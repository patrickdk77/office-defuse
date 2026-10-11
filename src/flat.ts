import { TextDecoder } from 'node:util';
import { decideContained } from './contained';
import { clip, type Output, type PlanContext, type RunContext, type RunState } from './context';
import { checkTime, StopError } from './errors';
import type { Format } from './formats';
import { SpillSink } from './io';
import { EntryRules, NS as ODF, oattr, odfNamespace } from './odf';
import { CONTENT_TYPES_NS, inspectOoxml, isXmlType } from './ooxml';
import type { FlatOpcLayout, PackagePlan } from './plan';
import { partEdits, splice } from './rewrite-xml';
import { writePackage } from './rewrite-zip';
import { type ByteSink, type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';
import { attr, detectEncoding, type XmlElement, type XmlHandler, type XmlLimits, XmlTokenizer } from './xml';
import { nameKey, ZipArchive, type ZipLimits } from './zip';
import { ZipWriter } from './zip-writer';

/** The namespace of a Flat OPC package, the single-file form of OOXML that Word, Excel and PowerPoint save as XML. */
export const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
/** The name findings give the one part of a flat OpenDocument file. */
export const FLAT_ODF_PART = 'office:document';
/** The name findings give the wrapper of a Flat OPC package. */
const FLAT_OPC_PART = 'pkg:package';
const LOEXT_NS = 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0';

/** Flat OpenDocument formats, by the media type in office:mimetype. */
const FLAT_ODF = new Map<string, Format>(
  (
    [
      ['text', 'fodt'],
      ['spreadsheet', 'fods'],
      ['presentation', 'fodp'],
      ['graphics', 'fodg'],
    ] as const
  ).map(([kind, ext]) => [`application/vnd.oasis.opendocument.${kind}`, { format: ext, mimeType: `application/vnd.oasis.opendocument.${kind}-flat-xml` }]),
);

/** Flat OPC has no media type of its own, and Office saves it with the .xml extension. */
export const FLAT_OPC_FORMAT: Format = { format: 'xml', mimeType: 'application/xml' };

const CHUNK = 65536;

/** A source's bytes in chunks, checking the deadline before each. */
async function* chunksOf(src: ByteSource, size: number, deadline?: number): AsyncGenerator<Uint8Array> {
  for (let at = 0; at < size; at += CHUNK) {
    checkTime(deadline);
    yield await src.read(at, Math.min(CHUNK, size - at));
  }
}

/** Whether the first bytes of a file read as XML: a byte order mark, white space and then "<". */
export function startsAsXml(head: Uint8Array): boolean {
  const found = detectEncoding(head);
  if (!found) return false;
  const text = new TextDecoder(found.encoding, { fatal: false }).decode(head.subarray(found.bom));
  return text.trimStart().startsWith('<');
}

/**
 * Whether the start of an XML file names a flat OpenDocument document or a Flat OPC package as its root element. The
 * reading of the whole file decides; this tells a router which package to try.
 */
export function isFlatHead(text: string): boolean {
  const body = text.replace(/^\uFEFF/, '').replace(/^(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>)*/, '');
  const qname = /^<([^\s/>]+)/.exec(body)?.[1] ?? '';
  const local = qname.slice(qname.indexOf(':') + 1);
  return (local === 'document' && /urn:oasis:names:tc:opendocument:xmlns:office:1\.\d/i.test(text)) || (local === 'package' && text.includes(PKG_NS));
}

/** Decodes base64 that arrives in pieces of any size, white space and line ends included. */
class Base64 {
  private rest = '';
  push(t: string): Buffer {
    const s = this.rest + t.replace(/\s+/g, '');
    const n = s.length - (s.length % 4);
    this.rest = s.slice(n);
    return Buffer.from(s.slice(0, n), 'base64');
  }
  end(): Buffer {
    const r = this.rest;
    this.rest = '';
    return Buffer.from(r, 'base64');
  }
}

/** Decoded bytes the handlers gather while the tokenizer runs, which go to their sinks between chunks. */
class Pending {
  private readonly queue: Array<[SpillSink, Uint8Array]> = [];
  private readonly closing: SpillSink[] = [];
  add(sink: SpillSink, b: Uint8Array): void {
    if (b.length) this.queue.push([sink, b]);
  }
  close(sink: SpillSink): void {
    this.closing.push(sink);
  }
  async flush(): Promise<void> {
    for (const [sink, b] of this.queue.splice(0)) await sink.write(b);
    for (const sink of this.closing.splice(0)) await sink.close();
  }
}

/** A file held in base64 inside a flat file: the element that holds it and its office:binary-data element. */
interface Binary {
  parent?: XmlElement;
  data: XmlElement;
  sink: SpillSink;
}

/**
 * The rules of a flat OpenDocument file: the OpenDocument rules over its one part, and the files it holds in
 * office:binary-data, such as pictures and OLE objects, which go to the plugins.
 */
class FlatOdf implements XmlHandler {
  readonly rules: EntryRules;
  readonly binaries: Binary[] = [];
  private readonly stack: XmlElement[] = [];
  private current?: { binary: Binary; b64: Base64 };

  constructor(
    ctx: PlanContext,
    private readonly run: RunState,
    private readonly pending: Pending,
  ) {
    this.rules = new EntryRules(ctx, FLAT_ODF_PART, new Map(), true);
  }

  start(el: XmlElement): void {
    this.rules.start(el);
    if (!this.current && odfNamespace(el.ns) === ODF.office && el.local === 'binary-data') {
      const sink = new SpillSink(this.run.temp, this.run.threshold);
      this.run.held.push(sink);
      this.current = { binary: { parent: this.stack[this.stack.length - 1], data: el, sink }, b64: new Base64() };
    }
    this.stack.push(el);
  }

  text(t: string): void {
    if (this.current) this.pending.add(this.current.binary.sink, this.current.b64.push(t));
    else this.rules.text(t);
  }

  end(el: XmlElement): void {
    this.stack.pop();
    const c = this.current;
    if (c && c.binary.data.index === el.index) {
      this.pending.add(c.binary.sink, c.b64.end());
      this.pending.close(c.binary.sink);
      this.binaries.push(c.binary);
      this.current = undefined;
    }
    this.rules.end(el);
  }
}

/** One part of a Flat OPC package, as the first reading finds it. */
interface FlatPart {
  name: string;
  contentType: string;
  /** The part's attributes besides its name and content type, such as pkg:compression, kept for the output. */
  extra: Array<[string, string]>;
  /** Where the part's XML lies in the decoded text, and the namespaces in scope there that its root does not declare. */
  xml?: { start: number; end: number; inherited: Map<string, string> };
  binary?: SpillSink;
}

/** The parts of a Flat OPC package. XML parts are cut from the text in a second reading, by their offsets. */
class FlatOpc implements XmlHandler {
  readonly parts: FlatPart[] = [];
  /** Where the pkg:package start tag begins, which is where the prolog ends. */
  rootStart = 0;
  private part?: FlatPart;
  private xmlDepth = 0;
  private b64?: Base64;
  private readonly scopes: Array<Array<[string, string]>> = [];

  constructor(
    private readonly run: RunState,
    private readonly pending: Pending,
  ) {}

  start(el: XmlElement): void {
    this.scopes.push(el.namespaces);
    if (el.depth === 1) this.rootStart = el.start;
    if (this.xmlDepth) {
      const xml = this.part?.xml;
      // The part's root inherits the namespaces in scope, and the ones it declares itself it does not need.
      if (xml && el.depth === this.xmlDepth + 1) for (const [prefix] of el.namespaces) xml.inherited.delete(prefix);
      return;
    }
    if (el.ns !== PKG_NS) return;
    if (el.depth === 2 && el.local === 'part') {
      const name = el.attributes.find(a => a.ns === PKG_NS && a.local === 'name')?.value;
      if (name === undefined || name.replace(/^\//, '') === '') throw new StopError(C.Package, D.Unrecognized, 'A Flat OPC part has no name', { reason: 'a part without a name' });
      const contentType = el.attributes.find(a => a.ns === PKG_NS && a.local === 'contentType')?.value ?? '';
      const extra = el.attributes.filter(a => !(a.ns === PKG_NS && (a.local === 'name' || a.local === 'contentType'))).map((a): [string, string] => [a.qname, a.value]);
      this.part = { name: name.replace(/^\//, ''), contentType, extra };
    } else if (el.depth === 3 && this.part && el.local === 'xmlData') {
      const inherited = new Map<string, string>();
      for (const scope of this.scopes) for (const [prefix, uri] of scope) inherited.set(prefix, uri);
      this.part.xml = { start: el.end, end: el.end, inherited };
      this.xmlDepth = el.depth;
    } else if (el.depth === 3 && this.part && el.local === 'binaryData') {
      const sink = new SpillSink(this.run.temp, this.run.threshold);
      this.run.held.push(sink);
      this.part.binary = sink;
      this.b64 = new Base64();
    }
  }

  text(t: string): void {
    if (this.b64 && this.part?.binary) this.pending.add(this.part.binary, this.b64.push(t));
  }

  end(el: XmlElement, start: number): void {
    this.scopes.pop();
    if (this.xmlDepth === el.depth && this.part?.xml) {
      this.part.xml.end = start;
      this.xmlDepth = 0;
    } else if (this.b64 && el.depth === 3 && this.part?.binary) {
      this.pending.add(this.part.binary, this.b64.end());
      this.pending.close(this.part.binary);
      this.b64 = undefined;
    } else if (el.depth === 2 && this.part) {
      this.parts.push(this.part);
      this.part = undefined;
    }
  }
}

/** Reads the root element and hands everything to the rules of the flat format it names. */
class FlatReader implements XmlHandler {
  odf?: FlatOdf;
  opc?: FlatOpc;
  format?: Format;

  constructor(
    private readonly ctx: PlanContext,
    private readonly pending: Pending,
  ) {}

  start(el: XmlElement): void {
    if (el.depth === 1) {
      if (odfNamespace(el.ns) === ODF.office && el.local === 'document') {
        const type = oattr(el, ODF.office, 'mimetype') ?? '';
        this.format = FLAT_ODF.get(type);
        if (!this.format) throw new StopError(C.Package, D.Unrecognized, 'The flat OpenDocument media type is not one this package reads', { mediaType: clip(type) });
        this.ctx.format = this.format;
        this.odf = new FlatOdf(this.ctx, this.ctx.run, this.pending);
      } else if (el.ns === PKG_NS && el.local === 'package') {
        this.opc = new FlatOpc(this.ctx.run, this.pending);
      } else throw new StopError(C.Package, D.Unrecognized, 'The XML file is neither flat OpenDocument nor Flat OPC', { root: clip(el.qname), namespace: clip(el.ns) });
    }
    this.odf?.start(el);
    this.opc?.start(el);
  }

  text(t: string): void {
    this.odf?.text(t);
    this.opc?.text(t);
  }

  end(el: XmlElement, start: number): void {
    this.odf?.end(el);
    this.opc?.end(el, start);
  }
}

/** Base64 of everything a sink holds. */
async function base64Of(sink: SpillSink): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const c of sink.read()) chunks.push(Buffer.from(c));
  return Buffer.concat(chunks).toString('base64');
}

/**
 * The plugins' verdicts on the files a flat OpenDocument file holds. A file a plugin removed goes with the element
 * that holds it, and one a plugin rewrote goes back in as base64. An OLE object a plugin kept stays, and one no plugin
 * took gets the rule EMBEDDED/OLE_OBJECT gives it. A picture no plugin took stays, since no rule names it.
 */
async function decideBinaries(ctx: PlanContext, odf: FlatOdf): Promise<void> {
  const plan = ctx.plan;
  const decided = new Set<number>();
  let n = 0;
  for (const b of odf.binaries) {
    n++;
    const parent = b.parent;
    if (!parent) continue;
    const declaredType = oattr(parent, ODF.draw, 'mime-type') ?? parent.attributes.find(a => a.ns === LOEXT_NS && a.local === 'mime-type')?.value;
    const decision = await decideContained(ctx, { location: `${FLAT_ODF_PART} > ${parent.qname} ${n}`, declaredType, chunks: () => b.sink.read() });
    const object = plan.objects.find(o => o.index === parent.index);
    if (object) {
      decided.add(parent.index);
      if (decision?.keep) {
        if (object.finding) {
          object.finding.action = 'info';
          ctx.run.superseded.add(object.finding);
        }
      } else if (decision || object.strips) plan.element(FLAT_ODF_PART, parent.index, { remove: true });
    } else if (decision && !decision.keep) plan.element(FLAT_ODF_PART, parent.index, { remove: true });
    if (decision?.keep && decision.replacement) plan.element(FLAT_ODF_PART, b.data.index, { content: await base64Of(decision.replacement) });
  }
  // An object with no data of its own gets its rule alone.
  for (const o of plan.objects) if (!decided.has(o.index) && o.strips) plan.element(FLAT_ODF_PART, o.index, { remove: true });
}

/** The ranges of the decoded text that `wanted` names, cut in one reading in order, and the text before `prolog`. */
async function cutText(src: ByteSource, size: number, deadline: number | undefined, prolog: number, wanted: Array<{ start: number; end: number }>): Promise<{ prolog: string; texts: string[] }> {
  const order = wanted.map((w, i) => ({ ...w, i })).sort((a, b) => a.start - b.start);
  const texts = wanted.map(() => '');
  let head = '';
  let pos = 0;
  let decoder: TextDecoder | undefined;
  let k = 0;
  const take = (s: string) => {
    if (pos < prolog) head += s.slice(0, Math.max(0, prolog - pos));
    for (let j = k; j < order.length && order[j].start < pos + s.length; j++) {
      const w = order[j];
      const from = Math.max(w.start, pos) - pos;
      const to = Math.min(w.end, pos + s.length) - pos;
      if (to > from) texts[w.i] += s.slice(from, to);
    }
    pos += s.length;
    while (k < order.length && order[k].end <= pos) k++;
  };
  let first = true;
  for await (const c of chunksOf(src, size, deadline)) {
    if (first) {
      first = false;
      const found = detectEncoding(c) ?? { encoding: 'utf-8' as const, bom: 0 };
      decoder = new TextDecoder(found.encoding, { fatal: true, ignoreBOM: true });
      take(decoder.decode(c.subarray(found.bom), { stream: true }));
      continue;
    }
    take((decoder as TextDecoder).decode(c, { stream: true }));
  }
  if (decoder) take(decoder.decode());
  return { prolog: head, texts };
}

const escapeAttribute = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/** A part's XML, cut from a Flat OPC package, as a part of its own: with an XML declaration and the namespaces it inherited. */
function partXml(text: string, inherited: Map<string, string>): Buffer {
  let body = text.trim();
  if (inherited.size) {
    // The root's start tag is the first "<" that does not open a comment or a processing instruction.
    const m = /^(?:\s|<!--[\s\S]*?-->|<\?[\s\S]*?\?>)*<[^\s/>]+/.exec(body);
    if (m) {
      const decls = [...inherited].map(([prefix, uri]) => ` ${prefix ? `xmlns:${prefix}` : 'xmlns'}="${escapeAttribute(uri)}"`).join('');
      body = body.slice(0, m[0].length) + decls + body.slice(m[0].length);
    }
  }
  return Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`, 'utf8');
}

/**
 * Inspects a flat XML file: flat OpenDocument, read with the OpenDocument rules over its one part, or Flat OPC,
 * which is unpacked into a package of its parts and read with the OOXML rules. Returns what the rewrite needs.
 */
export async function inspectFlat(ctx: RunContext, src: ByteSource, size: number, zipLimits: ZipLimits): Promise<Output | undefined> {
  const run = ctx.run;
  const pctx = ctx.withPlan();
  const pending = new Pending();
  const reader = new FlatReader(pctx, pending);
  // A finding about the document type comes before the root element, so the first bytes pick the name it gives.
  const head = await src.read(0, Math.min(CHUNK, size));
  const found = detectEncoding(head) ?? { encoding: 'utf-8' as const, bom: 0 };
  const opcHead = new TextDecoder(found.encoding, { fatal: false }).decode(head.subarray(found.bom)).includes(PKG_NS);
  await pctx.parseXml(opcHead ? FLAT_OPC_PART : FLAT_ODF_PART, chunksOf(src, size, ctx.settings.deadline), reader, false, pctx.plan, () => pending.flush());
  if (reader.odf) {
    await decideBinaries(pctx, reader.odf);
    return { kind: 'flatodf', plan: pctx.plan };
  }
  const opc = reader.opc;
  if (!opc) return undefined;
  const xmlParts = opc.parts.filter(p => p.xml);
  const cut = await cutText(
    src,
    size,
    ctx.settings.deadline,
    opc.rootStart,
    xmlParts.map(p => p.xml ?? { start: 0, end: 0 }),
  );
  // The parts go into a package of their own, with a [Content_Types].xml that gives each its content type.
  const packed = new SpillSink(run.temp, run.threshold);
  run.held.push(packed);
  const writer = new ZipWriter(packed);
  const seen = new Set<string>();
  for (const p of opc.parts) {
    const k = nameKey(p.name);
    if (seen.has(k) || k === nameKey('[Content_Types].xml')) throw new StopError(C.Container, D.DuplicateName, `Two Flat OPC parts are named ${clip(p.name)}`, { name: clip(p.name) });
    seen.add(k);
  }
  const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CONTENT_TYPES_NS}">${opc.parts.map(p => `<Override PartName="/${escapeAttribute(p.name)}" ContentType="${escapeAttribute(p.contentType)}"/>`).join('')}</Types>`;
  await writer.add('[Content_Types].xml', Buffer.from(types, 'utf8'), { compress: false });
  for (const p of opc.parts) {
    checkTime(ctx.settings.deadline);
    let data: Uint8Array = new Uint8Array(0);
    if (p.xml) data = partXml(cut.texts[xmlParts.indexOf(p)], p.xml.inherited);
    else if (p.binary) {
      const chunks: Buffer[] = [];
      for await (const c of p.binary.read()) chunks.push(Buffer.from(c));
      data = Buffer.concat(chunks);
    }
    await writer.add(p.name, data, { compress: false });
  }
  await writer.finish();
  await packed.close();
  const packedSource = packed.source();
  run.sources.push(packedSource);
  const zip = await ZipArchive.open(packedSource, zipLimits, ctx.settings.deadline);
  const zctx = ctx.withZip(zip);
  const contentTypes = zip.get('[Content_Types].xml');
  if (!contentTypes) return undefined;
  await inspectOoxml(zctx, contentTypes);
  const inner = ctx.format;
  ctx.format = FLAT_OPC_FORMAT;
  if (!inner) return undefined;
  const layout: FlatOpcLayout = {
    instructions: [...cut.prolog.matchAll(/<\?mso-application\s[^?]*\?>/g)].map(m => m[0]),
    extra: new Map(opc.parts.map(p => [nameKey(p.name), p.extra])),
  };
  return { kind: 'flatopc', zip, plan: zctx.plan, format: inner, layout };
}

/** Writes a flat OpenDocument file without what the plan removes. */
export async function writeFlatOdf(run: RunState, src: ByteSource, plan: PackagePlan, limits: XmlLimits, sink: ByteSink): Promise<void> {
  const size = await src.size();
  const r = await partEdits(chunksOf(src, size, run.deadline), FLAT_ODF_PART, limits, false, plan.parts.get(FLAT_ODF_PART), { element: () => undefined });
  for await (const c of splice(chunksOf(src, size, run.deadline), r.edits)) await sink.write(c);
}

/** The content types of a package's parts, by key, from its [Content_Types].xml. */
async function contentTypesOf(zip: ZipArchive, limits: XmlLimits): Promise<(name: string) => string> {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  const entry = zip.get('[Content_Types].xml');
  if (entry) {
    const t = new XmlTokenizer(
      {
        start(el) {
          if (el.ns.toLowerCase() !== CONTENT_TYPES_NS) return;
          const type = attr(el, '', 'ContentType') ?? '';
          if (el.local === 'Default') defaults.set((attr(el, '', 'Extension') ?? '').toLowerCase(), type);
          else if (el.local === 'Override') overrides.set(nameKey((attr(el, '', 'PartName') ?? '').replace(/^\//, '')), type);
        },
      },
      limits,
      entry.name,
    );
    for await (const c of zip.read(entry)) t.write(c);
    t.end();
  }
  return name => overrides.get(nameKey(name)) ?? defaults.get(/\.([^./]*)$/.exec(name)?.[1]?.toLowerCase() ?? '') ?? '';
}

/** A part's XML as it goes inside pkg:xmlData: decoded, without its byte order mark and XML declaration. */
function embeddedXml(bytes: Uint8Array): string {
  const found = detectEncoding(bytes) ?? { encoding: 'utf-8' as const, bom: 0 };
  const text = new TextDecoder(found.encoding, { fatal: true }).decode(bytes.subarray(found.bom));
  return text.replace(/^\s*<\?xml[\s\S]*?\?>\s*/, '');
}

/**
 * Writes a Flat OPC package without what the plan removes: the package is rewritten as any OOXML package is, and its
 * parts are written back into one XML file, each with its name, content type and attributes.
 */
export async function writeFlatOpc(run: RunState, output: Extract<Output, { kind: 'flatopc' }>, limits: XmlLimits, zipLimits: ZipLimits, sink: ByteSink): Promise<void> {
  const packed = new SpillSink(run.temp, run.threshold);
  run.held.push(packed);
  await writePackage(run, output.zip, output.plan, output.format, limits, packed);
  await packed.close();
  const packedSource = packed.source();
  run.sources.push(packedSource);
  const zip = await ZipArchive.open(packedSource, zipLimits, run.deadline);
  const typeOf = await contentTypesOf(zip, limits);
  const put = (s: string) => sink.write(Buffer.from(s, 'utf8'));
  await put(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${output.layout.instructions.map(i => `${i}\n`).join('')}<pkg:package xmlns:pkg="${PKG_NS}">`);
  for (const e of zip.entries) {
    checkTime(run.deadline);
    if (e.directory || e.name === '[Content_Types].xml') continue;
    const chunks: Buffer[] = [];
    for await (const c of zip.read(e)) chunks.push(Buffer.from(c));
    const bytes = Buffer.concat(chunks);
    const type = typeOf(e.name);
    const extra = (output.layout.extra.get(nameKey(e.name)) ?? []).map(([q, v]) => ` ${q}="${escapeAttribute(v)}"`).join('');
    await put(`<pkg:part pkg:name="/${escapeAttribute(e.name)}" pkg:contentType="${escapeAttribute(type)}"${extra}>`);
    if (isXmlType(type.toLowerCase()) || type.toLowerCase().endsWith('xml')) await put(`<pkg:xmlData>${embeddedXml(bytes)}</pkg:xmlData>`);
    else await put(`<pkg:binaryData>${(bytes.toString('base64').match(/.{1,76}/g) ?? []).join('\n')}</pkg:binaryData>`);
    await put('</pkg:part>');
  }
  await put('</pkg:package>');
}
