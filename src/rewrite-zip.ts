import * as stream from 'node:stream';
import * as zlib from 'node:zlib';
import type { RunState } from './context';
import { crc32 } from './crc32';
import { checkTime } from './errors';
import type { Format } from './formats';
import { SpillSink } from './io';
import { isAbsoluteTarget, isInsidePackage } from './links';
import { folderOf, hrefOf, NS as ODF, odfNamespace, packagePath } from './odf';
import { CONTENT_TYPES_NS, isXmlType, type OoxmlGraph, type Part, R_NS, RELS_NS, resolvePartRef, resolveTarget, SHEET_NS, TYPE, WORD_NS, XPS_NS, xpsResources } from './ooxml';
import type { PackagePlan } from './plan';
import { type Edit, partEdits, type RewriteRules, splice } from './rewrite-xml';
import type { ByteSink } from './types';
import { attr, type XmlElement, type XmlLimits } from './xml';
import { nameKey, type ZipArchive, type ZipEntry } from './zip';
import { type CopiedEntry, ZipWriter } from './zip-writer';

const VML_OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const MACRO_RELATIONSHIPS = new Set(['http://schemas.microsoft.com/office/2006/relationships/xlmacrosheet', 'http://schemas.microsoft.com/office/2006/relationships/xlintlmacrosheet']);
/**
 * SpreadsheetML elements whose schema wants at least one child, which go once their last child went. An external
 * link part whose link went goes whole.
 */
const LISTS = new Set(['oleObjects', 'controls', 'hyperlinks', 'externalReferences', 'externalLink']);

/** An entry ready to write: its local header's fields, and compressed bytes that can be read more than once. */
interface Ready {
  info: CopiedEntry;
  data(): AsyncIterable<Uint8Array>;
  dispose(): Promise<void>;
  /** The entry deflated again, for a copy whose own bytes would put a PDF header where PDF readers look. */
  again?(): Promise<Ready>;
}

async function* bytesOf(b: Uint8Array): AsyncGenerator<Uint8Array> {
  yield b;
}

const nothing = async () => undefined;

/** Deflates what `chunks` yields. The compressed bytes wait in a spill sink for their size and CRC. */
async function deflatedEntry(run: RunState, e: Pick<ZipEntry, 'name' | 'dosTime' | 'dosDate'>, chunks: AsyncIterable<Uint8Array>): Promise<Ready> {
  let crc = 0;
  let size = 0;
  const counted = (async function* () {
    for await (const c of chunks) {
      checkTime(run.deadline);
      crc = crc32(c, crc);
      size += c.length;
      yield c;
    }
  })();
  const deflated = new SpillSink(run.temp, run.threshold);
  try {
    const src = stream.Readable.from(counted);
    const deflater = zlib.createDeflateRaw();
    src.on('error', err => deflater.destroy(err));
    src.pipe(deflater);
    for await (const c of deflater as AsyncIterable<Buffer>) await deflated.write(c);
    await deflated.close();
  } catch (err) {
    await deflated.dispose();
    throw err;
  }
  const info = { name: e.name, crc, size, dosTime: e.dosTime, dosDate: e.dosDate };
  if (size === 0) {
    await deflated.dispose();
    return { info: { ...info, method: 0, compressedSize: 0 }, data: () => bytesOf(new Uint8Array(0)), dispose: nothing };
  }
  return { info: { ...info, method: 8, compressedSize: deflated.length }, data: () => deflated.read(), dispose: () => deflated.dispose() };
}

/** An entry whose compressed bytes are copied as they are. The run inflated them once and checked them against their CRC. */
function copiedEntry(run: RunState, zip: ZipArchive, e: ZipEntry): Ready {
  return {
    info: { name: e.name, method: e.method, crc: e.crc, compressedSize: e.compressedSize, size: e.size, dosTime: e.dosTime, dosDate: e.dosDate },
    data: () => zip.raw(e),
    dispose: nothing,
    again: () => deflatedEntry(run, e, zip.read(e)),
  };
}

/** An entry made from bytes in memory, deflated, with the time ZipWriter gives new entries. */
function bytesEntry(name: string, data: Uint8Array): Ready {
  const body = zlib.deflateRawSync(data);
  return { info: { name, method: 8, crc: crc32(data), compressedSize: body.length, size: data.length }, data: () => bytesOf(body), dispose: nothing };
}

/** PDF readers take a header that starts in the first 1024 bytes of a file. */
const PDF_WINDOW = 1024;
const PDF_HEADER = Buffer.from('%PDF-', 'latin1');
/** The bytes a header that starts in the window reaches. */
const PDF_SPAN = PDF_WINDOW + PDF_HEADER.length - 1;
/** The most entries that wait, so a package of many such entries cannot keep them all. */
const MAX_WAITING = 16;

/**
 * Writes entries so that no PDF header starts in the first 1024 bytes of the archive. Entries that go, and entries
 * rewritten to another size, move what follows, so a stored PDF, a name or a local header from deep in the upload
 * can land where PDF readers look. A copied entry that would put a PDF header there is deflated again. Any other
 * entry that would waits until later entries have filled the window, up to 16 of them. What still lands there fails
 * the output check, which rejects the file.
 */
class PdfWindow {
  readonly writer: ZipWriter;
  private head = Buffer.alloc(0);
  private readonly waiting: Ready[] = [];

  constructor(sink: ByteSink) {
    this.writer = new ZipWriter({
      write: async b => {
        if (this.head.length < PDF_SPAN) this.head = Buffer.concat([this.head, b.subarray(0, PDF_SPAN - this.head.length)]);
        await sink.write(b);
      },
      close: () => sink.close(),
    });
  }

  /** True when writing `r` next keeps a PDF header out of the window. */
  private async clear(r: Ready): Promise<boolean> {
    if (this.head.length >= PDF_SPAN) return true;
    const parts = [this.head, Buffer.from(this.writer.localHeader(r.info))];
    let got = parts[0].length + parts[1].length;
    if (got < PDF_SPAN) {
      for await (const c of r.data()) {
        parts.push(Buffer.from(c));
        got += c.length;
        if (got >= PDF_SPAN) break;
      }
    }
    return !Buffer.concat(parts).subarray(0, PDF_SPAN).includes(PDF_HEADER);
  }

  private async write(r: Ready): Promise<void> {
    try {
      await this.writer.copy(r.info, r.data());
    } finally {
      await r.dispose();
    }
  }

  async add(r: Ready): Promise<void> {
    let ready = r;
    if (!(await this.clear(ready)) && ready.again) {
      const other = await ready.again();
      await ready.dispose();
      ready = other;
    }
    if (this.waiting.length < MAX_WAITING && !(await this.clear(ready))) {
      this.waiting.push(ready);
      return;
    }
    await this.write(ready);
    while (this.waiting.length && this.head.length >= PDF_SPAN) await this.write(this.waiting.shift() as Ready);
  }

  /** Writes the entries still waiting, each time the first that keeps the window clear, or the first when none does. */
  async end(): Promise<void> {
    while (this.waiting.length) {
      let i = 0;
      while (i < this.waiting.length && !(await this.clear(this.waiting[i]))) i++;
      const [r] = this.waiting.splice(i < this.waiting.length ? i : 0, 1);
      await this.write(r);
    }
  }

  /** Frees what the waiting entries hold, after a failure. */
  async dispose(): Promise<void> {
    for (const r of this.waiting.splice(0)) await r.dispose();
  }
}

/** The parts the relationships reach from the package, past parts that go and out of no macro sheet that is emptied. */
function reach(graph: OoxmlGraph, drops: Set<string>, emptied: Set<string>): Set<Part> {
  const key = (p: Part) => nameKey(p.entry.name);
  const reached = new Set<Part>();
  const queue: Part[] = [];
  for (const r of graph.roots) {
    if (drops.has(key(r)) || reached.has(r)) continue;
    reached.add(r);
    queue.push(r);
  }
  for (let p = queue.pop(); p; p = queue.pop()) {
    if (emptied.has(key(p))) continue;
    for (const t of p.targets) {
      if (drops.has(key(t)) || reached.has(t)) continue;
      reached.add(t);
      queue.push(t);
    }
  }
  return reached;
}

const escapeAttribute = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/**
 * Writes an OOXML package without what the plan removes. A part goes with its own relationships, every relationship
 * that targets it, its content type override, every element that names one of those relationships, and the parts
 * only it reached. A macro sheet becomes an empty worksheet. Parts nothing changes keep their compressed bytes.
 */
async function writeOoxml(run: RunState, zip: ZipArchive, plan: PackagePlan, graph: OoxmlGraph, format: Format, limits: XmlLimits, sink: ByteSink): Promise<void> {
  const drops = new Set(plan.drop);
  const sheets = plan.sheets;
  const before = reach(graph, new Set(), new Set());
  const relsEdits = new Map<string, Edit[]>();
  const partEditsByName = new Map<string, Edit[]>();
  const strict = graph.strict;
  const worksheetRelationship = strict ? 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  const binarySheet = (k: string) => !isXmlType(graph.parts.get(k)?.contentType) && format.format === 'xlsb';
  const sheetType = (k: string) => (binarySheet(k) ? 'application/vnd.ms-excel.worksheet' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  // Removing an element can empty a part of its root, and that part then goes too, so this runs until nothing more goes.
  for (;;) {
    const after = reach(graph, drops, sheets);
    for (const p of before) if (!after.has(p)) drops.add(nameKey(p.entry.name));
    for (const [source, rels] of graph.relsOf) if (source !== '' && (drops.has(source) || sheets.has(source))) drops.add(nameKey(rels.name));
    const removedIds = new Map<string, Set<string>>();
    relsEdits.clear();
    for (const [source, rels] of graph.relsOf) {
      if (drops.has(nameKey(rels.name))) continue;
      const ids = plan.rels.get(rels.name) ?? new Set<string>();
      const removed = new Set<string>();
      const rules: RewriteRules = {
        element(el) {
          if (el.depth !== 2 || el.local !== 'Relationship' || el.ns.toLowerCase() !== RELS_NS) return undefined;
          const id = attr(el, '', 'Id') ?? '';
          const target = attr(el, '', 'Target');
          if (ids.has(id)) {
            removed.add(id);
            return { remove: true };
          }
          if (target === undefined) return undefined;
          const external = (attr(el, '', 'TargetMode') ?? 'Internal').trim().toLowerCase() !== 'internal' || isAbsoluteTarget(target);
          if (external) return undefined;
          const k = resolveTarget(rels.name, target);
          if (drops.has(k)) {
            removed.add(id);
            return { remove: true };
          }
          const type = (attr(el, '', 'Type') ?? '').trim().toLowerCase();
          if (sheets.has(k) && MACRO_RELATIONSHIPS.has(type)) return { set: new Map([['Type', worksheetRelationship]]) };
          return undefined;
        },
      };
      const r = await partEdits(zip.read(rels), rels.name, limits, false, plan.parts.get(rels.name), rules);
      if (r.edits.length) relsEdits.set(rels.name, r.edits);
      if (removed.size) removedIds.set(source, removed);
    }
    let more = false;
    partEditsByName.clear();
    // An XPS part names the parts it uses in its markup, so every XPS part is read again once a part goes.
    const xpsDrops = graph.xps && drops.size > 0;
    for (const [k, p] of graph.parts) {
      const removed = removedIds.get(k);
      const own = plan.parts.get(p.entry.name);
      // An empty part, such as the presentation part PowerPoint writes into an add-in, refers to nothing.
      if (drops.has(k) || sheets.has(k) || plan.replace.has(k) || !isXmlType(p.contentType) || p.entry.size === 0 || (!removed && !own && !xpsDrops)) continue;
      const rules: RewriteRules = {
        element(el) {
          // An XPS element that names a part that went goes with it.
          if (xpsDrops && XPS_NS.has(el.ns.toLowerCase())) {
            for (const [, uri] of xpsResources(el)) if (!isAbsoluteTarget(uri) && drops.has(resolvePartRef(p.entry.name, uri))) return { remove: true };
          }
          if (!removed) return undefined;
          for (const a of el.attributes) {
            const ns = a.ns.toLowerCase();
            if ((R_NS.has(ns) || (ns === VML_OFFICE_NS && a.local === 'relid')) && removed.has(a.value)) {
              // A hyperlink's text stays. Any other element that names a relationship that went goes with it.
              return WORD_NS.has(el.ns.toLowerCase()) && el.local === 'hyperlink' ? { unwrap: true } : { remove: true };
            }
          }
          return undefined;
        },
        // An XPS property element such as Path.Fill holds one value, and goes once that value went.
        collapses: el => (SHEET_NS.has(el.ns.toLowerCase()) && LISTS.has(el.local)) || (graph.xps && XPS_NS.has(el.ns.toLowerCase()) && el.local.includes('.')),
        word: el => WORD_NS.has(el.ns.toLowerCase()),
      };
      const r = await partEdits(zip.read(p.entry), p.entry.name, limits, p.contentType === TYPE.vml, own, rules);
      if (r.rootRemoved) {
        drops.add(k);
        more = true;
      } else if (r.edits.length) partEditsByName.set(p.entry.name, r.edits);
    }
    if (!more) break;
  }

  // The content types of parts that went, and of macro sheets that became worksheets.
  const overridden = new Set<string>();
  let prefix = '';
  const typesRules: RewriteRules = {
    element(el) {
      if (el.ns.toLowerCase() !== CONTENT_TYPES_NS) return undefined;
      if (el.depth === 1) prefix = el.qname.includes(':') ? el.qname.slice(0, el.qname.indexOf(':') + 1) : '';
      if (el.local !== 'Override') return undefined;
      const name = attr(el, '', 'PartName');
      if (name === undefined) return undefined;
      const k = nameKey(name.replace(/^\//, ''));
      if (drops.has(k)) return { remove: true };
      if (!sheets.has(k)) return undefined;
      overridden.add(k);
      return { set: new Map([['ContentType', sheetType(k)]]) };
    },
    append: () =>
      [...sheets]
        .filter(k => !overridden.has(k) && !drops.has(k))
        .map(k => `<${prefix}Override PartName="/${escapeAttribute(graph.parts.get(k)?.entry.name ?? k)}" ContentType="${sheetType(k)}"/>`)
        .join(''),
  };
  const types = await partEdits(zip.read(graph.contentTypes), graph.contentTypes.name, limits, false, plan.parts.get(graph.contentTypes.name), typesRules);

  const window = new PdfWindow(sink);
  try {
    for (const e of zip.entries) {
      checkTime(run.deadline);
      const k = nameKey(e.name);
      if (drops.has(k)) continue;
      const replaced = plan.replace.get(k);
      const edits = e === graph.contentTypes ? types.edits : (relsEdits.get(e.name) ?? partEditsByName.get(e.name));
      if (replaced) await window.add(await deflatedEntry(run, e, replaced.read()));
      else if (sheets.has(k)) await window.add(await deflatedEntry(run, e, bytesOf(emptySheet(binarySheet(k), strict))));
      else if (edits?.length) await window.add(await deflatedEntry(run, e, splice(zip.read(e), edits)));
      else await window.add(copiedEntry(run, zip, e));
    }
    await window.end();
  } finally {
    await window.dispose();
  }
  await window.writer.finish();
}

/** An empty worksheet, as XML or as the records of a binary workbook. */
function emptySheet(binary: boolean, strict: boolean): Uint8Array {
  // BrtBeginSheet, BrtBeginSheetData, BrtEndSheetData and BrtEndSheet, each with no data.
  if (binary) return Uint8Array.of(0x81, 0x01, 0x00, 0x91, 0x01, 0x00, 0x92, 0x01, 0x00, 0x82, 0x01, 0x00);
  const ns = strict ? 'http://purl.oclc.org/ooxml/spreadsheetml/main' : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  return Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${ns}"><sheetData/></worksheet>`, 'utf8');
}

/** A relative path resolved against a folder of the package, as a key. "." and ".." segments are read. */
function resolveIn(folder: string, href: string): string {
  const out: string[] = [];
  for (const seg of `${folder}${href.trim()}`.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return nameKey(out.join('/'));
}

/**
 * Writes an OpenDocument package without what the plan removes, with the mimetype entry first, stored and without an
 * extra field. An entry goes with its manifest entry and every element whose link names it. A package without a
 * manifest gets one that lists its entries.
 */
async function writeOdf(run: RunState, zip: ZipArchive, plan: PackagePlan, format: Format, limits: XmlLimits, sink: ByteSink): Promise<void> {
  const layout = plan.odf;
  const drops = new Set(plan.drop);
  const folders = plan.folders;
  const gone = (k: string) => drops.has(k) || [...folders].some(f => k.startsWith(f) || `${k}/` === f);
  for (const e of zip.entries) if (gone(nameKey(e.name))) drops.add(nameKey(e.name));
  const xmlEntries = zip.entries.filter(e => !e.directory && e !== layout?.mimetype && e !== layout?.manifest && /\.xml$/i.test(e.name) && !folderOf(e.name));
  const edits = new Map<string, Edit[]>();
  for (;;) {
    let more = false;
    edits.clear();
    for (const e of xmlEntries) {
      const k = nameKey(e.name);
      const own = plan.parts.get(e.name);
      if (drops.has(k) || plan.replace.has(k) || (!own && !drops.size)) continue;
      const folder = e.name.includes('/') ? e.name.slice(0, e.name.lastIndexOf('/') + 1) : '';
      const rules: RewriteRules = {
        element(el: XmlElement) {
          const written = hrefOf(el);
          // LibreOffice's vnd.sun.star.Package: URLs name an entry too.
          const href = written === undefined ? undefined : (/^\s*vnd\.sun\.star\.package:(.*)$/is.exec(written)?.[1] ?? written);
          if (href === undefined || !isInsidePackage(href)) return undefined;
          // A link inside a sub-document may name its entry from the package root or from the sub-document's folder.
          return gone(packagePath(href)) || gone(resolveIn(folder, href)) ? { remove: true } : undefined;
        },
      };
      const r = await partEdits(zip.read(e), e.name, limits, false, own, rules);
      if (r.rootRemoved) {
        drops.add(k);
        more = true;
      } else if (r.edits.length) edits.set(e.name, r.edits);
    }
    if (!more) break;
  }
  const manifest = layout?.manifest;
  if (manifest && drops.size) {
    const rules: RewriteRules = {
      element(el) {
        if (odfNamespace(el.ns) !== ODF.manifest || el.local !== 'file-entry') return undefined;
        const path = el.attributes.find(a => a.local === 'full-path' && odfNamespace(a.ns) === ODF.manifest)?.value;
        return path !== undefined && path !== '/' && gone(packagePath(path)) ? { remove: true } : undefined;
      },
    };
    const r = await partEdits(zip.read(manifest), manifest.name, limits, false, plan.parts.get(manifest.name), rules);
    if (r.edits.length) edits.set(manifest.name, r.edits);
  } else if (manifest && plan.parts.has(manifest.name)) {
    const r = await partEdits(zip.read(manifest), manifest.name, limits, false, plan.parts.get(manifest.name), { element: () => undefined });
    if (r.edits.length) edits.set(manifest.name, r.edits);
  }

  const window = new PdfWindow(sink);
  const writer = window.writer;
  await writer.add('mimetype', Buffer.from(layout?.mediaType || format.mimeType, 'latin1'), { compress: false, dosTime: layout?.mimetype?.dosTime, dosDate: layout?.mimetype?.dosDate });
  const kept: ZipEntry[] = [];
  try {
    for (const e of zip.entries) {
      checkTime(run.deadline);
      const k = nameKey(e.name);
      if (e === layout?.mimetype || drops.has(k)) continue;
      kept.push(e);
      const replaced = plan.replace.get(k);
      const own = edits.get(e.name);
      if (replaced) await window.add(await deflatedEntry(run, e, replaced.read()));
      else if (own?.length) await window.add(await deflatedEntry(run, e, splice(zip.read(e), own)));
      else await window.add(copiedEntry(run, zip, e));
    }
    if (!manifest) {
      const entries = kept
        .filter(e => !e.directory)
        .map(e => `<manifest:file-entry manifest:full-path="${escapeAttribute(e.name)}" manifest:media-type="${/\.xml$/i.test(e.name) ? 'text/xml' : ''}"/>`)
        .join('');
      const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="${ODF.manifest}"><manifest:file-entry manifest:full-path="/" manifest:media-type="${escapeAttribute(layout?.mediaType || format.mimeType)}"/>${entries}</manifest:manifest>`;
      await window.add(bytesEntry('META-INF/manifest.xml', Buffer.from(xml, 'utf8')));
    }
    await window.end();
  } finally {
    await window.dispose();
  }
  await writer.finish();
}

/** Writes the rewritten package to `sink`, which the caller closes. */
export async function writePackage(run: RunState, zip: ZipArchive, plan: PackagePlan, format: Format, limits: XmlLimits, sink: ByteSink): Promise<void> {
  if (plan.graph) await writeOoxml(run, zip, plan, plan.graph, format, limits, sink);
  else await writeOdf(run, zip, plan, format, limits, sink);
}
