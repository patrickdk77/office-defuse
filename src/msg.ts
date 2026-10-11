import { type CfbEntry, type CompoundFile, entryPath } from './cfb';
import { decideContained, presence, pushContained, spillOf } from './contained';
import { clip, type Reported, RunContext } from './context';
import { StopError, stopsTree } from './errors';
import { bufferSource, chunksOf, SpillSink } from './io';
import { latin1, utf16 } from './oleobj';
import type { CompoundPlan } from './plan';
import { writeRtf } from './rewrite-rtf';
import { inspectRtf } from './rtf';
import { decompressRtf, storeRtf } from './rtfcomp';
import { OfficeCategory as C, OfficeDetail as D } from './types';

/** Property ids of an attachment (MS-OXPROPS). */
const P = { method: 0x3705, longName: 0x3707, name: 0x3704, display: 0x3001, mime: 0x370e, data: 0x3701, longPath: 0x370d, path: 0x3708 };
/** How an attachment holds its file (MS-OXCMSG 2.2.2.9): by value, by reference, embedded message, OLE object, web reference. */
const METHODS: Record<number, string> = { 1: 'by value', 2: 'by reference', 3: 'by reference, resolved', 4: 'by reference only', 5: 'embedded message', 6: 'OLE object', 7: 'by web reference' };
/** Embedded messages can hold messages in turn. Past this depth they are still reported, and not read further. */
const MAX_DEPTH = 16;
/** A message's body properties (MS-OXPROPS): PidTagRtfCompressed, and PidTagHtml as binary, Unicode or 8-bit text. */
const RTF_BODY = 0x10090102;
const HTML_BODIES = [0x10130102, 0x1013001f, 0x1013001e];
/** The stream that holds a variable-size property (MS-OXMSG 2.1.4.1). */
const streamName = (tag: number) => `__SUBSTG1.0_${tag.toString(16).toUpperCase().padStart(8, '0')}`;

/** What the rewrite changes in a message's property stream: properties that go, and sizes that change. */
interface PropertyEdit {
  attachmentsDropped: number;
  drop: Set<number>;
  sizes: Map<number, number>;
}

/**
 * Inspects an Outlook message (MS-OXMSG): each attachment is reported with its name and type, and an attachment by
 * reference with what it points to. An embedded message's own attachments are read as well.
 */
export async function messageRules(ctx: RunContext, cfb: CompoundFile, message: CfbEntry, depth = 0, plan?: CompoundPlan): Promise<void> {
  const edit: PropertyEdit = { attachmentsDropped: 0, drop: new Set(), sizes: new Map() };
  await bodyRules(ctx, cfb, message, edit, plan);
  for (const a of message.children) {
    if (a.kind !== 'storage' || !/^__attach_version1\.0_#[0-9A-F]{8}$/i.test(a.name)) continue;
    // An attachment a strip finding names goes with its whole storage.
    if ((await attachmentRules(ctx, cfb, a, depth, plan)).some(r => r.action === 'strip') && plan) {
      plan.drop.add(a);
      edit.attachmentsDropped++;
    }
  }
  if (plan && (edit.attachmentsDropped || edit.drop.size || edit.sizes.size)) await editProperties(ctx, cfb, message, edit, plan);
}

/**
 * The property stream of a message after the rewrite (MS-OXMSG 2.4). Its header counts the attachments, at the same
 * place in a message and in an embedded one, and the count drops by the attachments that went. Each property held in
 * a stream has an entry giving that stream's size: the entry goes with a body that went, and takes the new size of a
 * body that was rewritten.
 */
async function editProperties(ctx: RunContext, cfb: CompoundFile, message: CfbEntry, edit: PropertyEdit, plan: CompoundPlan): Promise<void> {
  const s = message.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === '__PROPERTIES_VERSION1.0');
  if (!s || s.size < 24 || s.size > ctx.settings.decompressedBytes) return;
  const b = Buffer.from(await cfb.bytes(s));
  const headerSize = message === cfb.root ? 32 : 24;
  if (edit.attachmentsDropped) b.writeUInt32LE(Math.max(0, b.readUInt32LE(20) - edit.attachmentsDropped), 20);
  const kept: Buffer[] = [b.subarray(0, headerSize)];
  for (let at = headerSize; at + 16 <= b.length; at += 16) {
    const entry = Buffer.from(b.subarray(at, at + 16));
    const tag = entry.readUInt32LE(0);
    if (edit.drop.has(tag)) continue;
    const size = edit.sizes.get(tag);
    if (size !== undefined) entry.writeUInt32LE(size, 8);
    kept.push(entry);
  }
  plan.replace.set(s, await spillOf(ctx.run, Buffer.concat(kept)));
}

/** A stream of the message storage by property tag. */
const bodyStream = (message: CfbEntry, tag: number) => message.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === streamName(tag));

/** Every byte a sink holds. */
async function bytesOf(sink: SpillSink): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const c of sink.read()) chunks.push(Buffer.from(c));
  return Buffer.concat(chunks);
}

async function* once(b: Uint8Array): AsyncGenerator<Uint8Array> {
  yield b;
}

/**
 * A message's bodies. The RTF body goes to the plugins as the RTF it decompresses to, and a plugin's copy is stored
 * back uncompressed. When no plugin takes it, the RTF rules read it: a strip finding inside rewrites it, and a finding
 * that rejects it, or a body that does not decompress, makes it go, leaving the plain text body. The HTML body goes
 * to the plugins, and when none takes it, it stays, as no rule names it.
 */
async function bodyRules(ctx: RunContext, cfb: CompoundFile, message: CfbEntry, edit: PropertyEdit, plan: CompoundPlan | undefined): Promise<void> {
  const replace = async (stream: CfbEntry, tag: number, bytes: Uint8Array) => {
    plan?.replace.set(stream, await spillOf(ctx.run, bytes));
    edit.sizes.set(tag, bytes.length);
  };
  const drop = (stream: CfbEntry, tag: number) => {
    plan?.drop.add(stream);
    edit.drop.add(tag);
  };
  const rtf = bodyStream(message, RTF_BODY);
  if (rtf) {
    const location = entryPath(rtf);
    const body = rtf.size <= ctx.settings.decompressedBytes ? decompressRtf(await cfb.bytes(rtf), ctx.settings.decompressedBytes) : undefined;
    if (!body) {
      if (ctx.report(C.Package, D.UnreadableBody, location, { stream: location }).action === 'strip') drop(rtf, RTF_BODY);
    } else {
      const decision = await decideContained(ctx, { name: 'RTF body', declaredType: 'application/rtf', location, chunks: () => once(body) });
      if (decision) {
        if (!decision.keep) drop(rtf, RTF_BODY);
        else if (decision.replacement) await replace(rtf, RTF_BODY, storeRtf(await bytesOf(decision.replacement)));
      } else await rtfRules(ctx, body, location, rtf, edit, plan);
    }
  }
  for (const tag of HTML_BODIES) {
    const html = bodyStream(message, tag);
    if (!html || html.size > ctx.settings.decompressedBytes) continue;
    const raw = await cfb.bytes(html);
    const unicode = (tag & 0xffff) === 0x001f;
    const text = unicode ? Buffer.from(utf16(raw).replace(/\0+$/, ''), 'utf8') : raw;
    const decision = await decideContained(ctx, { name: 'HTML body', declaredType: 'text/html', location: entryPath(html), chunks: () => once(text) });
    if (decision && !decision.keep) drop(html, tag);
    else if (decision?.replacement) {
      const out = await bytesOf(decision.replacement);
      await replace(html, tag, unicode ? Buffer.from(`${out.toString('utf8')}\0`, 'utf16le') : out);
    }
    break;
  }
}

/**
 * The RTF rules over a message's RTF body, reporting into this run with the body's place. A strip finding rewrites the
 * body, and a reject drops it.
 */
async function rtfRules(ctx: RunContext, body: Uint8Array, location: string, stream: CfbEntry, edit: PropertyEdit, plan: CompoundPlan | undefined): Promise<void> {
  const sub = new RunContext(ctx.factory, ctx.xmlLimits, ctx.settings, ctx.run);
  // The RTF reader leaves its cuts where a whole file's rewrite finds them, which is the message's own.
  const saved = ctx.run.output;
  let cuts: Extract<NonNullable<typeof saved>, { kind: 'rtf' }> | undefined;
  try {
    await inspectRtf(sub, bufferSource(body), body.length);
    const out = ctx.run.output;
    if (out?.kind === 'rtf') cuts = out;
  } catch (e) {
    if (stopsTree(e) || !(e instanceof StopError)) throw e;
    sub.report(e.category, e.detail, undefined, { ...e.data, reason: e.data?.reason ?? e.message });
    cuts = undefined;
  } finally {
    ctx.run.output = saved;
  }
  const found = sub.findings;
  pushContained(ctx, found, location, 'RTF body');
  if (found.some(f => f.action === 'reject') || !cuts) {
    if (found.some(f => f.action !== 'info')) {
      plan?.drop.add(stream);
      edit.drop.add(RTF_BODY);
    }
    return;
  }
  if (!found.some(f => f.action === 'strip') || !plan) return;
  const sink = new SpillSink(ctx.run.temp, ctx.run.threshold);
  ctx.run.held.push(sink);
  await writeRtf(ctx.run, bufferSource(body), body.length, cuts.cuts, sink);
  await sink.close();
  const stored = storeRtf(await bytesOf(sink));
  plan.replace.set(stream, await spillOf(ctx.run, stored));
  edit.sizes.set(RTF_BODY, stored.length);
}

async function attachmentRules(ctx: RunContext, cfb: CompoundFile, a: CfbEntry, depth: number, plan: CompoundPlan | undefined): Promise<Reported[]> {
  const location = entryPath(a);
  const props = await properties(cfb, a);
  const text = async (id: number) => {
    for (const [type, decode] of [
      ['001F', utf16],
      ['001E', latin1],
    ] as const) {
      const s = a.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === `__SUBSTG1.0_${id.toString(16).toUpperCase().padStart(4, '0')}${type}`);
      if (s && s.size <= 65536) return decode(await cfb.bytes(s)).replace(/\0+$/, '');
    }
    return '';
  };
  const method = props.get(P.method) ?? 0;
  const name = (await text(P.longName)) || (await text(P.name)) || (await text(P.display));
  const declaredType = await text(P.mime);
  const dataId = `__SUBSTG1.0_${P.data.toString(16).toUpperCase()}`;
  const data = a.children.find(c => c.name.toUpperCase().startsWith(dataId));
  const info: Record<string, string | number> = { attachment: location, method: METHODS[method] ?? String(method) };
  if (name) info.name = clip(name);
  if (declaredType) info.type = clip(declaredType);
  if (method >= 2 && method <= 4) {
    const path = (await text(P.longPath)) || (await text(P.path));
    return path ? [ctx.external(D.Other, path, location, { source: 'attachment by reference', ...(name ? { name: clip(name) } : {}) }, 'attachment')] : [];
  }
  if (method === 7) {
    const path = (await text(P.longPath)) || (await text(P.path));
    return path ? [ctx.external(D.Other, path, location, { source: 'attachment by web reference', ...(name ? { name: clip(name) } : {}) }, 'attachment')] : [];
  }
  if (!data) return [];
  if (data.kind === 'stream') {
    return [await attachment(ctx, cfb, data, { name: name || undefined, declaredType: declaredType || undefined, location }, info, plan)];
  }
  // An embedded message or an OLE object is a storage of this file. The object rules read an object's storage, and
  // an embedded message's attachments are read here.
  info.size = 0;
  if (method === 5) info.type ??= 'application/vnd.ms-outlook';
  const r = ctx.report(C.Embedded, D.Attachment, location, info);
  if (method === 5 && depth < MAX_DEPTH) await messageRules(ctx, cfb, data, depth + 1, plan);
  return [r];
}

/**
 * An attached file held in a stream. It goes to the plugins, whose verdict decides whether it stays. When no plugin
 * takes it, EMBEDDED/ATTACHMENT does.
 */
async function attachment(
  ctx: RunContext,
  cfb: CompoundFile,
  data: CfbEntry,
  file: { name?: string; declaredType?: string; location: string },
  info: Record<string, string | number>,
  plan: CompoundPlan | undefined,
): Promise<Reported> {
  const source = cfb.source(data);
  const decision = await decideContained(ctx, { ...file, chunks: () => chunksOf(source, data.size) });
  if (decision?.replacement) plan?.replace.set(data, decision.replacement);
  const r = ctx.report(C.Embedded, D.Attachment, file.location, { ...info, size: data.size }, presence(decision));
  return decision && !decision.keep ? { ...r, action: 'strip' } : r;
}

/** The fixed-size values of a property stream (MS-OXMSG 2.4.2), by property id. An attachment's has an 8-byte header. */
async function properties(cfb: CompoundFile, storage: CfbEntry): Promise<Map<number, number>> {
  const out = new Map<number, number>();
  const s = storage.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === '__PROPERTIES_VERSION1.0');
  if (!s || s.size > 65536) return out;
  const b = await cfb.bytes(s);
  for (let at = 8; at + 16 <= b.length; at += 16) {
    const tag = (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
    // PT_LONG values sit in the first four bytes of the value.
    if ((tag & 0xffff) === 0x0003) out.set(tag >>> 16, (b[at + 8] | (b[at + 9] << 8) | (b[at + 10] << 16) | (b[at + 11] << 24)) >>> 0);
  }
  return out;
}
