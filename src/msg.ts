import { type CfbEntry, type CompoundFile, entryPath } from './cfb';
import { clip, type RunContext } from './context';
import { latin1, utf16 } from './oleobj';
import { OfficeCategory as C, type ContainedFile, OfficeDetail as D } from './types';

/** Property ids of an attachment (MS-OXPROPS). */
const P = { method: 0x3705, longName: 0x3707, name: 0x3704, display: 0x3001, mime: 0x370e, data: 0x3701, longPath: 0x370d, path: 0x3708 };
/** How an attachment holds its file (MS-OXCMSG 2.2.2.9): by value, by reference, embedded message, OLE object, web reference. */
const METHODS: Record<number, string> = { 1: 'by value', 2: 'by reference', 3: 'by reference, resolved', 4: 'by reference only', 5: 'embedded message', 6: 'OLE object', 7: 'by web reference' };
/** Embedded messages can hold messages in turn. Past this depth they are still reported, and not read further. */
const MAX_DEPTH = 16;

/**
 * Inspects an Outlook message (MS-OXMSG): each attachment is reported with its name and type, and an attachment by
 * reference with what it points to. An embedded message's own attachments are read as well.
 */
export async function messageRules(ctx: RunContext, cfb: CompoundFile, message: CfbEntry, depth = 0): Promise<void> {
  for (const a of message.children) {
    if (a.kind !== 'storage' || !/^__attach_version1\.0_#[0-9A-F]{8}$/i.test(a.name)) continue;
    await attachmentRules(ctx, cfb, a, depth);
  }
}

async function attachmentRules(ctx: RunContext, cfb: CompoundFile, a: CfbEntry, depth: number): Promise<void> {
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
    if (path) ctx.external(D.Other, path, location, { source: 'attachment by reference', ...(name ? { name: clip(name) } : {}) }, 'attachment');
    return;
  }
  if (method === 7) {
    const path = (await text(P.longPath)) || (await text(P.path));
    if (path) ctx.external(D.Other, path, location, { source: 'attachment by web reference', ...(name ? { name: clip(name) } : {}) }, 'attachment');
    return;
  }
  if (!data) return;
  if (data.kind === 'stream') {
    attachment(ctx, { name: name || undefined, declaredType: declaredType || undefined, size: data.size, location, depth: 1, source: cfb.source(data) }, info);
    return;
  }
  // An embedded message or an OLE object is a storage of this file. The object rules read an object's storage, and
  // an embedded message's attachments are read here.
  info.size = 0;
  if (method === 5) info.type ??= 'application/vnd.ms-outlook';
  ctx.report(C.Embedded, D.Attachment, location, info);
  if (method === 5 && depth < MAX_DEPTH) await messageRules(ctx, cfb, data, depth + 1);
}

/**
 * The place every attached file passes, with what a contained file plugin needs to read it. office-defuse reports it;
 * the family's router, once office-defuse has a plugin for it, is where it would be handed on.
 */
function attachment(ctx: RunContext, file: ContainedFile, info: Record<string, string | number>): void {
  ctx.report(C.Embedded, D.Attachment, file.location, { ...info, size: file.size });
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
