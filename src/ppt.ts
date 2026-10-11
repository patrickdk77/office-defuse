import * as zlib from 'node:zlib';
import { type CfbEntry, CompoundFile } from './cfb';
import { decideContained } from './contained';
import { clip, type RunContext } from './context';
import { checkTime, StopError } from './errors';
import { bufferSource, chunksOf, Reader } from './io';
import { isFragment } from './links';
import { compoundRules, nestedNotes } from './objects';
import { unlockPresentation } from './offcrypto';
import { utf16 } from './oleobj';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';

/** Record types of the PowerPoint binary format (MS-PPT 2.13.24). */
const T = {
  currentUserAtom: 0x0ff6,
  userEditAtom: 0x0ff5,
  persistDirectoryAtom: 0x1772,
  exOleObjStg: 0x1011,
  exOleObjAtom: 0x0fc3,
  vbaInfoAtom: 0x0400,
  interactiveInfo: 0x0ff2,
  interactiveInfoAtom: 0x0ff3,
  cString: 0x0fba,
  exHyperlink: 0x0fd7,
  exHyperlinkAtom: 0x0fd3,
  cryptSession: 0x2f14,
};
/** ExVideoContainer, ExMIDIAudioContainer and ExWAVAudioLinkContainer, whose path names a file outside the presentation. */
const MEDIA_LINKS = new Set([0x1005, 0x100d, 0x1010]);
/** The CurrentUserAtom's header token of an encrypted presentation. */
const ENCRYPTED_TOKEN = 0xf3d1c4df;
const MAX_DEPTH = 64;

interface Header {
  version: number;
  instance: number;
  type: number;
  length: number;
}

const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
const header = (b: Uint8Array): Header => ({ version: b[0] & 0x0f, instance: (b[0] | (b[1] << 8)) >> 4, type: b[2] | (b[3] << 8), length: u32(b, 4) });

/** An ExOleObjStg record: where its data lies and what refers to it. */
interface Storage {
  offset: number;
  header: Header;
}

/**
 * Inspects a PowerPoint 97 to 2003 presentation. PowerPoint finds its objects through the persist directory that the
 * current edit names. This reads every record at the top level of the stream as well as every object the persist
 * directories name, so an object that one route hides is still read.
 */
export async function powerPointRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry): Promise<void> {
  const location = stream.name;
  let source = cfb.source(stream);
  let reader = new Reader(source, stream.size, 16384, 8);
  const persist = new Map<number, number>();
  const current = cfb.child(cfb.root, 'Current User');
  if (current?.kind === 'stream') {
    const cu = await cfb.source(current).read(0, 24);
    if (cu.length >= 20 && header(cu).type === T.currentUserAtom) {
      await readPersist(reader, stream.size, u32(cu, 16), persist, ctx.settings.deadline);
      if (u32(cu, 12) === ENCRYPTED_TOKEN) {
        source = await decryptPresentation(ctx, reader, stream, u32(cu, 16), persist);
        reader = new Reader(source, stream.size, 16384, 8);
      }
    }
  }

  const storages = new Map<number, Storage>();
  /** ExOleObjAtom types by the persist id of the storage they name: 0 embedded, 1 linked, 2 a control. */
  const roles = new Map<number, number>();
  const vba = new Set<number>();
  const visited = new Set<number>();
  const walk = async (start: number, end: number, depth: number, siblings: Array<{ h: Header; offset: number }>) => {
    for (let pos = start; pos + 8 <= end; ) {
      checkTime(ctx.settings.deadline);
      const h = header(await reader.read(pos, 8));
      const bodyEnd = pos + 8 + h.length;
      if (bodyEnd > end || visited.has(pos)) return;
      visited.add(pos);
      siblings.push({ h, offset: pos });
      if (h.type === T.exOleObjStg) storages.set(pos, { offset: pos, header: h });
      else if (h.type === T.exOleObjAtom && h.length >= 20) {
        const b = await reader.read(pos + 8, 20);
        roles.set(u32(b, 16), u32(b, 4));
      } else if (h.type === T.vbaInfoAtom && h.length >= 4) vba.add(u32(await reader.read(pos + 8, 4), 0));
      if (h.version === 0x0f && depth < MAX_DEPTH) {
        const kids: Array<{ h: Header; offset: number }> = [];
        await walk(pos + 8, bodyEnd, depth + 1, kids);
        await containerRules(ctx, reader, h, pos, kids, location);
      }
      pos = bodyEnd;
    }
  };
  // Every record at the top level, then each persist object that does not start on a record already read.
  await walk(0, stream.size, 0, []);
  for (const offset of persist.values()) {
    if (visited.has(offset) || offset + 8 > stream.size) continue;
    const h = header(await reader.read(offset, 8));
    await walk(offset, Math.min(stream.size, offset + 8 + h.length), 0, []);
  }

  const byOffset = new Map([...persist].map(([id, offset]) => [offset, id]));
  for (const s of storages.values()) {
    const id = byOffset.get(s.offset);
    const role = id !== undefined && vba.has(id) ? 'vba' : id !== undefined && roles.get(id) === 2 ? 'control' : 'object';
    await storageRules(ctx, source, s, role, `${location}, offset ${s.offset}`);
  }
}

/**
 * The PowerPoint Document stream of an encrypted presentation, decrypted (MS-PPT 2.3.7). The current edit names the
 * persist object that holds the CryptSession10Container, whose data is an RC4 CryptoAPI EncryptionHeader. Each other
 * persist object, its record header included, is RC4 with the key of block number its persist id. The edits and
 * persist directories were never encrypted.
 */
async function decryptPresentation(ctx: RunContext, reader: Reader, stream: CfbEntry, edit: number, persist: Map<number, number>): Promise<ByteSource> {
  const required = (reason: string) => new StopError(C.Encrypted, D.PasswordRequired, `The presentation is encrypted: ${reason}`, { stream: stream.name, reason });
  const u = await reader.read(edit, 40);
  const sessionAt = u.length >= 40 && header(u).type === T.userEditAtom ? persist.get(u32(u, 36)) : undefined;
  if (sessionAt === undefined || sessionAt + 8 > stream.size) throw required('the current edit names no CryptSession10Container');
  const h = header(await reader.read(sessionAt, 8));
  if (h.type !== T.cryptSession) throw required('the current edit names no CryptSession10Container');
  const keyOf = unlockPresentation(ctx, await reader.read(sessionAt + 8, Math.min(h.length, 65536)), stream.name);
  if (stream.size > ctx.settings.decompressedBytes)
    throw new StopError(C.Limit, D.DecompressedSize, `The encrypted stream ${stream.name} is ${stream.size} bytes`, { name: stream.name, size: stream.size, limit: ctx.settings.decompressedBytes });
  const copy = Buffer.from(await reader.read(0, stream.size));
  for (const [id, offset] of persist) {
    checkTime(ctx.settings.deadline);
    if (offset === sessionAt || offset + 8 > copy.length) continue;
    const rc4 = keyOf(id);
    const head = rc4.crypt(copy.subarray(offset, offset + 8));
    const length = Math.min(copy.length - offset - 8, u32(head, 4));
    copy.set(head, offset);
    copy.set(rc4.crypt(copy.subarray(offset + 8, offset + 8 + length)), offset + 8);
  }
  return bufferSource(copy);
}

/** Reads the persist directories of the current edit and every edit before it. Newer entries win. */
async function readPersist(reader: Reader, size: number, start: number, persist: Map<number, number>, deadline?: number): Promise<void> {
  const seen = new Set<number>();
  for (let edit = start; edit > 0 && edit + 36 <= size && !seen.has(edit); ) {
    checkTime(deadline);
    seen.add(edit);
    const u = await reader.read(edit, 36);
    if (header(u).type !== T.userEditAtom) return;
    const dirAt = u32(u, 20);
    const d = await reader.read(dirAt, 8);
    if (d.length === 8 && header(d).type === T.persistDirectoryAtom) {
      const body = await reader.read(dirAt + 8, Math.min(header(d).length, size - dirAt - 8));
      for (let p = 0; p + 4 <= body.length; ) {
        const entry = u32(body, p);
        const first = entry & 0xfffff;
        const count = entry >>> 20;
        p += 4;
        for (let i = 0; i < count && p + 4 <= body.length; i++, p += 4) if (!persist.has(first + i)) persist.set(first + i, u32(body, p));
      }
    }
    edit = u32(u, 16);
  }
}

/** One ExOleObjStg: a VBA project, a control or an OLE object, stored as a compound file that may be compressed. */
async function storageRules(ctx: RunContext, source: ByteSource, s: Storage, role: 'vba' | 'control' | 'object', location: string): Promise<void> {
  let inner: ByteSource;
  if (s.header.instance === 1) {
    const raw = await source.read(s.offset + 8, s.header.length);
    const declared = raw.length >= 4 ? u32(raw, 0) : 0;
    if (declared > ctx.settings.decompressedBytes)
      throw new StopError(C.Limit, D.DecompressedSize, `An object at ${location} declares ${declared} bytes`, { name: location, size: declared, limit: ctx.settings.decompressedBytes });
    let bytes: Uint8Array;
    try {
      bytes = zlib.inflateSync(raw.subarray(4), { maxOutputLength: Math.max(1, declared) });
    } catch {
      unreadable(ctx, role, location, 'its compressed data does not inflate to its declared size');
      return;
    }
    inner = { size: async () => bytes.length, read: async (offset, length) => bytes.subarray(offset, offset + length) };
  } else {
    const start = s.offset + 8;
    const length = s.header.length;
    inner = { size: async () => length, read: (offset, n) => source.read(start + offset, Math.max(0, Math.min(n, length - offset))) };
  }
  const head = await inner.read(0, 8);
  let cfb: CompoundFile;
  try {
    if (head.length < 8 || head[0] !== 0xd0 || head[1] !== 0xcf) throw new StopError(C.Container, D.BadHeader, 'not a compound file');
    cfb = await CompoundFile.open(inner, ctx.settings, ctx.settings.deadline);
  } catch (e) {
    if (!(e instanceof StopError) || e.category !== C.Container) throw e;
    unreadable(ctx, role, location, `its storage does not read: ${e.message}`);
    return;
  }
  nestedNotes(ctx, cfb, location);
  if (role === 'control') ctx.report(C.Control, D.ActiveX, location, { storage: location, ...(cfb.root.clsid === '00000000-0000-0000-0000-000000000000' ? {} : { clsid: cfb.root.clsid }) });
  // An object's storage goes to the plugins as the compound file it is. The rewrite cannot change the presentation's
  // records, so only a plugin that passes the object keeps it in the output.
  const size = await inner.size();
  const decision = role === 'object' ? await decideContained(ctx, { location, chunks: () => chunksOf(inner, size) }) : undefined;
  await compoundRules(ctx, cfb, location, role === 'object', undefined, decision);
}

function unreadable(ctx: RunContext, role: 'vba' | 'control' | 'object', location: string, reason: string): void {
  if (role === 'vba') ctx.report(C.Macro, D.VbaProject, location, { storage: location, reason });
  else if (role === 'control') ctx.report(C.Control, D.ActiveX, location, { storage: location, reason });
  else ctx.report(C.Embedded, D.OleObject, location, { storage: location, reason });
}

/** The rules of one container record, from its children: actions, hyperlinks and linked media. */
async function containerRules(ctx: RunContext, reader: Reader, h: Header, offset: number, kids: Array<{ h: Header; offset: number }>, stream: string): Promise<void> {
  const location = `${stream}, offset ${offset}`;
  const strings = async () => {
    const out = new Map<number, string>();
    for (const k of kids) if (k.h.type === T.cString && !out.has(k.h.instance)) out.set(k.h.instance, utf16(await reader.read(k.offset + 8, Math.min(k.h.length, 65536))));
    return out;
  };
  if (h.type === T.interactiveInfo) {
    const atom = kids.find(k => k.h.type === T.interactiveInfoAtom && k.h.length >= 16);
    if (!atom) return;
    const action = (await reader.read(atom.offset + 8 + 8, 1))[0];
    const name = (await strings()).get(2) ?? '';
    const data = { trigger: h.instance === 1 ? 'mouse over' : 'click', ...(name ? { action: clip(name) } : {}) };
    if (action === 2) ctx.report(C.Action, D.RunProgram, location, data);
    else if (action === 1) ctx.report(C.Macro, D.Link, location, data);
    else if (action === 5) ctx.report(C.Action, D.OleVerb, location, data);
  } else if (h.type === T.exHyperlink) {
    const s = await strings();
    const target = s.get(1) ?? '';
    if (target.trim() && !isFragment(target)) ctx.link(target, location, { source: 'PowerPoint hyperlink' });
  } else if (MEDIA_LINKS.has(h.type)) {
    const path = (await strings()).get(0) ?? '';
    if (path.trim()) ctx.external(D.Other, path, location, { source: 'linked sound or movie' }, 'linked media');
  }
}
