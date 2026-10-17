import * as zlib from 'node:zlib';
import { type CfbEntry, CompoundFile } from './cfb';
import { decideContained, spillOf } from './contained';
import { clip, type Reported, type RunContext } from './context';
import { checkTime, StopError } from './errors';
import { controlStorageProblem, OOXML_CONTROL_STREAMS } from './forms';
import { bufferSink, bufferSource, chunksOf, Reader, readAll, SpillSink } from './io';
import { isFragment } from './links';
import { compoundRules, nestedNotes } from './objects';
import { unlockPresentation } from './offcrypto';
import { utf16 } from './oleobj';
import { CompoundPlan } from './plan';
import { hasPresentationEdits, noPresentationEdits, type PresentationEdits, rewritePresentation } from './ppt-writer';
import { writeCompound } from './rewrite-cfb';
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
 *
 * With `plan`, what the strip findings name is removed by a full save of the stream: the persist objects the current
 * edit reaches, written again with the records that go left out, a new persist directory and a new UserEditAtom.
 */
export async function powerPointRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry, plan?: CompoundPlan): Promise<void> {
  const location = stream.name;
  let source = cfb.source(stream);
  let reader = new Reader(source, stream.size, 16384, 8);
  const persist = new Map<number, number>();
  const edits = noPresentationEdits();
  const current = cfb.child(cfb.root, 'Current User');
  let editAt = 0;
  let session: number | undefined;
  if (current?.kind === 'stream') {
    const cu = await cfb.source(current).read(0, 24);
    if (cu.length >= 20 && header(cu).type === T.currentUserAtom) {
      editAt = u32(cu, 16);
      await readPersist(reader, stream.size, editAt, persist, ctx.settings.deadline);
      if (u32(cu, 12) === ENCRYPTED_TOKEN) {
        const before = ctx.findings.length;
        const opened = await decryptPresentation(ctx, reader, stream, editAt, persist);
        source = opened.source;
        session = opened.session;
        reader = new Reader(source, stream.size, 16384, 8);
        // The rewrite writes the presentation decrypted, so one VelvetSweatshop opens is rewritten when that finding strips.
        if (ctx.findings.slice(before).some(f => f.category === C.Encrypted && f.detail === D.DefaultPassword && f.action === 'strip')) edits.decrypt = true;
      }
    }
  }

  const storages = new Map<number, Storage>();
  /** ExOleObjAtoms by the persist id of the storage they name: the object's type, 0 embedded, 1 linked, 2 a control, and the container that holds the atom. */
  const roles = new Map<number, { type: number; container: number }>();
  /** VBAInfoAtoms by the persist id of the project's storage, with the VBAInfoContainer around each. */
  const vba = new Map<number, number[]>();
  const containers: Array<{ h: Header; offset: number; kids: Array<{ h: Header; offset: number }> }> = [];
  const visited = new Set<number>();
  const walk = async (start: number, end: number, depth: number, siblings: Array<{ h: Header; offset: number }>, parent: number) => {
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
        roles.set(u32(b, 16), { type: u32(b, 4), container: parent });
      } else if (h.type === T.vbaInfoAtom && h.length >= 4) {
        const id = u32(await reader.read(pos + 8, 4), 0);
        vba.set(id, [...(vba.get(id) ?? []), parent]);
      }
      if (h.version === 0x0f && depth < MAX_DEPTH) {
        const kids: Array<{ h: Header; offset: number }> = [];
        await walk(pos + 8, bodyEnd, depth + 1, kids, pos);
        containers.push({ h, offset: pos, kids });
      }
      pos = bodyEnd;
    }
  };
  // Every record at the top level, then each persist object that does not start on a record already read.
  await walk(0, stream.size, 0, [], -1);
  for (const offset of persist.values()) {
    if (visited.has(offset) || offset + 8 > stream.size) continue;
    const h = header(await reader.read(offset, 8));
    await walk(offset, Math.min(stream.size, offset + 8 + h.length), 0, [], -1);
  }

  // The storages come first: a VBA project that stays keeps the actions that run its macros.
  const byOffset = new Map([...persist].map(([id, offset]) => [offset, id]));
  for (const s of storages.values()) {
    const id = byOffset.get(s.offset);
    const role = id !== undefined && vba.has(id) ? 'vba' : id !== undefined && roles.get(id)?.type === 2 ? 'control' : 'object';
    const verdict = await storageRules(ctx, source, s, role, `${location}, offset ${s.offset}`, plan !== undefined);
    if (id === undefined || verdict === 'keep') continue;
    if (verdict === 'remove') {
      // The storage goes, and so does what names it: the VBAInfoContainer, or the object's container in ExObjList.
      edits.dropPersist.add(id);
      if (role === 'vba') for (const c of vba.get(id) ?? []) edits.remove.add(c);
      else {
        const holder = roles.get(id)?.container;
        if (holder !== undefined && holder >= 0) edits.remove.add(holder);
      }
    } else edits.replace.set(s.offset, verdict.body);
  }

  const links = new Map<number, number>();
  const actions: Array<{ offset: number; hyperlink: number; action: number }> = [];
  for (const c of containers) await containerRules(ctx, reader, c.h, c.offset, c.kids, location, edits, links, actions);
  // An action that opens a hyperlink that went does nothing.
  for (const a of actions) {
    if (a.action !== 4 || !links.has(a.hyperlink) || edits.replace.has(a.offset)) continue;
    const body = Uint8Array.from(await reader.read(a.offset + 8, 16));
    body[8] = 0;
    edits.replace.set(a.offset, body);
  }

  if (!plan || !current || !hasPresentationEdits(edits)) return;
  // An encrypted presentation's pictures and summary are encrypted too, in streams this rewrite does not decrypt. It
  // is left as it was, and the check of the output says what it still holds.
  if (session !== undefined && ['Pictures', 'EncryptedSummary'].some(n => (cfb.child(cfb.root, n)?.size ?? 0) > 0)) return;
  if (session !== undefined) {
    edits.decrypt = true;
    edits.dropPersist.add(session);
  }
  const userEdit = header(await reader.read(editAt, 8));
  const body = await reader.read(editAt + 8, Math.min(userEdit.length, 64));
  if (userEdit.type !== T.userEditAtom || body.length < 28) return;
  const sink = new SpillSink(ctx.run.temp, ctx.run.threshold);
  ctx.run.held.push(sink);
  const newEdit = await rewritePresentation(source, stream.size, persist, { body }, edits, sink, ctx.settings.deadline);
  await sink.close();
  plan.replace.set(stream, sink);
  const user = Uint8Array.from(await cfb.bytes(current));
  user.set([newEdit & 0xff, (newEdit >>> 8) & 0xff, (newEdit >>> 16) & 0xff, (newEdit >>> 24) & 0xff], 16);
  if (edits.decrypt) user.set([0x5f, 0xc0, 0x91, 0xe3], 12);
  plan.replace.set(current, await spillOf(ctx.run, user));
}

/**
 * The PowerPoint Document stream of an encrypted presentation, decrypted (MS-PPT 2.3.7). The current edit names the
 * persist object that holds the CryptSession10Container, whose data is an RC4 CryptoAPI EncryptionHeader. Each other
 * persist object, its record header included, is RC4 with the key of block number its persist id. The edits and
 * persist directories were never encrypted.
 */
async function decryptPresentation(ctx: RunContext, reader: Reader, stream: CfbEntry, edit: number, persist: Map<number, number>): Promise<{ source: ByteSource; session: number }> {
  const required = (reason: string) => new StopError(C.Encrypted, D.PasswordRequired, `The presentation is encrypted: ${reason}`, { stream: stream.name, reason });
  const u = await reader.read(edit, 40);
  const sessionId = u.length >= 40 && header(u).type === T.userEditAtom ? u32(u, 36) : -1;
  const sessionAt = persist.get(sessionId);
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
  return { source: bufferSource(copy), session: sessionId };
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

/** What becomes of an ExOleObjStg: it stays as it is, it goes, or its data is replaced. */
type StorageVerdict = 'keep' | 'remove' | { body: Uint8Array };

/**
 * One ExOleObjStg: a VBA project, a control or an OLE object, stored as a compound file that may be compressed. With
 * `rewrite`, a VBA project goes to the script plugins and stays, without its compiled code, when they pass it; a
 * Forms 2.0 control stays; an object stays as the file plugins left it. Anything a strip finding names goes.
 */
async function storageRules(ctx: RunContext, source: ByteSource, s: Storage, role: 'vba' | 'control' | 'object', location: string, rewrite: boolean): Promise<StorageVerdict> {
  const compressed = s.header.instance === 1;
  let inner: ByteSource;
  if (compressed) {
    const raw = await source.read(s.offset + 8, s.header.length);
    const declared = raw.length >= 4 ? u32(raw, 0) : 0;
    if (declared > ctx.settings.decompressedBytes)
      throw new StopError(C.Limit, D.DecompressedSize, `An object at ${location} declares ${declared} bytes`, { name: location, size: declared, limit: ctx.settings.decompressedBytes });
    let bytes: Uint8Array;
    try {
      bytes = zlib.inflateSync(raw.subarray(4), { maxOutputLength: Math.max(1, declared) });
    } catch {
      return unreadable(ctx, role, location, 'its compressed data does not inflate to its declared size');
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
    return unreadable(ctx, role, location, `its storage does not read: ${e.message}`);
  }
  nestedNotes(ctx, cfb, location);
  const strips = ctx.strips;
  if (role === 'control') {
    const data = { storage: location, ...(cfb.root.clsid === '00000000-0000-0000-0000-000000000000' ? {} : { clsid: cfb.root.clsid }) };
    const problem = await controlStorageProblem(cfb, cfb.root, OOXML_CONTROL_STREAMS);
    if (problem) ctx.report(C.Control, D.ActiveX, location, { ...data, reason: clip(problem) });
    else ctx.report(C.Control, D.FormsControl, location, data);
    // What the storage holds besides the control gets the rules any compound file gets.
    await compoundRules(ctx, cfb, location, false);
    return ctx.strips > strips ? 'remove' : 'keep';
  }
  if (role === 'vba') {
    // The project's storage is the root of the compound file, so the rules give it a plan of its own.
    const own = new CompoundPlan();
    await compoundRules(ctx, cfb, location, false, rewrite ? own : undefined);
    const project = cfb.child(cfb.root, 'VBA');
    if (!rewrite || !project || own.drop.has(project)) return ctx.strips > strips ? 'remove' : 'keep';
    // The project stays: written again without its compiled code and whatever else its rules removed.
    if (ctx.strips === strips && !own.replace.size && !own.drop.size) return 'keep';
    return { body: await storageBody(ctx, cfb, own, compressed) };
  }
  // An object's storage goes to the plugins as the compound file it is, and stays as they leave it.
  const size = await inner.size();
  const decision = await decideContained(ctx, { location, chunks: () => chunksOf(inner, size) });
  // A plugin's answer is final, so the object's own rules do not run.
  if (!decision?.answered) await compoundRules(ctx, cfb, location, true, undefined, decision);
  if ((decision && !decision.keep) || ctx.strips > strips) return 'remove';
  if (decision?.replacement) {
    const bytes = await readAll(decision.replacement);
    return { body: compressed ? Buffer.concat([Buffer.from(le32(bytes.length)), zlib.deflateSync(bytes)]) : bytes };
  }
  return 'keep';
}

/** An ExOleObjStg's data for a compound file written with `plan`: deflated after its size when it was compressed. */
async function storageBody(ctx: RunContext, cfb: CompoundFile, plan: CompoundPlan, compressed: boolean): Promise<Uint8Array> {
  const sink = bufferSink();
  await writeCompound(ctx.run, cfb, plan, sink);
  const bytes = sink.result();
  return compressed ? Buffer.concat([Buffer.from(le32(bytes.length)), zlib.deflateSync(bytes)]) : bytes;
}

const le32 = (n: number) => [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff];

function unreadable(ctx: RunContext, role: 'vba' | 'control' | 'object', location: string, reason: string): StorageVerdict {
  const r =
    role === 'vba'
      ? ctx.report(C.Macro, D.VbaProject, location, { storage: location, reason })
      : role === 'control'
        ? ctx.report(C.Control, D.ActiveX, location, { storage: location, reason })
        : ctx.report(C.Embedded, D.OleObject, location, { storage: location, reason });
  return r.action === 'strip' ? 'remove' : 'keep';
}

/**
 * The rules of one container record, from its children: actions, hyperlinks and linked media. What strips goes in the
 * edits: an action becomes no action, a hyperlink goes, and a linked file's path is emptied. `links` collects the ids
 * of the hyperlinks that go, and `actions` every action, for the actions that open those.
 */
async function containerRules(
  ctx: RunContext,
  reader: Reader,
  h: Header,
  offset: number,
  kids: Array<{ h: Header; offset: number }>,
  stream: string,
  edits: PresentationEdits,
  links: Map<number, number>,
  actions: Array<{ offset: number; hyperlink: number; action: number }>,
): Promise<void> {
  const location = `${stream}, offset ${offset}`;
  const strings = async () => {
    const out = new Map<number, string>();
    for (const k of kids) if (k.h.type === T.cString && !out.has(k.h.instance)) out.set(k.h.instance, utf16(await reader.read(k.offset + 8, Math.min(k.h.length, 65536))));
    return out;
  };
  if (h.type === T.interactiveInfo) {
    const atom = kids.find(k => k.h.type === T.interactiveInfoAtom && k.h.length >= 16);
    if (!atom) return;
    const body = Uint8Array.from(await reader.read(atom.offset + 8, 16));
    const action = body[8];
    actions.push({ offset: atom.offset, hyperlink: u32(body, 4), action });
    const name = (await strings()).get(2) ?? '';
    const data = { trigger: h.instance === 1 ? 'mouse over' : 'click', ...(name ? { action: clip(name) } : {}) };
    let r: Reported | undefined;
    if (action === 2) r = ctx.report(C.Action, D.RunProgram, location, data);
    // A macro of a VBA project that stays keeps its action.
    else if (action === 1) r = ctx.macroTrigger(C.Macro, D.Link, location, data, name);
    else if (action === 5) r = ctx.report(C.Action, D.OleVerb, location, data);
    if (r?.action === 'strip') {
      body[8] = 0;
      edits.replace.set(atom.offset, body);
    }
  } else if (h.type === T.exHyperlink) {
    const s = await strings();
    const target = s.get(1) ?? '';
    const atom = kids.find(k => k.h.type === T.exHyperlinkAtom && k.h.length >= 4);
    if (target.trim() && !isFragment(target) && ctx.link(target, location, { source: 'PowerPoint hyperlink' }).action === 'strip') {
      edits.remove.add(offset);
      if (atom) links.set(u32(await reader.read(atom.offset + 8, 4), 0), offset);
    }
  } else if (MEDIA_LINKS.has(h.type)) {
    const path = (await strings()).get(0) ?? '';
    const string = kids.find(k => k.h.type === T.cString && k.h.instance === 0);
    if (path.trim() && ctx.external(D.Other, path, location, { source: 'linked sound or movie' }, 'linked media').action === 'strip' && string) edits.replace.set(string.offset, new Uint8Array(0));
  }
}
