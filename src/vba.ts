import { createHash } from 'node:crypto';
import type { CfbEntry, CompoundFile } from './cfb';

/**
 * A VBA project as MS-OVBA stores one: a VBA storage with a dir stream that lists the modules, and a stream per module
 * that holds the module's compiled p-code and then its compressed source. Office runs the p-code when the same VBA
 * version compiled it, so the p-code can do something other than the source says (VBA stomping). A project kept after
 * its source was checked is written without the p-code, and with the source compressed again, so Office compiles the
 * text that was checked.
 */

export class VbaError extends Error {}

const CHUNK = 4096;

/**
 * Decompresses an MS-OVBA CompressedContainer (2.4.1.3.1). Throws VbaError when a copy token reaches before its chunk,
 * a chunk decompresses past 4096 bytes, or the output passes `limit`. A chunk header without its signature bits is
 * read like any other, and fewer than 3 bytes at the end are taken as padding: Office opens files written that way.
 */
export function decompress(data: Uint8Array, limit = 64 * 1024 * 1024): Uint8Array {
  if (data.length === 0 || data[0] !== 0x01) throw new VbaError('a compressed container without its signature byte');
  let out = new Uint8Array(Math.min(limit, Math.max(CHUNK, data.length * 4)));
  let n = 0;
  const room = (k: number) => {
    if (n + k <= out.length) return;
    if (n + k > limit) throw new VbaError('decompressed data past the limit');
    const bigger = new Uint8Array(Math.min(limit, Math.max(out.length * 2, n + k)));
    bigger.set(out.subarray(0, n));
    out = bigger;
  };
  let pos = 1;
  while (pos + 3 <= data.length) {
    const header = data[pos] | (data[pos + 1] << 8);
    const end = Math.min(pos + (header & 0x0fff) + 3, data.length);
    const start = n;
    pos += 2;
    if ((header & 0x8000) === 0) {
      if (pos + CHUNK > data.length) throw new VbaError('an uncompressed chunk cut short');
      room(CHUNK);
      out.set(data.subarray(pos, pos + CHUNK), n);
      n += CHUNK;
      pos += CHUNK;
      continue;
    }
    while (pos < end) {
      const flags = data[pos++];
      for (let bit = 0; bit < 8 && pos < end; bit++) {
        if ((flags & (1 << bit)) === 0) {
          room(1);
          out[n++] = data[pos++];
          continue;
        }
        if (pos + 2 > end) throw new VbaError('a copy token cut short');
        const token = data[pos] | (data[pos + 1] << 8);
        pos += 2;
        const help = copyTokenHelp(n - start);
        const length = (token & help.lengthMask) + 3;
        const from = n - ((token >> (16 - help.bitCount)) + 1);
        if (from < start) throw new VbaError('a copy token that reaches before its chunk');
        room(length);
        // The source and the copy can overlap, so it goes a byte at a time.
        for (let i = 0; i < length; i++) out[n + i] = out[from + i];
        n += length;
      }
      if (n - start > CHUNK) throw new VbaError('a chunk that decompresses past 4096 bytes');
    }
  }
  return out.slice(0, n);
}

/** MS-OVBA 2.4.1.3.19.1: the masks of a copy token at `difference` bytes into its chunk. */
function copyTokenHelp(difference: number): { lengthMask: number; bitCount: number; maximumLength: number } {
  let bitCount = 4;
  while (1 << bitCount < difference) bitCount++;
  const lengthMask = 0xffff >> bitCount;
  return { lengthMask, bitCount, maximumLength: lengthMask + 3 };
}

/** Candidates a match search looks at for one position. Past it, a shorter match only costs a little space. */
const MAX_CANDIDATES = 256;

/**
 * Compresses `data` into an MS-OVBA CompressedContainer (2.4.1.3.6). The longest match comes from a hash chain of
 * the positions in the chunk. Throws VbaError for a last chunk that compresses past 4096 bytes, which MS-OVBA can only
 * write as a raw chunk padded to 4096 bytes, with padding Office would read as source.
 */
export function compress(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  for (let chunkStart = 0; chunkStart < data.length; chunkStart += CHUNK) {
    const chunkEnd = Math.min(chunkStart + CHUNK, data.length);
    const body: number[] = [];
    const heads = new Map<number, number>();
    const prev = new Int32Array(chunkEnd - chunkStart).fill(-1);
    const key = (i: number) => data[i] | (data[i + 1] << 8) | (data[i + 2] << 16);
    let indexed = chunkStart;
    const index = (upTo: number) => {
      for (; indexed < upTo; indexed++) {
        if (indexed + 3 > chunkEnd) continue;
        const k = key(indexed);
        prev[indexed - chunkStart] = heads.get(k) ?? -1;
        heads.set(k, indexed);
      }
    };
    let cur = chunkStart;
    while (cur < chunkEnd) {
      const flagAt = body.length;
      body.push(0);
      let flags = 0;
      for (let bit = 0; bit < 8 && cur < chunkEnd; bit++) {
        index(cur);
        const help = copyTokenHelp(cur - chunkStart);
        let bestLength = 0;
        let bestCandidate = 0;
        if (cur + 3 <= chunkEnd) {
          let candidate = heads.get(key(cur)) ?? -1;
          for (let tries = 0; candidate >= 0 && tries < MAX_CANDIDATES; tries++) {
            let len = 0;
            while (cur + len < chunkEnd && len < help.maximumLength && data[candidate + len] === data[cur + len]) len++;
            if (len > bestLength) {
              bestLength = len;
              bestCandidate = candidate;
              if (len === help.maximumLength) break;
            }
            candidate = prev[candidate - chunkStart];
          }
        }
        if (bestLength >= 3) {
          const token = (((cur - bestCandidate - 1) << (16 - help.bitCount)) | (bestLength - 3)) & 0xffff;
          body.push(token & 0xff, token >> 8);
          flags |= 1 << bit;
          cur += bestLength;
        } else {
          body.push(data[cur]);
          cur += 1;
        }
      }
      body[flagAt] = flags;
    }
    if (body.length <= CHUNK) {
      const header = (body.length - 1) | 0xb000;
      out.push(header & 0xff, header >> 8);
      for (const b of body) out.push(b);
    } else if (chunkEnd - chunkStart === CHUNK) {
      // A raw chunk: the header with the flag clear, then the 4096 bytes as they are.
      const header = 0x3fff;
      out.push(header & 0xff, header >> 8);
      for (let i = chunkStart; i < chunkEnd; i++) out.push(data[i]);
    } else throw new VbaError('a last chunk that does not compress to 4096 bytes');
  }
  return Uint8Array.from(out);
}

/** One record of the dir stream (MS-OVBA 2.3.4.2), as read. */
interface DirRecord {
  id: number;
  data: Uint8Array;
}

const R = {
  codePage: 0x0003,
  version: 0x0009,
  referenceName: 0x0016,
  referenceRegistered: 0x000d,
  referenceProject: 0x000e,
  referenceControl: 0x002f,
  referenceExtended: 0x0030,
  referenceOriginal: 0x0033,
  moduleName: 0x0019,
  moduleStreamName: 0x001a,
  moduleOffset: 0x0031,
  moduleProcedural: 0x0021,
  moduleTerminator: 0x002b,
  dirTerminator: 0x0010,
} as const;

function readRecords(dir: Uint8Array): DirRecord[] {
  const records: DirRecord[] = [];
  let pos = 0;
  while (pos < dir.length) {
    if (pos + 6 > dir.length) throw new VbaError('a dir record cut short');
    const id = dir[pos] | (dir[pos + 1] << 8);
    // PROJECTVERSION says 4 and holds 6: a major version of 4 bytes and a minor one of 2.
    const size = id === R.version ? 6 : u32(dir, pos + 2);
    if (pos + 6 + size > dir.length) throw new VbaError('a dir record that runs past the stream');
    records.push({ id, data: dir.subarray(pos + 6, pos + 6 + size) });
    pos += 6 + size;
    if (id === R.dirTerminator) return records;
  }
  throw new VbaError('a dir stream without its terminator');
}

function writeRecords(records: DirRecord[]): Uint8Array {
  const parts: number[] = [];
  for (const r of records) {
    const size = r.id === R.version ? 4 : r.data.length;
    parts.push(r.id & 0xff, r.id >> 8, size & 0xff, (size >> 8) & 0xff, (size >> 16) & 0xff, (size >>> 24) & 0xff);
    for (const b of r.data) parts.push(b);
  }
  return Uint8Array.from(parts);
}

/** A module of a VBA project. */
export interface VbaModule {
  name: string;
  /**
   * What the PROJECT stream says the module is: a document's or a sheet's module, a code module, a class module or a
   * user form. A module the PROJECT stream leaves out is a code or class module as its dir record says.
   */
  kind: 'document' | 'standard' | 'class' | 'form';
  stream: CfbEntry;
  /** Where the compressed source starts in the stream. The bytes before it are p-code. */
  offset: number;
  /** The source as stored, in the project's code page. */
  bytes: Uint8Array;
  source: string;
}

/** A type library or project the VBA project refers to (MS-OVBA 2.3.4.2.2). */
export interface VbaReference {
  name: string;
  kind: 'registered' | 'control' | 'project';
  /** The library's LibidReference (2.1.1.8), or for a project reference its absolute path (2.1.1.12). */
  libid: string;
}

export interface VbaProject {
  storage: CfbEntry;
  dir: CfbEntry;
  codePage: number;
  modules: VbaModule[];
  references: VbaReference[];
  records: DirRecord[];
  /** The MD5 of the references and the modules' names, kinds and sources: what a plugin's verdict is about. */
  digest: string;
}

const u32 = (b: Uint8Array, at = 0) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;

interface Decoder {
  decode(bytes: Uint8Array): string;
}

/** The text decoder for a project's code page, with Latin-1 for one Node has no decoder for. */
function decoderFor(codePage: number): Decoder {
  const labels: Record<number, string> = { 65001: 'utf-8', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 10000: 'macintosh' };
  try {
    return new TextDecoder(labels[codePage] ?? `windows-${codePage}`);
  } catch {
    return new TextDecoder('latin1');
  }
}

/** A string that starts with its 4-byte length, as REFERENCEREGISTERED and the others hold their libids. */
function sized(b: Uint8Array, at: number, decoder: Decoder): { text: string; next: number } {
  if (at + 4 > b.length) throw new VbaError('a reference record cut short');
  const n = u32(b, at);
  if (at + 4 + n > b.length) throw new VbaError('a reference record cut short');
  return { text: decoder.decode(b.subarray(at + 4, at + 4 + n)), next: at + 4 + n };
}

function readReferences(records: DirRecord[], decoder: Decoder): VbaReference[] {
  const refs: VbaReference[] = [];
  let name = '';
  let original: string | undefined;
  let control: VbaReference | undefined;
  for (const r of records) {
    if (r.id === R.moduleName) break;
    if (r.id === R.referenceName) {
      // Inside a REFERENCECONTROL, a name record is the control library's extended name, not a new reference.
      if (!control) name = decoder.decode(r.data);
    } else if (r.id === R.referenceOriginal) original = decoder.decode(r.data);
    else if (r.id === R.referenceRegistered) refs.push({ name, kind: 'registered', libid: sized(r.data, 0, decoder).text });
    else if (r.id === R.referenceProject) refs.push({ name, kind: 'project', libid: sized(r.data, 0, decoder).text });
    else if (r.id === R.referenceControl) {
      control = { name, kind: 'control', libid: original ?? sized(r.data, 0, decoder).text };
      refs.push(control);
    } else if (r.id === R.referenceExtended && control) {
      // The extended libid names the library the control comes from, which the twiddled one only stands in for.
      if (original === undefined) control.libid = sized(r.data, 0, decoder).text;
      control = undefined;
      original = undefined;
    }
  }
  return refs;
}

/** The module kinds the PROJECT stream lists (MS-OVBA 2.3.1), by module name in upper case. */
function projectKinds(text: string): Map<string, VbaModule['kind']> {
  const kinds = new Map<string, VbaModule['kind']>();
  for (const line of text.split(/\r\n|\r|\n/)) {
    if (line.startsWith('[')) break;
    const m = /^(Document|Module|Class|BaseClass)=([^/]*)/.exec(line);
    if (!m) continue;
    const kind = m[1] === 'Document' ? 'document' : m[1] === 'Module' ? 'standard' : m[1] === 'Class' ? 'class' : 'form';
    kinds.set(m[2].toUpperCase(), kind);
  }
  return kinds;
}

/**
 * Reads the project in `vba`, the VBA storage: the dir stream, the references, and each module's source. Throws
 * VbaError when a stream the dir names is missing or a source does not decompress, which leaves nothing to check.
 */
export async function readProject(cfb: CompoundFile, vba: CfbEntry, limit = 64 * 1024 * 1024): Promise<VbaProject> {
  const find = (storage: CfbEntry, name: string) => storage.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === name.toUpperCase());
  const dirEntry = find(vba, 'dir');
  if (!dirEntry) throw new VbaError('a VBA storage without its dir stream');
  if (dirEntry.size > limit) throw new VbaError('a dir stream past the limit');
  const records = readRecords(decompress(await cfb.bytes(dirEntry), limit));
  const codePageRecord = records.find(r => r.id === R.codePage);
  const codePage = codePageRecord && codePageRecord.data.length >= 2 ? codePageRecord.data[0] | (codePageRecord.data[1] << 8) : 1252;
  const decoder = decoderFor(codePage);
  const projectStream = vba.parent ? find(vba.parent, 'PROJECT') : undefined;
  const kinds = projectStream && projectStream.size <= limit ? projectKinds(decoder.decode(await cfb.bytes(projectStream))) : new Map<string, VbaModule['kind']>();
  const modules: VbaModule[] = [];
  let current: { name?: string; streamName?: string; offset?: number; procedural: boolean } | undefined;
  let total = 0;
  for (const r of records) {
    if (r.id === R.moduleName) current = { name: decoder.decode(r.data), procedural: false };
    else if (!current) continue;
    else if (r.id === R.moduleStreamName) current.streamName = decoder.decode(r.data);
    else if (r.id === R.moduleOffset && r.data.length >= 4) current.offset = u32(r.data);
    else if (r.id === R.moduleProcedural) current.procedural = true;
    else if (r.id === R.moduleTerminator) {
      const name = current.name ?? '';
      const streamName = current.streamName ?? name;
      const stream = find(vba, streamName);
      if (!stream) throw new VbaError(`a module stream the dir names is missing: ${streamName}`);
      if (stream.size > limit) throw new VbaError('a module stream past the limit');
      const offset = current.offset ?? 0;
      const raw = await cfb.bytes(stream);
      if (offset > raw.length) throw new VbaError('a module whose source starts past its stream');
      const bytes = decompress(raw.subarray(offset), limit);
      total += bytes.length;
      if (total > limit) throw new VbaError('sources past the limit');
      const kind = kinds.get(name.toUpperCase()) ?? (current.procedural ? 'standard' : 'class');
      modules.push({ name, kind, stream, offset, bytes, source: decoder.decode(bytes) });
      current = undefined;
    }
  }
  const references = readReferences(records, decoder);
  const hash = createHash('md5');
  hash.update(JSON.stringify(references));
  for (const m of modules) hash.update(JSON.stringify([m.name, m.kind, m.stream.name])).update(m.bytes);
  return { storage: vba, dir: dirEntry, codePage, modules, references, records, digest: hash.digest('hex') };
}

/** The `_VBA_PROJECT` stream of a project no VBA version compiled, which makes Office compile the source. */
export const PURGED_VBA_PROJECT = Uint8Array.of(0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00);

/** Streams of the VBA storage that are neither the dir stream, `_VBA_PROJECT`, nor a module the dir names. */
function otherStreams(project: VbaProject): CfbEntry[] {
  const known = new Set<CfbEntry>([project.dir, ...project.modules.map(m => m.stream)]);
  return project.storage.children.filter(c => !known.has(c) && c.name.toUpperCase() !== '_VBA_PROJECT');
}

/**
 * The project's streams without p-code, as Mandiant's OfficePurge writes them: each module stream holds only its
 * source, compressed again; the dir stream says each source starts at 0; `_VBA_PROJECT` names no VBA version. The
 * `__SRP_` caches and anything else in the VBA storage that the dir does not name go.
 */
export function purge(project: VbaProject): { replace: Map<CfbEntry, Uint8Array>; drop: CfbEntry[] } {
  const replace = new Map<CfbEntry, Uint8Array>();
  for (const m of project.modules) replace.set(m.stream, compress(m.bytes));
  const records = project.records.map(r => (r.id === R.moduleOffset ? { id: r.id, data: new Uint8Array(4) } : r));
  replace.set(project.dir, compress(writeRecords(records)));
  const vbaProject = project.storage.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === '_VBA_PROJECT');
  if (vbaProject) replace.set(vbaProject, PURGED_VBA_PROJECT);
  return { replace, drop: otherStreams(project) };
}

/** Whether the project is as `purge` writes it, so writing it again changes nothing. */
export async function isPurged(cfb: CompoundFile, project: VbaProject): Promise<boolean> {
  if (otherStreams(project).length) return false;
  const vbaProject = project.storage.children.find(c => c.kind === 'stream' && c.name.toUpperCase() === '_VBA_PROJECT');
  if (!vbaProject || !same(await cfb.bytes(vbaProject), PURGED_VBA_PROJECT)) return false;
  for (const m of project.modules) if (m.offset !== 0 || !same(await cfb.bytes(m.stream), compress(m.bytes))) return false;
  return true;
}

const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((x, i) => x === b[i]);
