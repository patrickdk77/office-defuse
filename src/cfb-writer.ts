import { checkTime } from './errors';
import type { ByteSink, ByteSource } from './types';

const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
const FREE = 0xffffffff;
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const NO_STREAM = 0xffffffff;
const HEADER_FAT = 109;

/** A stream or storage to write. A storage's children may come in any order; the writer sorts them. */
export type CfbNode = { kind: 'storage'; name: string; clsid: string; children: CfbNode[] } | { kind: 'stream'; name: string; size: number; source: ByteSource };

/** The root of the file to write: its class id and what it holds. */
export interface CfbTree {
  clsid: string;
  children: CfbNode[];
}

/** A GUID as Windows shows it, back to its 16 bytes. '' gives zeros. */
function guidBytes(clsid: string): Uint8Array {
  const out = new Uint8Array(16);
  const h = clsid.replace(/-/g, '');
  if (h.length !== 32) return out;
  const byte = (i: number) => Number.parseInt(h.slice(i * 2, i * 2 + 2), 16);
  const order = [3, 2, 1, 0, 5, 4, 7, 6, 8, 9, 10, 11, 12, 13, 14, 15];
  for (const [i, from] of order.entries()) out[i] = byte(from);
  return out;
}

/** MS-CFB's order of names in a storage: the shorter first, then by their upper-case code units. */
function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

interface Entry {
  name: string;
  type: 1 | 2 | 5;
  clsid: Uint8Array;
  left: number;
  right: number;
  child: number;
  red: boolean;
  start: number;
  size: number;
  node?: CfbNode;
}

class Bytes {
  readonly b: Buffer;
  constructor(n: number) {
    this.b = Buffer.alloc(n);
  }
  u16(at: number, v: number): void {
    this.b.writeUInt16LE(v, at);
  }
  u32(at: number, v: number): void {
    this.b.writeUInt32LE(v >>> 0, at);
  }
}

/**
 * Writes a compound file (MS-CFB, version 3 with 512-byte sectors) to an append-only sink. The children of each
 * storage go into a balanced tree in MS-CFB's name order, colored as a valid red-black tree. Streams under 4,096 bytes
 * go into the mini stream. The writer holds the allocation tables and the directory, and reads each stream once, in
 * order, from its source.
 */
export async function writeCfb(tree: CfbTree, sink: ByteSink, deadline?: number): Promise<void> {
  const ss = 512;
  const per = ss / 4;
  // The directory: the root, then every storage's children, depth first.
  const entries: Entry[] = [{ name: 'Root Entry', type: 5, clsid: guidBytes(tree.clsid), left: NO_STREAM, right: NO_STREAM, child: NO_STREAM, red: false, start: END_OF_CHAIN, size: 0 }];
  const place = (children: CfbNode[]): number => {
    const sorted = [...children].sort((a, b) => compareNames(a.name, b.name));
    const ids: number[] = [];
    for (const c of sorted) {
      ids.push(entries.length);
      entries.push({
        name: c.name,
        type: c.kind === 'storage' ? 1 : 2,
        clsid: c.kind === 'storage' ? guidBytes(c.clsid) : new Uint8Array(16),
        left: NO_STREAM,
        right: NO_STREAM,
        child: NO_STREAM,
        red: false,
        start: END_OF_CHAIN,
        size: c.kind === 'stream' ? c.size : 0,
        node: c,
      });
    }
    // A median split puts every leaf on the last two levels. The deepest level is red and the rest black, so every
    // path from the root down passes the same number of black entries.
    const height = Math.ceil(Math.log2(sorted.length + 1));
    const build = (lo: number, hi: number, depth: number): number => {
      if (lo > hi) return NO_STREAM;
      const mid = (lo + hi) >>> 1;
      const e = entries[ids[mid]];
      e.left = build(lo, mid - 1, depth + 1);
      e.right = build(mid + 1, hi, depth + 1);
      e.red = depth === height - 1 && height > 1;
      return ids[mid];
    };
    const root = build(0, sorted.length - 1, 0);
    for (const [i, c] of sorted.entries()) if (c.kind === 'storage') entries[ids[i]].child = place(c.children);
    return root;
  };
  entries[0].child = place(tree.children);

  // Sectors: the FAT and the DIFAT, then the mini stream, the large streams, the directory and the mini FAT, then any
  // free sectors the tail needs. The first sector after the header is always a FAT sector, whose sector numbers can
  // never spell a signature, so no stream lands where PDF readers look for a header.
  // Version 3 keeps sizes in 32 bits.
  if (entries.some(e => e.size > 0xffffffff)) throw new Error('A stream is too large for a version 3 compound file');
  const small = entries.filter(e => e.type === 2 && e.size > 0 && e.size < MINI_CUTOFF);
  const large = entries.filter(e => e.type === 2 && e.size >= MINI_CUTOFF);
  let miniSectors = 0;
  const miniStart = new Map<Entry, number>();
  for (const e of small) {
    miniStart.set(e, miniSectors);
    miniSectors += Math.ceil(e.size / MINI_SECTOR);
  }
  const miniStreamSize = miniSectors * MINI_SECTOR;
  const miniStreamSectors = Math.ceil(miniStreamSize / ss);
  const largeSectors = large.reduce((t, e) => t + Math.ceil(e.size / ss), 0);
  const dirSectors = Math.ceil(entries.length / (ss / 128));
  const miniFatSectors = Math.ceil(miniSectors / per);
  const content = miniStreamSectors + largeSectors + dirSectors + miniFatSectors;

  /** Where everything goes with `pad` free sectors at the end. */
  const layout = (pad: number) => {
    // The FAT must describe its own sectors and the DIFAT's, so their counts grow together until they fit.
    let fatSectors = 0;
    let difatSectors = 0;
    for (;;) {
      const total = content + pad + fatSectors + difatSectors;
      const needFat = Math.ceil(total / per);
      const needDifat = needFat > HEADER_FAT ? Math.ceil((needFat - HEADER_FAT) / (per - 1)) : 0;
      if (needFat === fatSectors && needDifat === difatSectors) break;
      fatSectors = needFat;
      difatSectors = needDifat;
    }
    const base = fatSectors + difatSectors;
    let next = base + miniStreamSectors;
    const starts = new Map<Entry, number>();
    for (const e of large) {
      starts.set(e, next);
      next += Math.ceil(e.size / ss);
    }
    const dirStart = next;
    const miniFatStart = miniSectors ? dirStart + dirSectors : END_OF_CHAIN;
    const sectors = base + content + pad;
    return { fatSectors, difatSectors, base, starts, dirStart, miniFatStart, pad, sectors, size: (sectors + 1) * ss };
  };
  type Layout = ReturnType<typeof layout>;

  const apply = (l: Layout) => {
    for (const e of small) e.start = miniStart.get(e) ?? 0;
    for (const e of large) e.start = l.starts.get(e) ?? 0;
    entries[0].start = miniStreamSectors ? l.base : END_OF_CHAIN;
    entries[0].size = miniStreamSize;
  };

  const fatOf = (l: Layout) => {
    const fat = new Uint32Array(l.fatSectors * per).fill(FREE);
    const chain = (start: number, count: number) => {
      for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? END_OF_CHAIN : start + i + 1;
    };
    for (let i = 0; i < l.fatSectors; i++) fat[i] = FAT_SECTOR;
    for (let i = 0; i < l.difatSectors; i++) fat[l.fatSectors + i] = DIFAT_SECTOR;
    if (miniStreamSectors) chain(l.base, miniStreamSectors);
    for (const e of large) chain(e.start, Math.ceil(e.size / ss));
    chain(l.dirStart, dirSectors);
    if (miniFatSectors) chain(l.miniFatStart, miniFatSectors);
    return fat;
  };

  const directoryOf = () => {
    const dir = new Bytes(dirSectors * ss);
    for (let id = 0; id < dirSectors * (ss / 128); id++) {
      const at = id * 128;
      const e = entries[id];
      if (!e) {
        dir.u32(at + 0x44, NO_STREAM);
        dir.u32(at + 0x48, NO_STREAM);
        dir.u32(at + 0x4c, NO_STREAM);
        continue;
      }
      const name = Buffer.from(e.name, 'utf16le').subarray(0, 62);
      name.copy(dir.b, at);
      dir.u16(at + 0x40, name.length + 2);
      dir.b[at + 0x42] = e.type;
      dir.b[at + 0x43] = e.red ? 0 : 1;
      dir.u32(at + 0x44, e.left);
      dir.u32(at + 0x48, e.right);
      dir.u32(at + 0x4c, e.child);
      Buffer.from(e.clsid).copy(dir.b, at + 0x50);
      dir.u32(at + 0x74, e.type === 1 ? 0 : e.size === 0 && e.type === 2 ? END_OF_CHAIN : e.start);
      dir.u32(at + 0x78, e.size);
    }
    return dir.b;
  };

  const miniFat = new Uint32Array(miniFatSectors * per).fill(FREE);
  for (const e of small) {
    const first = miniStart.get(e) ?? 0;
    const n = Math.ceil(e.size / MINI_SECTOR);
    for (let i = 0; i < n; i++) miniFat[first + i] = i === n - 1 ? END_OF_CHAIN : first + i + 1;
  }

  /**
   * The bytes the file holds in [from, to), with streams read from their sources. The header and the DIFAT read as
   * zeros, since what the header holds is fixed and the DIFAT holds only sector numbers.
   */
  const bytesAt = async (l: Layout, from: number, to: number): Promise<Buffer> => {
    const out = Buffer.alloc(Math.max(0, to - from));
    const put = (at: number, b: Uint8Array) => {
      const s = Math.max(at, from);
      const e = Math.min(at + b.length, to);
      if (e > s) out.set(b.subarray(s - at, e - at), s - from);
    };
    const putStream = async (at: number, e: Entry) => {
      const node = e.node;
      const s = Math.max(at, from);
      const end = Math.min(at + e.size, to);
      if (end <= s || node?.kind !== 'stream') return;
      put(s, await node.source.read(s - at, end - s));
    };
    const fat = fatOf(l);
    put(ss, Buffer.from(fat.buffer));
    for (const e of small) await putStream((l.base + 1) * ss + (miniStart.get(e) ?? 0) * MINI_SECTOR, e);
    for (const e of large) await putStream((e.start + 1) * ss, e);
    put((l.dirStart + 1) * ss, directoryOf());
    if (miniFatSectors) put((l.miniFatStart + 1) * ss, Buffer.from(miniFat.buffer));
    return out;
  };

  // ZIP readers look for an end record in the last 65,557 bytes. A stream that ends with one, such as an embedded
  // package, would make the file open as that archive too, so free sectors of zeros push it out of their reach.
  const ZIP_WINDOW = 22 + 0xffff;
  const END_RECORD = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
  const lastEndRecord = async (l: Layout): Promise<number> => {
    const from = Math.max(ss, l.size - ZIP_WINDOW);
    const found = (await bytesAt(l, from, l.size)).lastIndexOf(END_RECORD);
    return found < 0 ? -1 : from + found;
  };
  let plan = layout(0);
  apply(plan);
  const endAt = await lastEndRecord(plan);
  if (endAt >= 0) {
    plan = layout(Math.ceil((endAt + 1 - (plan.size - ZIP_WINDOW)) / ss));
    apply(plan);
  }
  // Both rules hold by construction. Checking them costs two small reads and turns a mistake into a failed write.
  if ((await lastEndRecord(plan)) >= 0) throw new Error('A ZIP end record is still within reach of ZIP readers');
  if ((await bytesAt(plan, 0, Math.min(1024, plan.size))).includes('%PDF-')) throw new Error('A PDF header is within the first 1024 bytes');

  const { fatSectors, difatSectors, dirStart, miniFatStart } = plan;
  const fatStart = 0;
  const difatStart = fatSectors;
  const fat = fatOf(plan);

  const header = new Bytes(ss);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(header.b, 0);
  header.u16(0x18, 0x3e);
  header.u16(0x1a, 3);
  header.u16(0x1c, 0xfffe);
  header.u16(0x1e, 9);
  header.u16(0x20, 6);
  header.u32(0x2c, fatSectors);
  header.u32(0x30, dirStart);
  header.u32(0x38, MINI_CUTOFF);
  header.u32(0x3c, miniFatStart);
  header.u32(0x40, miniFatSectors);
  header.u32(0x44, difatSectors ? difatStart : END_OF_CHAIN);
  header.u32(0x48, difatSectors);
  for (let i = 0; i < HEADER_FAT; i++) header.u32(0x4c + 4 * i, i < fatSectors ? fatStart + i : FREE);
  await sink.write(header.b);

  await sink.write(Buffer.from(fat.buffer));
  if (difatSectors) {
    const d = new Bytes(difatSectors * ss);
    const rest = Math.max(0, fatSectors - HEADER_FAT);
    for (let k = 0; k < difatSectors; k++) {
      for (let i = 0; i < per - 1; i++) {
        const n = k * (per - 1) + i;
        d.u32(k * ss + 4 * i, n < rest ? fatStart + HEADER_FAT + n : FREE);
      }
      d.u32(k * ss + 4 * (per - 1), k === difatSectors - 1 ? END_OF_CHAIN : difatStart + k + 1);
    }
    await sink.write(d.b);
  }

  /** Copies a stream's bytes and pads them to `unit`. */
  const copy = async (e: Entry, unit: number) => {
    const node = e.node;
    if (node?.kind !== 'stream') return;
    for (let at = 0; at < e.size; at += 262144) {
      checkTime(deadline);
      const want = Math.min(262144, e.size - at);
      const b = await node.source.read(at, want);
      if (b.length !== want) throw Object.assign(new Error(`A stream returned ${b.length} bytes where ${want} were expected`), { code: 'EIO' });
      await sink.write(b);
    }
    const pad = (unit - (e.size % unit)) % unit;
    if (pad) await sink.write(new Uint8Array(pad));
  };
  for (const e of small) await copy(e, MINI_SECTOR);
  const miniPad = miniStreamSectors * ss - miniStreamSize;
  if (miniPad) await sink.write(new Uint8Array(miniPad));
  for (const e of large) await copy(e, ss);
  await sink.write(directoryOf());
  if (miniFatSectors) await sink.write(Buffer.from(miniFat.buffer));
  for (let i = 0; i < plan.pad; i++) await sink.write(new Uint8Array(ss));
}
