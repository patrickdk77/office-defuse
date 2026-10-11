import { checkTime, StopError } from './errors';
import { Reader } from './io';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';

/** The first eight bytes of every compound file. */
export const CFB_SIGNATURE = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);

const MAX_REGULAR = 0xfffffffa;
const END_OF_CHAIN = 0xfffffffe;
const FREE = 0xffffffff;
const NO_STREAM = 0xffffffff;
const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
/** The header lists the first 109 FAT sectors itself. */
const HEADER_FAT = 109;

/** Owners of a sector besides the streams, which take 1 + their position among the entries. */
const OWN = { fat: -1, difat: -2, directory: -3, miniFat: -4, miniStream: -5 };

export interface CfbLimits {
  entries: number;
  decompressedBytes: number;
  totalDecompressedBytes: number;
}

export interface CfbEntry {
  /** The entry's position in the directory. */
  id: number;
  name: string;
  kind: 'root' | 'storage' | 'stream';
  /** The class id, as Windows writes a GUID, in upper case. */
  clsid: string;
  size: number;
  parent?: CfbEntry;
  children: CfbEntry[];
}

/** Something the reader saw that it does not refuse. Each is a strip finding, since a rewrite leaves the bytes out. */
export interface CfbNote {
  detail: D.HiddenData | D.TrailingBytes;
  bytes: number;
}

const fail = (detail: D, message: string, data?: Record<string, string | number>): never => {
  throw new StopError(C.Container, detail, message, data);
};
const limit = (detail: D, message: string, data?: Record<string, string | number>): never => {
  throw new StopError(C.Limit, detail, message, data);
};

const u16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;

/** A GUID as Windows shows it: the first three fields little-endian, upper case, without braces. */
export function guid(b: Uint8Array, at = 0): string {
  const h = (i: number) => b[at + i].toString(16).padStart(2, '0');
  const r = (...is: number[]) => is.map(h).join('');
  return `${r(3, 2, 1, 0)}-${r(5, 4)}-${r(7, 6)}-${r(8, 9)}-${r(10, 11, 12, 13, 14, 15)}`.toUpperCase();
}

/** An entry name with control and format characters written as \xNN, so a location can show it. */
export const showName = (name: string) => name.replace(/[\p{Cc}\p{Cf}]/gu, c => `\\x${c.charCodeAt(0).toString(16).padStart(2, '0')}`);

/** The entry's path from the root, its names joined by "/", which no name may hold. */
export function entryPath(e: CfbEntry): string {
  const names: string[] = [];
  for (let p: CfbEntry | undefined = e; p && p.kind !== 'root'; p = p.parent) names.push(showName(p.name));
  return names.reverse().join('/');
}

/** Names compare as MS-CFB compares them, ignoring case. Python's lower() and Windows' upper case both count. */
const nameKeys = (name: string) => [`u:${name.toUpperCase()}`, `l:${name.toLowerCase()}`];

/**
 * A compound file (MS-CFB), read from its header, its allocation tables and its directory. Every chain is followed as
 * far as its stream needs and every sector has one owner, so two readers cannot read the file differently. Reading a
 * stream maps its offsets through its chain; nothing is held but the tables.
 */
export class CompoundFile {
  private readonly streamChains = new Map<CfbEntry, Uint32Array>();

  private constructor(
    private readonly reader: Reader,
    readonly sectorSize: number,
    readonly root: CfbEntry,
    /** Every entry the directory tree reaches, the root first, each storage before its children. */
    readonly entries: CfbEntry[],
    readonly notes: CfbNote[],
    private readonly miniStream: Uint32Array,
    private readonly deadline: number | undefined,
  ) {}

  /** The child of a storage with this name, compared ignoring case. */
  child(storage: CfbEntry, name: string): CfbEntry | undefined {
    const key = name.toUpperCase();
    return storage.children.find(c => c.name.toUpperCase() === key);
  }

  /** Reads a compound file. The caller has seen its signature. */
  static async open(source: ByteSource, limits: CfbLimits, deadline?: number): Promise<CompoundFile> {
    const size = await source.size();
    const reader = new Reader(source, size);
    if (size < 512) fail(D.BadHeader, 'The file is too short for a compound file header');
    const h = await reader.read(0, 512);
    for (let i = 0; i < 8; i++) if (h[i] !== CFB_SIGNATURE[i]) fail(D.BadHeader, 'The file does not start with the compound file signature');
    const major = u16(h, 0x1a);
    const shift = u16(h, 0x1e);
    // olefile reads any sector size the header gives, and Windows only the two that match their version.
    if (!((major === 3 && shift === 9) || (major === 4 && shift === 12))) fail(D.BadHeader, `Version ${major} with sectors of 2^${shift} bytes`, { reason: 'version and sector size' });
    if (u16(h, 0x20) !== 6) fail(D.BadHeader, 'Mini sectors other than 64 bytes', { reason: 'mini sector size' });
    // olefile reads mini streams up to 4096 bytes whatever the header says, and other readers take the header's word.
    if (u32(h, 0x38) !== MINI_CUTOFF) fail(D.BadHeader, 'A mini stream cutoff other than 4096 bytes', { reason: 'mini stream cutoff' });
    if (u16(h, 0x1c) !== 0xfffe) fail(D.BadHeader, 'A byte order mark other than little-endian', { reason: 'byte order' });
    const ss = 2 ** shift;
    if (size < ss) fail(D.BadHeader, 'The file is too short for its header sector');
    const numFat = u32(h, 0x2c);
    const firstDir = u32(h, 0x30);
    const firstMiniFat = u32(h, 0x3c);
    const numMiniFat = u32(h, 0x40);
    const firstDifat = u32(h, 0x44);
    const numDifat = u32(h, 0x48);
    const notes: CfbNote[] = [];

    /** Sectors that start inside the file. The last may be cut short. */
    const fileSectors = Math.ceil((size - ss) / ss);
    const fullSectors = Math.floor((size - ss) / ss);
    const perSector = ss / 4;
    if (numFat === 0) fail(D.BadHeader, 'The header counts no FAT sectors', { reason: 'no FAT' });
    if (numFat > fullSectors) fail(D.OutOfRange, 'The header counts more FAT sectors than the file holds');

    // The FAT sectors: the first 109 from the header, the rest from the DIFAT chain.
    const fatSectors: number[] = [];
    const listed = Math.min(numFat, HEADER_FAT);
    for (let i = 0; i < listed; i++) fatSectors.push(u32(h, 0x4c + 4 * i));
    // olefile reads FAT sectors from the header until the first free or end mark, and Windows reads as many as the
    // header counts. A sector listed past the count would be read by one and not the other.
    if (numFat < HEADER_FAT) {
      const next = u32(h, 0x4c + 4 * numFat);
      if (next <= MAX_REGULAR) fail(D.BadHeader, 'The header lists more FAT sectors than it counts', { reason: 'FAT sector count' });
    }
    const difatSectors: number[] = [];
    if (numFat > HEADER_FAT) {
      const per = perSector - 1;
      const need = Math.ceil((numFat - HEADER_FAT) / per);
      if (numDifat !== need) fail(D.BadHeader, 'The DIFAT sector count does not fit the FAT sector count', { reason: 'DIFAT sector count' });
      let s = firstDifat;
      for (let k = 0; k < need; k++) {
        if (s > MAX_REGULAR) fail(D.BrokenChain, 'The DIFAT chain ends early');
        if (s >= fullSectors) fail(D.OutOfRange, 'A DIFAT sector lies outside the file');
        if (difatSectors.includes(s)) fail(D.ChainLoop, 'The DIFAT chain loops');
        difatSectors.push(s);
        const d = await reader.read((s + 1) * ss, ss);
        const take = Math.min(per, numFat - fatSectors.length);
        for (let i = 0; i < take; i++) fatSectors.push(u32(d, 4 * i));
        if (take < per && u32(d, 4 * take) <= MAX_REGULAR) fail(D.BadHeader, 'The DIFAT lists more FAT sectors than the header counts', { reason: 'FAT sector count' });
        s = u32(d, 4 * per);
      }
      if (s !== END_OF_CHAIN && s !== FREE) fail(D.BrokenChain, 'The DIFAT chain does not end after its last sector');
    } else if (numDifat !== 0) fail(D.BadHeader, 'The header counts DIFAT sectors it does not need', { reason: 'DIFAT sector count' });

    const fatLength = numFat * perSector;
    /** Sectors the FAT describes that start inside the file: the only ones a chain may name. */
    const sectors = Math.min(fatLength, fileSectors);
    if (fatLength < fileSectors) notes.push({ detail: D.TrailingBytes, bytes: size - (fatLength + 1) * ss });
    const owners = new Int32Array(sectors);
    const own = (s: number, owner: number, what: string) => {
      if (s >= sectors) fail(D.OutOfRange, `A sector of the ${what} lies outside the file`);
      if (owners[s] === owner) fail(D.ChainLoop, `The ${what} loops`);
      if (owners[s] !== 0) fail(D.Overlap, `The ${what} shares a sector with another stream`, { what });
      owners[s] = owner;
    };
    for (const s of fatSectors) {
      if (s > MAX_REGULAR) fail(D.BrokenChain, 'A FAT sector number is a reserved value');
      if (s >= fullSectors) fail(D.OutOfRange, 'A FAT sector lies outside the file');
      own(s, OWN.fat, 'FAT');
    }
    for (const s of difatSectors) own(s, OWN.difat, 'DIFAT');

    const fat = new Uint32Array(fatLength);
    for (const [i, s] of fatSectors.entries()) {
      if (i % 64 === 0) checkTime(deadline);
      const b = await reader.read((s + 1) * ss, ss);
      fat.set(new Uint32Array(b.buffer.slice(b.byteOffset, b.byteOffset + ss)), i * perSector);
    }

    /**
     * Follows a chain of the FAT, owning each sector: to its end mark, or for a stream only as far as its size needs,
     * since readers stop there. `full` sectors must lie wholly inside the file.
     */
    const chain = (start: number, owner: number, what: string, full: boolean, length = Number.POSITIVE_INFINITY): Uint32Array => {
      const out: number[] = [];
      for (let s = start; out.length < length; s = fat[s]) {
        if (out.length % 4096 === 0) checkTime(deadline);
        if (s === END_OF_CHAIN) {
          if (length === Number.POSITIVE_INFINITY) break;
          fail(D.DataError, `The ${what} is shorter than its declared size`, { reason: 'a chain shorter than the stream' });
        }
        if (s > MAX_REGULAR) fail(D.BrokenChain, `The ${what} runs into a free or reserved sector`, { what });
        own(s, owner, what);
        if (full && s >= fullSectors) fail(D.OutOfRange, `A sector of the ${what} is cut short by the end of the file`);
        out.push(s);
      }
      return Uint32Array.from(out);
    };

    const dir = chain(firstDir, OWN.directory, 'directory', true);
    if (dir.length === 0) fail(D.BadHeader, 'The directory is empty', { reason: 'no directory' });
    const perDirSector = ss / 128;
    const dirCount = dir.length * perDirSector;
    const rawEntry = (id: number) => reader.read((dir[Math.floor(id / perDirSector)] + 1) * ss + (id % perDirSector) * 128, 128);

    const v3 = major === 3;
    interface Raw {
      entry: CfbEntry;
      left: number;
      right: number;
      child: number;
      start: number;
    }
    const load = async (id: number, parent: CfbEntry | undefined): Promise<Raw> => {
      const b = await rawEntry(id);
      const type = b[0x42];
      const kind = type === 5 ? 'root' : type === 1 ? 'storage' : type === 2 ? 'stream' : undefined;
      if (kind === undefined) return fail(D.BadEntry, `Directory entry ${id} has type ${type}`, { reason: `entry type ${type}` });
      if ((kind === 'root') !== (id === 0)) fail(D.BadEntry, 'The root entry is not the first entry of the directory, or there are two', { reason: 'root entry' });
      const nameLength = u16(b, 0x40);
      let name = '';
      if (kind !== 'root') {
        // olefile and LibreOffice take the name from its length field. An odd or empty length, or one past the field,
        // makes them read different names, and a NUL inside one ends it for a reader that takes it as a C string.
        if (nameLength < 2 || nameLength > 64 || nameLength % 2) fail(D.BadEntry, `Directory entry ${id} has a name length of ${nameLength}`, { reason: 'name length' });
        name = Buffer.from(b.subarray(0, nameLength - 2)).toString('utf16le');
        if (u16(b, nameLength - 2) !== 0) fail(D.BadEntry, `The name of directory entry ${id} is not terminated where its length says`, { reason: 'name terminator' });
        if (name.includes('\0')) fail(D.BadEntry, `The name of directory entry ${id} holds a NUL`, { reason: 'NUL in name' });
        if (/[/\\:!]/.test(name)) fail(D.UnsafeName, `The name of directory entry ${id} holds a character MS-CFB forbids`, { name: showName(name), reason: 'a character MS-CFB forbids' });
      }
      const low = u32(b, 0x78);
      const high = u32(b, 0x7c);
      // A version 3 file keeps its sizes in 32 bits. Readers ignore the high half, which some writers leave set.
      let entrySize = low;
      if (!v3) {
        const big = high * 2 ** 32 + low;
        if (big > Number.MAX_SAFE_INTEGER) fail(D.OutOfRange, 'A stream size is beyond any file this reader can hold');
        entrySize = big;
      }
      return {
        entry: { id, name, kind, clsid: guid(b, 0x50), size: kind === 'storage' ? 0 : entrySize, parent, children: [] },
        left: u32(b, 0x44),
        right: u32(b, 0x48),
        child: u32(b, 0x4c),
        start: u32(b, 0x74),
      };
    };

    const rootRaw = await load(0, undefined);
    const root = rootRaw.entry;
    const entries: CfbEntry[] = [root];
    /** The first sector of each entry, by its position in `entries`. */
    const starts: number[] = [rootRaw.start];
    const reached = new Uint8Array(dirCount);
    reached[0] = 1;
    const storages: Raw[] = [rootRaw];
    while (storages.length) {
      checkTime(deadline);
      const storage = storages.pop() as Raw;
      const seen = new Set<string>();
      // An in-order walk of the storage's tree of children.
      const stack: Raw[] = [];
      let at = storage.child;
      while (at !== NO_STREAM || stack.length) {
        while (at !== NO_STREAM) {
          if (at >= dirCount) fail(D.OutOfRange, 'A directory entry names an entry past the directory');
          if (reached[at]) fail(D.DirectoryLoop, `Directory entry ${at} is reached twice`);
          reached[at] = 1;
          const r = await load(at, storage.entry);
          if (entries.length >= limits.entries) limit(D.EntryCount, `More than ${limits.entries} directory entries`, { limit: limits.entries });
          stack.push(r);
          at = r.left;
        }
        const r = stack.pop() as Raw;
        for (const k of nameKeys(r.entry.name)) {
          if (seen.has(k)) fail(D.DuplicateName, `Two entries of one storage are named ${showName(r.entry.name)}`, { name: showName(r.entry.name) });
          seen.add(k);
        }
        storage.entry.children.push(r.entry);
        entries.push(r.entry);
        starts.push(r.start);
        if (r.entry.kind === 'storage') storages.push(r);
        else if (r.child !== NO_STREAM) fail(D.BadEntry, `Stream ${showName(r.entry.name)} has children`, { reason: 'a stream with children' });
        at = r.right;
      }
    }

    let total = 0;
    const big: Array<[CfbEntry, Uint32Array]> = [];
    /** File offsets and lengths of the bytes past a stream's end in its last sector. */
    const slack: Array<[number, number]> = [];
    const small: Array<[CfbEntry, number]> = [];
    for (const [i, e] of entries.entries()) {
      if (e.kind === 'storage' || e.size === 0) continue;
      if (e.size > limits.decompressedBytes)
        limit(D.DecompressedSize, `${entryPath(e) || 'The mini stream'} is ${e.size} bytes`, { name: entryPath(e), size: e.size, limit: limits.decompressedBytes });
      total += e.size;
      if (total > limits.totalDecompressedBytes) limit(D.TotalDecompressedSize, `The streams hold more than ${limits.totalDecompressedBytes} bytes`, { limit: limits.totalDecompressedBytes });
      if (e.kind === 'stream' && e.size < MINI_CUTOFF) {
        small.push([e, i]);
        continue;
      }
      const what = e.kind === 'root' ? 'mini stream' : `stream ${showName(e.name)}`;
      const need = Math.ceil(e.size / ss);
      const c = chain(starts[i], e.kind === 'root' ? OWN.miniStream : i + 1, what, false, need);
      for (let k = 0; k < need - 1; k++) if (c[k] >= fullSectors) fail(D.OutOfRange, `The ${what} is cut short by the end of the file`);
      const tail = e.size - (need - 1) * ss;
      if ((c[need - 1] + 1) * ss + tail > size) fail(D.OutOfRange, `The ${what} is cut short by the end of the file`);
      big.push([e, c]);
      // The rest of the last sector belongs to no stream.
      if (tail < ss) slack.push([(c[need - 1] + 1) * ss + tail, Math.min(ss - tail, size - ((c[need - 1] + 1) * ss + tail))]);
    }
    const cf = new CompoundFile(reader, ss, root, entries, notes, big.find(([e]) => e.kind === 'root')?.[1] ?? new Uint32Array(0), deadline);
    for (const [e, c] of big) if (e.kind !== 'root') cf.streamChains.set(e, c);

    // The mini FAT, and the mini streams in the mini stream. olefile reads as many mini FAT sectors as the header
    // counts, and other readers follow the chain, so both must cover every mini sector a stream can name.
    const miniFatChain = firstMiniFat === END_OF_CHAIN ? new Uint32Array(0) : chain(firstMiniFat, OWN.miniFat, 'mini FAT', true);
    const miniSectors = Math.ceil(root.size / MINI_SECTOR);
    const miniOwners = new Int32Array(miniSectors);
    /** Mini stream offsets and lengths of the bytes past a mini stream's end in its last mini sector. */
    const miniSlack: Array<[number, number]> = [];
    if (small.length) {
      const needFat = Math.ceil(miniSectors / perSector);
      if (miniFatChain.length < needFat || numMiniFat < needFat) fail(D.BadHeader, 'The mini FAT is too short for the mini stream', { reason: 'mini FAT sector count' });
      const miniFat = new Uint32Array(miniFatChain.length * perSector);
      for (const [i, s] of miniFatChain.entries()) {
        const b = await reader.read((s + 1) * ss, ss);
        miniFat.set(new Uint32Array(b.buffer.slice(b.byteOffset, b.byteOffset + ss)), i * perSector);
      }
      for (const [e, i] of small) {
        const what = `stream ${showName(e.name)}`;
        const out: number[] = [];
        const owner = i + 1;
        const need = Math.ceil(e.size / MINI_SECTOR);
        for (let s = starts[i]; out.length < need; s = miniFat[s]) {
          if (s === END_OF_CHAIN) fail(D.DataError, `The ${what} is shorter than its declared size`, { name: entryPath(e), reason: 'a chain shorter than the stream' });
          if (s > MAX_REGULAR) fail(D.BrokenChain, `The ${what} runs into a free or reserved mini sector`, { what });
          if (s >= miniSectors) fail(D.OutOfRange, `A mini sector of the ${what} lies outside the mini stream`);
          if (miniOwners[s] === owner) fail(D.ChainLoop, `The ${what} loops`);
          if (miniOwners[s] !== 0) fail(D.Overlap, `The ${what} shares a mini sector with another stream`, { what });
          miniOwners[s] = owner;
          out.push(s);
        }
        cf.streamChains.set(e, Uint32Array.from(out));
        const tail = e.size - (need - 1) * MINI_SECTOR;
        if (tail < MINI_SECTOR) miniSlack.push([out[need - 1] * MINI_SECTOR + tail, Math.min(MINI_SECTOR - tail, root.size - (out[need - 1] * MINI_SECTOR + tail))]);
      }
    }

    // Bytes no stream covers: sectors nothing owns, mini sectors no stream owns, the rest of each stream's last
    // sector, and directory entries the tree does not reach. Zeros are padding and count for nothing; anything else
    // could carry content that a reader of the raw file finds and a rewrite leaves out.
    let hidden = 0;
    for (let s = 0; s < sectors; s++) {
      if (owners[s] !== 0) continue;
      checkTime(deadline);
      hidden += content(await reader.read((s + 1) * ss, Math.min(ss, size - (s + 1) * ss)));
    }
    if (major === 4) hidden += content((await reader.read(0, ss)).subarray(512));
    for (const [at, n] of slack) hidden += content(await reader.read(at, n));
    for (let m = 0; m < miniSectors; m++) {
      if (miniOwners[m] !== 0) continue;
      hidden += content(await cf.readMini(m * MINI_SECTOR, Math.min(MINI_SECTOR, root.size - m * MINI_SECTOR)));
    }
    for (const [at, n] of miniSlack) hidden += content(await cf.readMini(at, n));
    for (let id = 0; id < dirCount; id++) if (!reached[id]) hidden += unusedEntry(await rawEntry(id));
    if (hidden > 0) notes.push({ detail: D.HiddenData, bytes: hidden });
    return cf;
  }

  /** Bytes of the mini stream, by their offset in it. */
  private readMini(offset: number, length: number): Promise<Uint8Array> {
    return this.readChain(this.miniStream, this.sectorSize, this.root.size, offset, length, (s, at, n) => this.reader.read((s + 1) * this.sectorSize + at, n));
  }

  private async readChain(chain: Uint32Array, unit: number, size: number, offset: number, length: number, read: (sector: number, at: number, n: number) => Promise<Uint8Array>): Promise<Uint8Array> {
    const end = Math.min(size, offset + length);
    if (end <= offset) return new Uint8Array(0);
    const out = new Uint8Array(end - offset);
    let pos = offset;
    while (pos < end) {
      const k = Math.floor(pos / unit);
      const at = pos % unit;
      const n = Math.min(unit - at, end - pos);
      out.set(await read(chain[k], at, n), pos - offset);
      pos += n;
    }
    return out;
  }

  /** A random-access source over a stream's content. */
  source(e: CfbEntry): ByteSource {
    const chain = this.streamChains.get(e) ?? new Uint32Array(0);
    const mini = e.size < MINI_CUTOFF;
    const read = mini ? (s: number, at: number, n: number) => this.readMini(s * MINI_SECTOR + at, n) : (s: number, at: number, n: number) => this.reader.read((s + 1) * this.sectorSize + at, n);
    const unit = mini ? MINI_SECTOR : this.sectorSize;
    return {
      size: async () => e.size,
      read: (offset, length) => {
        checkTime(this.deadline);
        return this.readChain(chain, unit, e.size, offset, length, read);
      },
    };
  }

  /** A stream's whole content. Callers read only streams whose size they bounded. */
  bytes(e: CfbEntry): Promise<Uint8Array> {
    return this.source(e).read(0, e.size);
  }
}

/**
 * The length of a directory entry the tree does not reach, or 0 when it is unused as MS-CFB writes one: all zeros,
 * with its sibling and child fields 0xFFFFFFFF or zero.
 */
function unusedEntry(b: Uint8Array): number {
  const links = b.subarray(0x44, 0x50);
  const blank = links.every(x => x === 0xff) || links.every(x => x === 0);
  return blank && content(b.subarray(0, 0x44)) === 0 && content(b.subarray(0x50)) === 0 ? 0 : b.length;
}

/** The length of `b` when it holds anything but zeros, and 0 when it does not. */
const content = (b: Uint8Array) => (b.some(x => x !== 0) ? b.length : 0);
