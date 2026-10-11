import * as stream from 'node:stream';
import * as zlib from 'node:zlib';
import { crc32 } from './crc32';
import { checkTime, StopError } from './errors';
import { Reader } from './io';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_END = 0x06054b50;
const SIG_END64 = 0x06064b50;
const SIG_LOCATOR64 = 0x07064b50;
const SIG_DESCRIPTOR = 0x08074b50;
const MAX16 = 0xffff;
const MAX32 = 0xffffffff;
const UNKNOWN64 = 0xffffffffffffffffn;
/** The end record is 22 bytes and its comment at most 65,535. */
const END_SEARCH = 22 + MAX16;
/** Flags whose meaning a local header and the central directory must agree on: encryption, data descriptor, strong encryption, UTF-8 names and an encrypted directory. */
const FLAG_MASK = 0x0001 | 0x0008 | 0x0040 | 0x0800 | 0x2000;
const ENCRYPTION_FLAGS = 0x0001 | 0x0040 | 0x2000;
const MAX_NAME_BYTES = 1024;

export interface ZipLimits {
  entries: number;
  decompressedBytes: number;
  totalDecompressedBytes: number;
}

export interface ZipEntry {
  /** Position in the central directory. */
  index: number;
  name: string;
  /** A directory entry, whose name ends in "/". */
  directory: boolean;
  flags: number;
  /** 0 stored, 8 deflate. */
  method: number;
  dosTime: number;
  dosDate: number;
  crc: number;
  compressedSize: number;
  size: number;
  localOffset: number;
  dataOffset: number;
  /** Where the entry ends, after any data descriptor. */
  end: number;
  /** The local header has an extra field. ODF wants none on the mimetype entry. */
  localExtra: boolean;
}

/** Something the reader saw that it does not refuse. Each is a strip finding, since the rewrite drops the bytes. */
export interface ZipNote {
  detail: D.LeadingBytes | D.HiddenData | D.HiddenEntry | D.TrailingBytes;
  bytes: number;
  /** For a hidden entry, where its local header sits. */
  offset?: number;
}

const LOCAL_SIGNATURE = Buffer.from([0x50, 0x4b, 0x03, 0x04]);

/** The offset of the first local file header signature in [start, end) of the file, or undefined. */
async function findLocalHeader(reader: Reader, start: number, end: number, deadline: number | undefined): Promise<number | undefined> {
  let carry = Buffer.alloc(0);
  let at = start;
  for await (const c of reader.chunks(start, end - start)) {
    checkTime(deadline);
    const b = Buffer.concat([carry, Buffer.from(c.buffer, c.byteOffset, c.byteLength)]);
    const i = b.indexOf(LOCAL_SIGNATURE);
    if (i >= 0) return at - carry.length + i;
    carry = b.subarray(Math.max(0, b.length - 3));
    at += c.length;
  }
  return undefined;
}

const fail = (detail: D, message: string, data?: Record<string, string | number>): never => {
  throw new StopError(C.Container, detail, message, data);
};
const limit = (detail: D, message: string, data?: Record<string, string | number>): never => {
  throw new StopError(C.Limit, detail, message, data);
};

class Bytes {
  private readonly view: DataView;
  constructor(readonly b: Uint8Array) {
    this.view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  }
  u16(at: number): number {
    return this.view.getUint16(at, true);
  }
  u32(at: number): number {
    return this.view.getUint32(at, true);
  }
  /** An 8-byte value as it is, for comparing bytes that may not be a 64-bit field at all. */
  big(at: number): bigint {
    return this.view.getBigUint64(at, true);
  }
  /** An 8-byte value, which must fit a JavaScript number exactly. */
  u64(at: number): number {
    const v = this.big(at);
    if (v > BigInt(Number.MAX_SAFE_INTEGER)) fail(D.OutOfRange, 'A 64-bit field is beyond any file this reader can hold');
    return Number(v);
  }
}

/** Splits an extra field block into records. The block must split exactly, and no record id may repeat. */
function parseExtra(extra: Uint8Array, name: string): Map<number, Uint8Array> {
  const out = new Map<number, Uint8Array>();
  const b = new Bytes(extra);
  let p = 0;
  while (p < extra.length) {
    if (p + 4 > extra.length) fail(D.BadExtraField, 'An extra field block ends inside a record header', { name });
    const id = b.u16(p);
    const len = b.u16(p + 2);
    if (p + 4 + len > extra.length) fail(D.BadExtraField, 'An extra field record runs past its block', { name });
    if (out.has(id)) fail(D.BadExtraField, `Extra field 0x${id.toString(16)} appears twice`, { name });
    out.set(id, extra.subarray(p + 4, p + 4 + len));
    p += 4 + len;
  }
  return out;
}

const utf8 = new TextDecoder('utf-8', { fatal: true });

/** Decodes a name the way the flags say. A non-ASCII name without the UTF-8 flag has no single reading. */
function decodeName(raw: Uint8Array, flags: number): string {
  if (flags & 0x0800) {
    try {
      return utf8.decode(raw);
    } catch {
      return fail(D.NameEncoding, 'An entry name flagged UTF-8 is not valid UTF-8');
    }
  }
  for (const c of raw) if (c >= 0x80) fail(D.NameEncoding, 'An entry name has non-ASCII bytes without the UTF-8 flag');
  return Buffer.from(raw).toString('latin1');
}

/** The Info-ZIP Unicode path field gives a second name, which some readers prefer. It must name the same entry. */
function checkUnicodePath(extras: Map<number, Uint8Array>, name: string): void {
  const up = extras.get(0x7075);
  if (!up) return;
  let other: string | undefined;
  try {
    other = up.length >= 5 && up[0] === 1 ? utf8.decode(up.subarray(5)) : undefined;
  } catch {
    other = undefined;
  }
  if (other !== name) fail(D.NameEncoding, 'The Unicode path field names a different entry', { name });
}

/** True for a segment of nothing but dots and spaces, the empty one included. */
function onlyDotsAndSpaces(seg: string): boolean {
  for (const c of seg) if (c !== '.' && c !== ' ') return false;
  return true;
}

/** A segment without the dots and spaces Windows drops from its end. A loop, since /[. ]+$/ backtracks. */
function trimSegment(seg: string): string {
  let end = seg.length;
  while (end > 0 && (seg[end - 1] === '.' || seg[end - 1] === ' ')) end--;
  return seg.slice(0, end);
}

function checkName(name: string, directory: boolean): void {
  const why = (reason: string) => fail(D.UnsafeName, `Entry name ${reason}`, { name, reason });
  if (name === '') why('is empty');
  if (/\p{Cc}/u.test(name)) why('holds a control character');
  if (name.includes('\\')) why('holds a backslash');
  if (name.startsWith('/')) why('is absolute');
  if (/^[A-Za-z]:/.test(name)) why('starts with a drive letter');
  for (const seg of (directory ? name.slice(0, -1) : name).split('/')) {
    // Windows drops trailing dots and spaces from each segment, so ".. " and "..." climb out as ".." does.
    if (onlyDotsAndSpaces(seg)) why('has an empty segment, or one of only dots and spaces');
  }
}

/**
 * The key two names collide on: percent-decoded, NFKC and case-folded, without a directory's slash, and without the
 * dots and spaces Windows drops from the end of each segment. Upper then lower case is the full folding, which takes a
 * sharp s to ss as macOS's case-insensitive file system does. NFKC also joins compatibility forms such as ligatures
 * and full-width letters.
 */
export function nameKey(name: string): string {
  let n = name;
  try {
    n = decodeURIComponent(n);
  } catch {
    // Not valid percent-encoding, so it stays as written.
  }
  return n.normalize('NFKC').toUpperCase().toLowerCase().replace(/\/$/, '').split('/').map(trimSegment).join('/');
}

/** ZIP64 values the central directory defers to its extra field, in the order the format fixes. */
function zip64Central(
  extras: Map<number, Uint8Array>,
  fields: { size: number; compressedSize: number; offset: number; disk: number },
  name: string,
): { size: number; compressedSize: number; offset: number; disk: number } {
  const want = [fields.size === MAX32, fields.compressedSize === MAX32, fields.offset === MAX32, fields.disk === MAX16];
  const length = (want[0] ? 8 : 0) + (want[1] ? 8 : 0) + (want[2] ? 8 : 0) + (want[3] ? 4 : 0);
  const z = extras.get(0x0001);
  if (!z) {
    if (length) fail(D.Zip64Mismatch, 'A field at its maximum has no ZIP64 value behind it', { name });
    return fields;
  }
  if (z.length !== length) fail(D.Zip64Mismatch, 'A ZIP64 extra field holds other fields than its header defers to it', { name });
  const b = new Bytes(z);
  let p = 0;
  const out = { ...fields };
  if (want[0]) {
    out.size = b.u64(p);
    p += 8;
  }
  if (want[1]) {
    out.compressedSize = b.u64(p);
    p += 8;
  }
  if (want[2]) {
    out.offset = b.u64(p);
    p += 8;
  }
  if (want[3]) out.disk = b.u32(p);
  return out;
}

/**
 * A ZIP archive read from its central directory, after every local header was checked against it. Reading an entry
 * inflates it in chunks and checks its size and CRC.
 */
export class ZipArchive {
  private readonly byKey = new Map<string, ZipEntry>();

  private constructor(
    private readonly reader: Reader,
    readonly entries: ZipEntry[],
    readonly notes: ZipNote[],
    private readonly deadline: number | undefined,
  ) {
    for (const e of entries) this.byKey.set(nameKey(e.name), e);
  }

  /** The entry with this name, compared as entry names collide: ignoring case. */
  get(name: string): ZipEntry | undefined {
    return this.byKey.get(nameKey(name));
  }

  static async open(source: ByteSource, limits: ZipLimits, deadline?: number): Promise<ZipArchive> {
    const size = await source.size();
    const reader = new Reader(source, size);
    const notes: ZipNote[] = [];
    if (size < 22) fail(D.NotZip, 'The file is too short to be a ZIP archive');
    const tailStart = Math.max(0, size - END_SEARCH);
    const tail = new Bytes(await reader.read(tailStart, size - tailStart));
    // The last end signature is the one a reader scanning backwards takes. An earlier one whose comment reaches the
    // end of the file holds it inside that comment, and a reader that checks comment lengths can take either.
    let endAt = -1;
    let holder = false;
    for (let p = 0; p + 22 <= tail.b.length; p++) {
      if (tail.b[p] !== 0x50 || tail.u32(p) !== SIG_END) continue;
      if (endAt >= 0 && tailStart + endAt + 22 + tail.u16(endAt + 20) === size) holder = true;
      endAt = p;
    }
    if (endAt < 0) fail(D.NotZip, 'No end of central directory record');
    if (holder) fail(D.AmbiguousEnd, "A second end record sits inside an end record's comment");
    const end = tailStart + endAt;
    const commentLength = tail.u16(endAt + 20);
    if (end + 22 + commentLength > size) fail(D.NotZip, "The end record's comment runs past the end of the file");
    if (end + 22 + commentLength < size) notes.push({ detail: D.TrailingBytes, bytes: size - (end + 22 + commentLength) });
    let disk = tail.u16(endAt + 4);
    let dirDisk = tail.u16(endAt + 6);
    let countOnDisk = tail.u16(endAt + 8);
    let count = tail.u16(endAt + 10);
    let dirSize = tail.u32(endAt + 12);
    let dirOffset = tail.u32(endAt + 16);
    let dirEnd = end;
    const locatorAt = end - 20;
    const locator = locatorAt >= 0 ? new Bytes(await reader.read(locatorAt, 20)) : undefined;
    if (locator && locator.u32(0) === SIG_LOCATOR64) {
      if (locator.u32(4) !== 0 || locator.u32(16) > 1) fail(D.MultiDisk, 'The ZIP64 locator names another disk');
      const at = locator.u64(8);
      if (at + 56 > locatorAt) fail(D.Zip64Mismatch, 'The ZIP64 locator does not point at a ZIP64 end record');
      const r = new Bytes(await reader.read(at, 56));
      if (r.u32(0) !== SIG_END64) fail(D.Zip64Mismatch, 'The ZIP64 locator does not point at a ZIP64 end record');
      if (at + 12 + r.u64(4) !== locatorAt) fail(D.Zip64Mismatch, 'The ZIP64 end record does not end at its locator');
      const pick = (short: number, max: number, long: number, what: string) => {
        if (short === max) return long;
        if (short !== long) fail(D.Zip64Mismatch, `The end record and the ZIP64 end record disagree on ${what}`);
        return long;
      };
      disk = pick(disk, MAX16, r.u32(16), 'the disk number');
      dirDisk = pick(dirDisk, MAX16, r.u32(20), 'the directory disk');
      countOnDisk = pick(countOnDisk, MAX16, r.u64(24), 'the entry count');
      count = pick(count, MAX16, r.u64(32), 'the entry count');
      dirSize = pick(dirSize, MAX32, r.u64(40), 'the directory size');
      dirOffset = pick(dirOffset, MAX32, r.u64(48), 'the directory offset');
      dirEnd = at;
    } else if (disk === MAX16 || dirDisk === MAX16 || countOnDisk === MAX16 || count === MAX16 || dirSize === MAX32 || dirOffset === MAX32) {
      fail(D.Zip64Mismatch, 'The end record defers to a ZIP64 record that is not there');
    }
    if (disk !== 0 || dirDisk !== 0) fail(D.MultiDisk, 'The archive spans several disks');
    if (countOnDisk !== count) fail(D.DirectoryMismatch, 'The entry counts for this disk and for the archive differ');
    if (count > limits.entries) limit(D.EntryCount, `${count} entries`, { entries: count, limit: limits.entries });
    if (dirOffset > dirEnd || dirSize > dirEnd) fail(D.OutOfRange, 'The central directory lies outside the file');
    if (dirOffset + dirSize !== dirEnd) fail(D.DirectoryMismatch, 'The central directory does not end where the end record begins');

    const entries: ZipEntry[] = [];
    const seen = new Map<string, string>();
    const raws: Uint8Array[] = [];
    let total = 0;
    let pos = dirOffset;
    for (let index = 0; index < count; index++) {
      if (index % 256 === 0) checkTime(deadline);
      if (pos + 46 > dirEnd) fail(D.DirectoryMismatch, 'The central directory holds fewer entries than the end record says');
      const h = new Bytes(await reader.read(pos, 46));
      if (h.u32(0) !== SIG_CENTRAL) fail(D.DirectoryMismatch, 'The central directory holds fewer entries than the end record says');
      const flags = h.u16(8);
      const method = h.u16(10);
      const nameLength = h.u16(28);
      const extraLength = h.u16(30);
      const commentLen = h.u16(32);
      if (pos + 46 + nameLength + extraLength + commentLen > dirEnd) fail(D.DirectoryMismatch, 'A central directory header runs past the directory');
      if (nameLength > MAX_NAME_BYTES) fail(D.UnsafeName, `An entry name is ${nameLength} bytes long`, { reason: `longer than ${MAX_NAME_BYTES} bytes` });
      const tailBytes = await reader.read(pos + 46, nameLength + extraLength);
      const raw = tailBytes.slice(0, nameLength);
      const name = decodeName(raw, flags);
      const extras = parseExtra(tailBytes.subarray(nameLength), name);
      checkUnicodePath(extras, name);
      if (flags & ENCRYPTION_FLAGS) fail(D.EncryptedEntry, 'An entry is encrypted', { name });
      if (method !== 0 && method !== 8) fail(D.UnsupportedMethod, `Compression method ${method}`, { name, method });
      const z = zip64Central(extras, { size: h.u32(24), compressedSize: h.u32(20), offset: h.u32(42), disk: h.u16(34) }, name);
      if (z.disk !== 0) fail(D.MultiDisk, 'An entry starts on another disk', { name });
      const directory = name.endsWith('/');
      checkName(name, directory);
      if (directory && z.size !== 0) fail(D.UnsafeName, 'A directory entry has content', { name, reason: 'a directory entry with content' });
      if (method === 0 && z.size !== z.compressedSize) fail(D.DataError, 'A stored entry has two different sizes', { name });
      if (z.size > limits.decompressedBytes) limit(D.DecompressedSize, `${name} declares ${z.size} bytes`, { name, size: z.size, limit: limits.decompressedBytes });
      total += z.size;
      if (total > limits.totalDecompressedBytes) limit(D.TotalDecompressedSize, `The entries declare more than ${limits.totalDecompressedBytes} bytes`, { limit: limits.totalDecompressedBytes });
      const key = nameKey(name);
      const other = seen.get(key);
      if (other !== undefined) fail(D.DuplicateName, `${other} and ${name}`, { name, other });
      seen.set(key, name);
      entries.push({
        index,
        name,
        directory,
        flags,
        method,
        dosTime: h.u16(12),
        dosDate: h.u16(14),
        crc: h.u32(16),
        compressedSize: z.compressedSize,
        size: z.size,
        localOffset: z.offset,
        dataOffset: 0,
        end: 0,
        localExtra: false,
      });
      raws.push(raw);
      pos += 46 + nameLength + extraLength + commentLen;
    }
    if (pos !== dirEnd) fail(D.DirectoryMismatch, 'The central directory holds more than the end record counts');

    for (const [i, e] of entries.entries()) {
      if (i % 256 === 0) checkTime(deadline);
      await ZipArchive.checkLocal(reader, e, raws[i], dirOffset);
    }
    // A reader that walks the file cannot know where a stored entry with a data descriptor ends. It looks for the
    // next signature, so a local header inside the data is an entry to it, and part of this entry here.
    for (const e of entries) {
      if (e.method !== 0 || !(e.flags & 0x0008) || e.compressedSize === 0) continue;
      const at = await findLocalHeader(reader, e.dataOffset, e.dataOffset + e.compressedSize, deadline);
      if (at !== undefined) notes.push({ detail: D.HiddenEntry, bytes: e.dataOffset + e.compressedSize - at, offset: at });
    }

    // Entries and the central directory must tile the file. Bytes no entry covers are reported, overlaps refused.
    const sorted = [...entries].sort((a, b) => a.localOffset - b.localOffset);
    const gaps: Array<[number, number]> = [];
    let covered = 0;
    let hidden = 0;
    for (const e of sorted) {
      if (e.localOffset < covered) fail(D.Overlap, 'Two entries share bytes', { name: e.name });
      if (e.localOffset > covered) gaps.push([covered, e.localOffset]);
      if (covered === 0 && e.localOffset > 0) notes.push({ detail: D.LeadingBytes, bytes: e.localOffset });
      else hidden += e.localOffset - covered;
      covered = e.end;
    }
    if (dirOffset > covered) gaps.push([covered, dirOffset]);
    if (sorted.length === 0 && dirOffset > 0) notes.push({ detail: D.LeadingBytes, bytes: dirOffset });
    else hidden += dirOffset - covered;
    if (hidden > 0) notes.push({ detail: D.HiddenData, bytes: hidden });
    // A local header in a gap is an entry the central directory does not list. A reader that walks the file from the
    // start, as streaming unzippers do, extracts it; a reader of the central directory never sees it.
    for (const [from, to] of gaps) {
      const at = await findLocalHeader(reader, from, to, deadline);
      if (at !== undefined) notes.push({ detail: D.HiddenEntry, bytes: to - from, offset: at });
    }
    return new ZipArchive(reader, entries, notes, deadline);
  }

  /** Checks a local header and any data descriptor against the central directory, and fills in where the entry's data and the entry end. */
  private static async checkLocal(reader: Reader, e: ZipEntry, rawName: Uint8Array, dirOffset: number): Promise<void> {
    const name = e.name;
    const mismatch = (what: string) => fail(D.HeaderMismatch, `The local header ${what}`, { name, what });
    if (e.localOffset + 30 > dirOffset) fail(D.OutOfRange, 'A local header lies outside the entries', { name });
    const h = new Bytes(await reader.read(e.localOffset, 30));
    if (h.u32(0) !== SIG_LOCAL) mismatch('is missing');
    const flags = h.u16(6);
    const nameLength = h.u16(26);
    const extraLength = h.u16(28);
    e.dataOffset = e.localOffset + 30 + nameLength + extraLength;
    if (e.dataOffset > dirOffset) fail(D.OutOfRange, 'A local header runs into the central directory', { name });
    const rest = await reader.read(e.localOffset + 30, nameLength + extraLength);
    if (Buffer.compare(rest.subarray(0, nameLength), rawName) !== 0) mismatch('names another entry');
    if ((flags & FLAG_MASK) !== (e.flags & FLAG_MASK)) mismatch('has other flags');
    if (h.u16(8) !== e.method) mismatch('has another compression method');
    const extras = parseExtra(rest.subarray(nameLength), name);
    checkUnicodePath(extras, name);
    e.localExtra = extraLength > 0;
    let size = h.u32(22);
    let compressedSize = h.u32(18);
    const z = extras.get(0x0001);
    if (z) {
      // The local ZIP64 field holds both sizes, and may go on to the offset and disk.
      if (z.length !== 16 && z.length !== 24 && z.length !== 28) fail(D.Zip64Mismatch, 'A local ZIP64 extra field has an odd length', { name });
      const b = new Bytes(z);
      // A writer that streams the entry does not know its sizes yet. It sets the data descriptor flag and may fill
      // the ZIP64 sizes with ones, which then mean what zero means: the data descriptor has them.
      const wide = (at: number) => (e.flags & 0x0008 && b.big(at) === UNKNOWN64 ? 0 : b.u64(at));
      if (size !== MAX32 && size !== 0 && size !== wide(0)) fail(D.Zip64Mismatch, 'A local header and its ZIP64 field disagree on the size', { name });
      if (compressedSize !== MAX32 && compressedSize !== 0 && compressedSize !== wide(8)) fail(D.Zip64Mismatch, 'A local header and its ZIP64 field disagree on the size', { name });
      size = wide(0);
      compressedSize = wide(8);
    } else if (size === MAX32 || compressedSize === MAX32) {
      fail(D.Zip64Mismatch, 'A local size at its maximum has no ZIP64 value behind it', { name });
    }
    const crc = h.u32(14);
    if (e.flags & 0x0008) {
      // With a data descriptor the local fields may be zero. Any that are set must agree.
      if (crc !== 0 && crc !== e.crc) mismatch('has another CRC');
      if (size !== 0 && size !== e.size) mismatch('has another size');
      if (compressedSize !== 0 && compressedSize !== e.compressedSize) mismatch('has another compressed size');
    } else if (crc !== e.crc || size !== e.size || compressedSize !== e.compressedSize) {
      mismatch('has another CRC or size');
    }
    const dataEnd = e.dataOffset + e.compressedSize;
    if (dataEnd > dirOffset) fail(D.Overlap, "An entry's data runs into the central directory", { name });
    e.end = dataEnd;
    if (!(e.flags & 0x0008)) return;
    const d = new Bytes(await reader.read(dataEnd, Math.min(24, dirOffset - dataEnd)));
    let p = 0;
    if (d.b.length >= 8 && d.u32(0) === SIG_DESCRIPTOR && d.u32(4) === e.crc) p = 4;
    if (d.b.length < p + 12 || d.u32(p) !== e.crc) mismatch('has a data descriptor with another CRC');
    // The format says eight-byte sizes go with a ZIP64 field. Either width is accepted when its values agree; the
    // tiling check then sees any bytes the other reading would leave over. Past a narrow descriptor the wide reading
    // takes in the next header's signature, so it compares the bytes without demanding that they fit a number.
    const wide = d.b.length >= p + 20 && d.big(p + 4) === BigInt(e.compressedSize) && d.big(p + 12) === BigInt(e.size);
    const narrow = d.u32(p + 4) === e.compressedSize && d.u32(p + 8) === e.size;
    if (z ? wide : narrow) e.end = dataEnd + p + (z ? 20 : 12);
    else if (z ? narrow : wide) e.end = dataEnd + p + (z ? 12 : 20);
    else mismatch('has a data descriptor with other sizes');
  }

  /** The compressed bytes of an entry, as stored. */
  raw(e: ZipEntry): AsyncGenerator<Uint8Array> {
    return this.reader.chunks(e.dataOffset, e.compressedSize);
  }

  /** The entry's content in chunks. Its size, its CRC and, for deflate, the compressed bytes consumed must all agree with the central directory. */
  async *read(e: ZipEntry): AsyncGenerator<Uint8Array> {
    const deadline = this.deadline;
    const bad = (why: string) => fail(D.DataError, `${e.name}: ${why}`, { name: e.name, reason: why });
    async function* timed(chunks: AsyncGenerator<Uint8Array>): AsyncGenerator<Uint8Array> {
      for await (const c of chunks) {
        checkTime(deadline);
        yield c;
      }
    }
    let crc = 0;
    let total = 0;
    if (e.method === 0) {
      for await (const c of timed(this.raw(e))) {
        total += c.length;
        crc = crc32(c, crc);
        yield c;
      }
    } else {
      const inflater = zlib.createInflateRaw();
      const src = stream.Readable.from(timed(this.raw(e)));
      src.on('error', err => inflater.destroy(err));
      src.pipe(inflater);
      try {
        for await (const c of inflater as AsyncIterable<Buffer>) {
          total += c.length;
          if (total > e.size) bad('inflates past its declared size');
          crc = crc32(c, crc);
          yield c;
        }
      } catch (err) {
        // zlib's own errors mean corrupt or cut-short data. Anything else, such as an I/O error or the time limit, goes up.
        const code = (err as { code?: unknown })?.code;
        if (typeof code === 'string' && code.startsWith('Z_')) bad((err as Error).message);
        throw err;
      } finally {
        src.destroy();
        inflater.destroy();
      }
      if (inflater.bytesWritten !== e.compressedSize) bad('the deflate data ends before its compressed size');
    }
    if (total !== e.size) bad(`inflates to ${total} bytes, not ${e.size}`);
    if (crc !== e.crc) bad('has the wrong CRC');
  }
}
