import { guid } from './cfb';

/** Thrown by ByteReader when a structure runs past the end of its bytes. Parsers catch it and keep what they read. */
export class Truncated extends Error {}

/** Runs `read`. A structure that runs past its data gives what `partial` returns instead, and any other error goes up. */
export function bounded<T>(read: () => T, partial: () => T): T {
  try {
    return read();
  } catch (e) {
    if (e instanceof Truncated) return partial();
    throw e;
  }
}

/** Little-endian reads over bytes in memory, bounded by their end. */
export class ByteReader {
  pos = 0;
  constructor(readonly b: Uint8Array) {}

  get left(): number {
    return this.b.length - this.pos;
  }

  private need(n: number): number {
    if (n < 0 || this.pos + n > this.b.length) throw new Truncated('The structure runs past its data');
    const at = this.pos;
    this.pos += n;
    return at;
  }

  u8(): number {
    return this.b[this.need(1)];
  }

  u16(): number {
    const at = this.need(2);
    return this.b[at] | (this.b[at + 1] << 8);
  }

  u32(): number {
    const at = this.need(4);
    return (this.b[at] | (this.b[at + 1] << 8) | (this.b[at + 2] << 16) | (this.b[at + 3] << 24)) >>> 0;
  }

  bytes(n: number): Uint8Array {
    const at = this.need(n);
    return this.b.subarray(at, at + n);
  }

  skip(n: number): void {
    this.need(n);
  }

  /** An ANSI string ended by a NUL within `max` bytes, read as Latin-1. */
  ansiz(max = 1024): string {
    const end = this.b.indexOf(0, this.pos);
    if (end < 0 || end - this.pos >= max) throw new Truncated('No NUL ends the string');
    const s = latin1(this.b.subarray(this.pos, end));
    this.pos = end + 1;
    return s;
  }

  /** A 32-bit byte count, then that many bytes of ANSI text whose last is a NUL. */
  lengthPrefixedAnsi(max = 65536): string {
    const n = this.u32();
    if (n > max) throw new Truncated('A string length past any real string');
    if (n === 0) return '';
    const s = this.bytes(n);
    return latin1(s[n - 1] === 0 ? s.subarray(0, n - 1) : s);
  }

  /** A 32-bit count of UTF-16 code units, then the units, the last of which may be a NUL. */
  lengthPrefixedUnicode(max = 32768): string {
    const n = this.u32();
    if (n > max) throw new Truncated('A string length past any real string');
    return utf16(this.bytes(n * 2)).replace(/\0+$/, '');
  }
}

export const latin1 = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('latin1');
export const utf16 = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength - (b.byteLength % 2)).toString('utf16le');

/** The class an OLE object names in its CompObj stream (MS-OLEDS 2.3.8): the user type and the ProgID. */
export interface CompObj {
  userType?: string;
  progId?: string;
}

export function parseCompObj(b: Uint8Array): CompObj {
  const r = new ByteReader(b);
  const out: CompObj = {};
  return bounded(
    () => {
      r.skip(28);
      out.userType = r.lengthPrefixedAnsi();
      const marker = r.u32();
      if (marker === 0xffffffff || marker === 0xfffffffe) r.skip(4);
      else if (marker !== 0) r.skip(marker);
      // MS-OLEDS calls it Reserved1. Windows writes the ProgID there.
      out.progId = r.lengthPrefixedAnsi();
      return out;
    },
    () => out,
  );
}

/** URL moniker, file moniker and composite moniker class ids (MS-OSHARED 2.3.7). */
const URL_MONIKER = '79EAC9E0-BAF9-11CE-8C82-00AA004BA90B';
const FILE_MONIKER = '00000303-0000-0000-C000-000000000046';
const COMPOSITE_MONIKER = '00000309-0000-0000-C000-000000000046';

/** What a moniker names: a URL moniker's URL, or a file moniker's path. */
export interface MonikerTarget {
  target: string;
  /** A URL moniker, whose target loads from wherever it names, however it is spelled. */
  url: boolean;
}

/**
 * The address a moniker names: a URL moniker's URL, a file moniker's path, or the first of a composite's that names
 * one. Other monikers name nothing this reads, and give ''. The reader stops after the moniker it reads.
 */
export function readMoniker(r: ByteReader, depth = 0): MonikerTarget {
  const clsid = guid(r.bytes(16));
  if (clsid === URL_MONIKER) {
    const n = r.u32();
    const raw = r.bytes(Math.min(n, r.left));
    return { target: utf16(raw).split('\0')[0], url: true };
  }
  if (clsid === FILE_MONIKER) {
    const anti = r.u16();
    const ansi = r.lengthPrefixedAnsi();
    r.skip(2 + 2 + 16);
    const unicodeSize = r.u32();
    let path = ansi;
    if (unicodeSize > 0) {
      const bytes = r.u32();
      r.skip(2);
      path = utf16(r.bytes(bytes));
    }
    return { target: '..\\'.repeat(Math.min(anti, 64)) + path, url: false };
  }
  if (clsid === COMPOSITE_MONIKER && depth < 8) {
    const count = r.u32();
    for (let i = 0; i < Math.min(count, 64); i++) {
      const found = readMoniker(r, depth + 1);
      if (found.target) return found;
    }
  }
  return { target: '', url: false };
}

/** What an OLE stream (MS-OLEDS 2.3.3) says: whether the object is a link, and what it links to. */
export interface OleStreamInfo extends MonikerTarget {
  linked: boolean;
}

export function parseOleStream(b: Uint8Array): OleStreamInfo {
  const r = new ByteReader(b);
  const out = { linked: false, target: '', url: false };
  return bounded(
    () => {
      r.skip(4);
      out.linked = (r.u32() & 1) === 1;
      if (!out.linked) return out;
      r.skip(8);
      r.skip(r.u32());
      // The relative moniker comes first, then the absolute one, which Office prefers when both resolve.
      const monikers: MonikerTarget[] = [];
      for (let i = 0; i < 2; i++) {
        const size = r.u32();
        if (size > 0) monikers.push(readMoniker(new ByteReader(r.bytes(size))));
      }
      return Object.assign(
        out,
        monikers.reverse().find(m => m.target),
      );
    },
    () => out,
  );
}

/** An embedded file in Packager's native data, which oleobj reads the same way. */
export interface PackagedFile {
  fileName: string;
  sourcePath: string;
  tempPath: string;
  /** False when the data names a file it does not hold, a link. */
  embedded: boolean;
  size: number;
  /** Where the file's bytes start in the native data. */
  dataAt: number;
}

/**
 * Packager's native data, as oleobj reads it: a 16-bit type, the label and source path ended by NULs, eight bytes,
 * the temporary path ended by a NUL, and the size of the file's data. Without the size it is a link. Undefined when it
 * does not read that far.
 */
export function parsePackage(b: Uint8Array): PackagedFile | undefined {
  const r = new ByteReader(b);
  return bounded<PackagedFile | undefined>(
    () => {
      r.skip(2);
      const fileName = r.ansiz();
      const sourcePath = r.ansiz();
      r.skip(8);
      const tempPath = r.ansiz();
      return bounded<PackagedFile>(
        () => ({ fileName, sourcePath, tempPath, embedded: true, size: r.u32(), dataAt: r.pos }),
        () => ({ fileName, sourcePath, tempPath, embedded: false, size: 0, dataAt: r.pos }),
      );
    },
    () => undefined,
  );
}

/** An OLE 1.0 object (MS-OLEDS 2.2), as RTF's objdata holds one. */
export interface Ole1Object {
  /** 1 linked, 2 embedded, 5 a static picture. */
  formatId: number;
  className: string;
  topicName: string;
  itemName: string;
  /** The native data of an embedded object. */
  data?: Uint8Array;
  /** Where the native data starts in the object's bytes, after the size before it. */
  dataAt?: number;
}

export function parseOle1(b: Uint8Array): Ole1Object | undefined {
  const r = new ByteReader(b);
  return bounded<Ole1Object | undefined>(
    () => {
      r.skip(4);
      const formatId = r.u32();
      if (formatId !== 1 && formatId !== 2) return { formatId, className: '', topicName: '', itemName: '' };
      const className = r.lengthPrefixedAnsi();
      const topicName = r.lengthPrefixedAnsi();
      const itemName = r.lengthPrefixedAnsi();
      if (formatId === 1) return { formatId, className, topicName, itemName };
      const size = r.u32();
      const dataAt = r.pos;
      return { formatId, className, topicName, itemName, data: r.bytes(Math.min(size, r.left)), dataAt };
    },
    () => undefined,
  );
}

const le32 = (n: number) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n >>> 0);
  return b;
};

/**
 * Data that holds a 32-bit size at `dataAt - 4` and that many bytes at `dataAt`, with `data` in their place and the
 * size set to match. What came before and after stays. Packager's native data and an OLE 1.0 object have this layout.
 */
export function replaceSized(bytes: Uint8Array, dataAt: number, oldSize: number, data: Uint8Array): Uint8Array {
  return Buffer.concat([bytes.subarray(0, dataAt - 4), le32(data.length), data, bytes.subarray(Math.min(bytes.length, dataAt + oldSize))]);
}

/** A stream that starts with the size of what follows, such as Ole10Native, around `data`. */
export const sizedStream = (data: Uint8Array) => Buffer.concat([le32(data.length), data]);
