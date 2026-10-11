import * as crypto from 'node:crypto';
import type { BiffRecord } from './biff';
import type { CfbEntry, CompoundFile } from './cfb';
import type { RunContext } from './context';
import { checkTime, StopError } from './errors';
import { ByteReader, Truncated } from './oleobj';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';
import type { Decrypt } from './xls';

/** The password Excel opens a workbook with when nobody set one to open it (MS-OFFCRYPTO 2.4.3). */
export const DEFAULT_PASSWORD = 'VelvetSweatshop';
/** The password PowerPoint 97 to 2003 opens a presentation with when nobody set one to open it. */
export const POWERPOINT_PASSWORD = '/01Hannes Ruescher/01';

/** A password an application tries on its own, and the application. */
interface Default {
  password: string;
  application: string;
}
const EXCEL: Default = { password: DEFAULT_PASSWORD, application: 'Excel' };
const POWERPOINT: Default = { password: POWERPOINT_PASSWORD, application: 'PowerPoint' };
/** MS-OFFCRYPTO allows a spin count of at most 10,000,000. */
const MAX_SPIN = 10_000_000;
/** More of a Standard EncryptionInfo than its header and verifier take, which are a few hundred bytes. */
const STANDARD_INFO = 4096;

/** A source's bytes from `from` to `to`, in pieces. */
async function* chunks(source: ByteSource, from: number, to: number): AsyncGenerator<Uint8Array> {
  for (let at = from; at < to; at += 65536) yield source.read(at, Math.min(65536, to - at));
}

const required = (reason: string, data: Record<string, string> = {}) => new StopError(C.Encrypted, D.PasswordRequired, `No password at hand opens the file: ${reason}`, { ...data, reason });
/** A file encrypted with an algorithm this reader does not decrypt. It never reaches OpenSSL. */
const unsupported = (algorithm: string, data: Record<string, string> = {}) =>
  new StopError(C.Encrypted, D.UnsupportedAlgorithm, `The file is encrypted with ${algorithm}, which this reader does not decrypt`, { ...data, algorithm });

/**
 * The only hashes and ciphers that reach OpenSSL: those its default provider has in OpenSSL 3 and 4. A name a file
 * gives is looked up here, and one not here is never passed on.
 */
type HashName = 'md5' | 'sha1' | 'sha256' | 'sha384' | 'sha512';
const HASHES = new Map<string, HashName>([
  ['SHA1', 'sha1'],
  ['SHA256', 'sha256'],
  ['SHA384', 'sha384'],
  ['SHA512', 'sha512'],
]);
const AES = new Map([
  [128, { cbc: 'aes-128-cbc', ecb: 'aes-128-ecb', detail: D.Aes128, bits: 128 }],
  [192, { cbc: 'aes-192-cbc', ecb: 'aes-192-ecb', detail: D.Aes192, bits: 192 }],
  [256, { cbc: 'aes-256-cbc', ecb: 'aes-256-ecb', detail: D.Aes256, bits: 256 }],
]);
/** CALG_AES_128, CALG_AES_192 and CALG_AES_256, by key size. */
const STANDARD_AES = new Map([
  [128, 0x660e],
  [192, 0x660f],
  [256, 0x6610],
]);
/** A value from the file for a message, cut short. */
const shown = (v: string | undefined) => (v ?? 'none').slice(0, 40);

/** A password to try, and the application's default it is, when it is one. */
interface Tried {
  password: string;
  default?: Default;
}

/**
 * The passwords tried, in order: the caller's, each once, then the application's default, which Word has none of.
 * One upload can hold files with different passwords, so a password meant for another file must not keep the default
 * from opening this one.
 */
function passwords(ctx: RunContext, fallback: Default | undefined = EXCEL): Tried[] {
  const given = [...new Set(ctx.settings.passwords)];
  const out: Tried[] = given.map(password => ({ password }));
  return !fallback || given.includes(fallback.password) ? out : [...out, { password: fallback.password, default: fallback }];
}

/** Which password opened the file, and the cipher, as findings. */
function opened(ctx: RunContext, tried: Tried, cipher: D, location: string, bits?: number): void {
  const d = tried.default;
  ctx.report(C.Encrypted, d ? D.DefaultPassword : D.UserPassword, location, d ? { password: d.password, application: d.application } : {});
  ctx.report(C.Encrypted, cipher, location, bits === undefined ? {} : { bits });
}

const utf16le = (s: string) => Buffer.from(s, 'utf16le');
const le32 = (n: number) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n >>> 0);
  return b;
};
const hash = (algorithm: HashName, ...parts: Uint8Array[]) => {
  const h = crypto.createHash(algorithm);
  for (const p of parts) h.update(p);
  return h.digest();
};

/** RC4, which OpenSSL 3 leaves out of its default provider. */
export class Rc4 {
  private readonly s = new Uint8Array(256);
  private i = 0;
  private j = 0;
  constructor(key: Uint8Array) {
    for (let k = 0; k < 256; k++) this.s[k] = k;
    for (let k = 0, j = 0; k < 256; k++) {
      j = (j + this.s[k] + key[k % key.length]) & 255;
      [this.s[k], this.s[j]] = [this.s[j], this.s[k]];
    }
  }

  /** XORs the next keystream bytes into a copy of `data`. */
  crypt(data: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.length);
    for (let k = 0; k < data.length; k++) {
      this.i = (this.i + 1) & 255;
      this.j = (this.j + this.s[this.i]) & 255;
      [this.s[this.i], this.s[this.j]] = [this.s[this.j], this.s[this.i]];
      out[k] = data[k] ^ this.s[(this.s[this.i] + this.s[this.j]) & 255];
    }
    return out;
  }
}

/** The key of one 1024-byte block of an RC4-encrypted binary file, for a password. */
type BlockKey = (block: number) => Uint8Array;

/** Office binary document RC4 (MS-OFFCRYPTO 2.3.6): MD5, a 40-bit truncation and the salt. */
function rc4Key(password: string, salt: Uint8Array): BlockKey {
  const h0 = hash('md5', utf16le(password)).subarray(0, 5);
  const buffer = Buffer.concat(Array.from({ length: 16 }, () => Buffer.concat([h0, salt])));
  const h1 = hash('md5', buffer).subarray(0, 5);
  return block => hash('md5', h1, le32(block));
}

/** RC4 CryptoAPI (MS-OFFCRYPTO 2.3.5.2): SHA-1 of the salt and password, then of the block number. A 40-bit key is padded to 128 bits. */
function cryptoApiKey(password: string, salt: Uint8Array, bits: number): BlockKey {
  const h0 = hash('sha1', salt, utf16le(password));
  return block => {
    const h = hash('sha1', h0, le32(block));
    return bits === 40 ? Buffer.concat([h.subarray(0, 5), Buffer.alloc(11)]) : h.subarray(0, bits / 8);
  };
}

/**
 * The key of a binary document encrypted with RC4 (MS-OFFCRYPTO 2.3.6) or RC4 CryptoAPI (2.3.5), from its
 * EncryptionHeader, which `r` reads from its version on. Returns undefined when no password at hand opens it.
 * `cryptoApiOnly` refuses plain RC4, as PowerPoint does.
 */
function binaryKey(ctx: RunContext, r: ByteReader, stream: string, fallback: Default | undefined, cryptoApiOnly = false): BlockKey | undefined {
  const major = r.u16();
  const minor = r.u16();
  if (major === 1 && minor === 1 && !cryptoApiOnly) {
    const salt = r.bytes(16);
    const verifier = r.bytes(16);
    const verifierHash = r.bytes(16);
    for (const p of passwords(ctx, fallback)) {
      const k = rc4Key(p.password, salt);
      const plain = new Rc4(k(0)).crypt(Buffer.concat([verifier, verifierHash]));
      if (hash('md5', plain.subarray(0, 16)).equals(Buffer.from(plain.subarray(16, 32)))) {
        opened(ctx, p, D.Rc4_40, stream, 40);
        return k;
      }
    }
    return undefined;
  }
  if ((major === 2 || major === 3 || major === 4) && minor === 2) {
    r.skip(4);
    const headerSize = r.u32();
    const header = new ByteReader(r.bytes(headerSize));
    header.skip(8);
    const algorithm = header.u32();
    header.skip(4);
    const bits = header.u32() || 40;
    if (algorithm !== 0x6801 || bits < 40 || bits > 128 || bits % 8) throw unsupported(`CryptoAPI algorithm 0x${algorithm.toString(16)} with a ${bits}-bit key`, { stream });
    const cipher = bits === 40 ? D.Rc4_40 : bits === 128 ? D.Rc4_128 : D.Rc4Other;
    const saltSize = r.u32();
    const salt = r.bytes(saltSize);
    const verifier = r.bytes(16);
    const hashSize = r.u32();
    const verifierHash = r.bytes(Math.min(hashSize, r.left));
    for (const p of passwords(ctx, fallback)) {
      const k = cryptoApiKey(p.password, salt, bits);
      const plain = new Rc4(k(0)).crypt(Buffer.concat([verifier, verifierHash]));
      if (hash('sha1', plain.subarray(0, 16)).equals(Buffer.from(plain.subarray(16, 36)))) {
        opened(ctx, p, cipher, stream, bits);
        return k;
      }
    }
    return undefined;
  }
  throw unsupported(`encryption version ${major}.${minor}`, { stream });
}

/**
 * Opens a workbook that its FILEPASS record (MS-XLS 2.4.117) encrypts with RC4 or RC4 CryptoAPI, with the caller's
 * password or VelvetSweatshop. The records after FILEPASS are encrypted in 1024-byte blocks of the stream, each with its
 * own key. Their headers are not, but take keystream bytes, and some records are never encrypted. XOR obfuscation,
 * the only kind an Excel 5.0 or 95 workbook has, is undone the same way record by record.
 */
export function unlockWorkbook(ctx: RunContext, filePass: BiffRecord, stream: string): Decrypt {
  const r = new ByteReader(filePass.body);
  try {
    // A BIFF5 FILEPASS holds only the XOR key and verifier. BIFF8 names the method first.
    const xor = filePass.body.length === 4 || r.u16() === 0;
    if (xor) {
      const key = r.u16();
      const verifier = r.u16();
      for (const p of passwords(ctx)) {
        const pw = xorPassword(p.password);
        if (!pw.length || xorKey1(pw) !== key || verifier1(pw) !== verifier) continue;
        opened(ctx, p, D.XorObfuscation, stream);
        return xorDecrypt(xorArray1(pw));
      }
      throw required('the workbook is obfuscated with a password that is not VelvetSweatshop', { stream });
    }
    const key = binaryKey(ctx, r, stream, EXCEL);
    if (!key) throw required('the workbook is encrypted with a password that is not VelvetSweatshop', { stream });
    return biffDecrypt(key);
  } catch (e) {
    if (e instanceof Truncated) throw required('a FILEPASS record cut short', { stream });
    throw e;
  }
}

/** XOR obfuscation method 1 tables (MS-OFFCRYPTO 2.3.7.2). */
const PAD = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00];
const INITIAL_CODE = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];
// biome-ignore format: the table as the specification prints it
const XOR_MATRIX = [
  0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09, 0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf, 0x4563, 0x8ac6,
  0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0, 0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40, 0xd849, 0xa0b3, 0x5147, 0xa28e,
  0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a, 0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc,
  0x5fd9, 0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0, 0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec, 0x45a0,
  0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10, 0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168, 0x76b4, 0xed68, 0xcaf1,
  0x85c3, 0x1ba7, 0x374e, 0x6e9c, 0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373,
  0x06e6, 0x0dcc, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4,
];

/**
 * A password as the XOR methods read it (MS-OFFCRYPTO 2.3.7.4): each character's low byte, or its high byte when the
 * low one is 0, at most 15 of them.
 */
function xorPassword(password: string): Uint8Array {
  const out: number[] = [];
  for (let i = 0; i < password.length && out.length < 15; i++) {
    const c = password.charCodeAt(i);
    out.push(c & 0xff ? c & 0xff : (c >> 8) & 0xff);
  }
  return Uint8Array.from(out);
}

/** CreateXorKey_Method1 (MS-OFFCRYPTO 2.3.7.2). */
function xorKey1(pw: Uint8Array): number {
  let key = INITIAL_CODE[pw.length - 1];
  let element = 0x68;
  for (let i = pw.length - 1; i >= 0; i--) {
    let c = pw[i];
    for (let k = 0; k < 7; k++) {
      if (c & 0x40) key ^= XOR_MATRIX[element];
      c *= 2;
      element--;
    }
  }
  return key & 0xffff;
}

/** CreatePasswordVerifier_Method1 (MS-OFFCRYPTO 2.3.7.1). */
function verifier1(pw: Uint8Array): number {
  let v = 0;
  for (const b of [...pw].reverse().concat(pw.length)) v = (((v & 0x4000) >> 14) | ((v << 1) & 0x7fff)) ^ b;
  return (v ^ 0xce4b) & 0xffff;
}

/** One right rotation of a byte. */
const ror = (b: number) => ((b >> 1) | (b << 7)) & 0xff;

/** CreateXorArray_Method1 (MS-OFFCRYPTO 2.3.7.2). */
function xorArray1(pw: Uint8Array): Uint8Array {
  const key = xorKey1(pw);
  const hi = key >> 8;
  const lo = key & 0xff;
  const out = new Uint8Array(16);
  let i = pw.length;
  if (i % 2 === 1) {
    out[i] = ror(PAD[0] ^ hi);
    i--;
    out[i] = ror(pw[pw.length - 1] ^ lo);
  }
  while (i > 0) {
    i--;
    out[i] = ror(pw[i] ^ hi);
    i--;
    out[i] = ror(pw[i] ^ lo);
  }
  let at = 15;
  for (let pad = 15 - pw.length; pad > 0; ) {
    out[at--] = ror(PAD[pad--] ^ hi);
    out[at--] = ror((PAD[pad--] ?? 0) ^ lo);
  }
  return out;
}

/** CreateXorArray_Method2 (MS-OFFCRYPTO 2.3.7.5), which Word uses. */
function xorArray2(pw: Uint8Array): Uint8Array {
  const key = xorKey1(pw);
  const out = new Uint8Array(16);
  for (let i = 0; i < 16; i++) out[i] = i < pw.length ? pw[i] : PAD[i - pw.length];
  for (let i = 0; i < 16; i += 2) {
    out[i] = ror(out[i] ^ (key & 0xff));
    out[i + 1] = ror(out[i + 1] ^ (key >> 8));
  }
  return out;
}

/**
 * Undoes XOR obfuscation method 1 in each record after FILEPASS (MS-OFFCRYPTO 2.3.7.3). The array index of a record's
 * first byte is its stream offset plus its size, as Excel and LibreOffice count it, and each byte is XORed and then
 * rotated right by 5 bits. The records MS-XLS leaves plain, and BoundSheet8's lbPlyPos, stay as they are.
 */
function xorDecrypt(array: Uint8Array): Decrypt {
  return r => {
    if (PLAIN.has(r.type)) return r.body;
    const out = new Uint8Array(r.body);
    const start = r.offset + 4;
    for (let k = r.type === 0x0085 ? 4 : 0; k < out.length; k++) {
      const v = out[k] ^ array[(start + r.body.length + k) % 16];
      out[k] = ((v >> 5) | (v << 3)) & 0xff;
    }
    return out;
  };
}

/** A stream decrypted where it is read. */
export type StreamDecrypt = (offset: number, bytes: Uint8Array) => Uint8Array;

/** What opens an encrypted Word 97 to 2003 document: a decryption for each of its three encrypted streams. */
export interface WordKey {
  word: StreamDecrypt;
  table: StreamDecrypt;
  data: StreamDecrypt;
}

/** RC4 keyed per 512-byte block, as Word encrypts its streams, applied from `plain` on. */
function blockStream(key: BlockKey, plain: number): StreamDecrypt {
  const cache = new Map<number, Uint8Array>();
  const keystream = (b: number) => {
    let ks = cache.get(b);
    if (!ks) {
      if (cache.size > 64) cache.clear();
      ks = new Rc4(key(b)).crypt(new Uint8Array(512));
      cache.set(b, ks);
    }
    return ks;
  };
  return (offset, bytes) => {
    const out = new Uint8Array(bytes);
    for (let k = Math.max(0, plain - offset); k < out.length; k++) {
      const at = offset + k;
      out[k] ^= keystream(Math.floor(at / 512))[at % 512];
    }
    return out;
  };
}

/** XOR obfuscation method 2 over a stream (MS-OFFCRYPTO 2.3.7.6), from `plain` on. */
function xorStream(array: Uint8Array, plain: number): StreamDecrypt {
  return (offset, bytes) => {
    const out = new Uint8Array(bytes);
    for (let k = Math.max(0, plain - offset); k < out.length; k++) {
      const a = array[(offset + k) % 16];
      // A byte that was 0x00 or equal to its array element was left as it was.
      if (out[k] !== 0 && out[k] !== a) out[k] ^= a;
    }
    return out;
  };
}

/**
 * Opens an encrypted Word 97 to 2003 document (MS-DOC 2.2.6) with the caller's passwords. XOR obfuscation keeps its
 * verifier in FibBase.lKey. RC4 and RC4 CryptoAPI keep their EncryptionHeader in the table stream's first lKey bytes.
 * The WordDocument stream stays plain for its first 68 bytes, and Word has no default password.
 */
export async function unlockDocument(ctx: RunContext, fib: Uint8Array, table: ByteSource, stream: string): Promise<WordKey> {
  const flags = fib[10] | (fib[11] << 8);
  const lKey = (fib[14] | (fib[15] << 8) | (fib[16] << 16) | (fib[17] << 24)) >>> 0;
  if (flags & 0x8000) {
    for (const p of passwords(ctx, undefined)) {
      const pw = xorPassword(p.password);
      if (!pw.length || ((xorKey1(pw) << 16) | verifier1(pw)) >>> 0 !== lKey) continue;
      opened(ctx, p, D.XorObfuscation, stream);
      const array = xorArray2(pw);
      return { word: xorStream(array, 68), table: xorStream(array, 0), data: xorStream(array, 0) };
    }
    throw required('the document is obfuscated with a password not at hand', { stream });
  }
  if (lKey < 8 || lKey > 65536) throw required('an EncryptionHeader of an impossible size', { stream });
  try {
    const key = binaryKey(ctx, new ByteReader(await table.read(0, lKey)), stream, undefined);
    if (!key) throw required('the document is encrypted with a password not at hand', { stream });
    return { word: blockStream(key, 68), table: blockStream(key, lKey), data: blockStream(key, 0) };
  } catch (e) {
    if (e instanceof Truncated) throw required('an EncryptionHeader cut short', { stream });
    throw e;
  }
}

/**
 * Opens an encrypted PowerPoint 97 to 2003 presentation (MS-PPT 2.3.7): `session` is the CryptSession10Container's
 * data, an RC4 CryptoAPI EncryptionHeader. Returns the key of each persist object, whose block number is its persist
 * id, after the caller's passwords and PowerPoint's own default have been tried.
 */
export function unlockPresentation(ctx: RunContext, session: Uint8Array, stream: string): (id: number) => Rc4 {
  try {
    const key = binaryKey(ctx, new ByteReader(session), stream, POWERPOINT, true);
    if (!key) throw required('the presentation is encrypted with a password not at hand', { stream });
    return id => new Rc4(key(id));
  } catch (e) {
    if (e instanceof Truncated) throw required('a CryptSession10Container cut short', { stream });
    throw e;
  }
}

/** Records MS-XLS 2.2.10 never encrypts: BOF, FILEPASS, UsrExcl, FileLock, InterfaceHdr, RRDInfo and RRDHead. */
const PLAIN = new Set([0x0809, 0x002f, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138]);

/** Only the records after FILEPASS reach the function it returns. */
function biffDecrypt(key: BlockKey): Decrypt {
  let block = -1;
  let stream: Uint8Array = new Uint8Array(0);
  const keystream = (b: number) => {
    if (b !== block) {
      block = b;
      stream = new Rc4(key(b)).crypt(new Uint8Array(1024));
    }
    return stream;
  };
  return r => {
    if (PLAIN.has(r.type)) return r.body;
    const out = new Uint8Array(r.body);
    // BoundSheet8's lbPlyPos stays plain, so a reader can find each sheet without the key.
    const from = r.type === 0x0085 ? 4 : 0;
    for (let k = from; k < out.length; k++) {
      const at = r.offset + 4 + k;
      out[k] ^= keystream(Math.floor(at / 1024))[at % 1024];
    }
    return out;
  };
}

/** What an encrypted OOXML package's EncryptionInfo says, with the key a password gives. */
interface PackageCipher {
  /** The intermediate key that decrypts the package, or undefined when the password is wrong. */
  key(password: string): Uint8Array | undefined;
  /** Decrypts the package's data after its 8-byte size, which `read` reads from its start. */
  source(read: (offset: number, length: number) => Promise<Uint8Array>, size: number, key: Uint8Array): ByteSource;
  cipher: Aes;
}

/**
 * Inspects an encrypted OOXML package: EncryptionInfo (MS-OFFCRYPTO 2.3.4) names the cipher, and the caller's password
 * or VelvetSweatshop gives the key. The package it decrypts goes through the OOXML rules, read in place: Agile
 * encryption works in 4,096-byte segments and Standard in 16-byte blocks, so nothing is decrypted that is not read.
 */
export async function openPackage(ctx: RunContext, cfb: CompoundFile, info: CfbEntry, pkg: CfbEntry): Promise<{ source: ByteSource; size: number }> {
  const source = cfb.source(info);
  const version = await source.read(0, 8);
  const major = version.length >= 4 ? version[0] | (version[1] << 8) : 0;
  const minor = version.length >= 4 ? version[2] | (version[3] << 8) : 0;
  let cipher: PackageCipher;
  // Agile's XML goes to the tokenizer in pieces. Standard's binary header is a few hundred bytes.
  if (major === 4 && minor === 4) cipher = await agile(ctx, chunks(source, 8, info.size));
  else if ((major === 3 || major === 4) && minor === 2) cipher = standard(await source.read(8, STANDARD_INFO));
  else throw unsupported(`encryption version ${major}.${minor}`, { stream: info.name });
  const head = await cfb.source(pkg).read(0, 8);
  if (head.length < 8) throw required('an EncryptedPackage stream cut short', { stream: pkg.name });
  const size = Number(Buffer.from(head).readBigUInt64LE(0));
  if (size > pkg.size - 8) throw required('an EncryptedPackage stream shorter than the size it gives', { stream: pkg.name });
  for (const p of passwords(ctx)) {
    checkTime(ctx.settings.deadline);
    const key = cipher.key(p.password);
    if (!key) continue;
    opened(ctx, p, cipher.cipher.detail, pkg.name, cipher.cipher.bits);
    const data = cfb.source(pkg);
    return { source: cipher.source((offset, length) => data.read(8 + offset, length), size, key), size };
  }
  throw required('the package is encrypted with a password that is not VelvetSweatshop', { stream: pkg.name });
}

type Aes = { cbc: string; ecb: string; detail: D; bits: number };

/** AES without padding over whole blocks: ECB, or CBC from `iv`. */
function aes(cipher: Aes, key: Uint8Array, data: Uint8Array, iv?: Uint8Array): Uint8Array {
  const d = crypto.createDecipheriv(iv ? cipher.cbc : cipher.ecb, key, iv ?? null);
  d.setAutoPadding(false);
  const whole = data.subarray(0, data.length - (data.length % 16));
  return Buffer.concat([d.update(whole), d.final()]);
}

/**
 * AES in CFB with an 8-bit window, which Agile's ChainingModeCFB names, built on AES-ECB. Each byte is XORed with the
 * first byte of the encryption of the 16 ciphertext bytes before it, the IV standing in front. Decryption knows every
 * ciphertext byte, so all those blocks go through ECB in one call.
 */
function aesCfb8(cipher: Aes, key: Uint8Array, data: Uint8Array, iv: Uint8Array): Uint8Array {
  const feed = Buffer.concat([iv.subarray(0, 16), data]);
  const blocks = Buffer.alloc(data.length * 16);
  for (let i = 0; i < data.length; i++) feed.copy(blocks, i * 16, i, i + 16);
  const e = crypto.createCipheriv(cipher.ecb, key, null);
  e.setAutoPadding(false);
  const stream = Buffer.concat([e.update(blocks), e.final()]);
  const out = new Uint8Array(data.length);
  for (let i = 0; i < data.length; i++) out[i] = data[i] ^ stream[i * 16];
  return out;
}

/** Decrypts with the chaining an Agile element names. */
const agileDecrypt = (cipher: Aes, cfb: boolean, key: Uint8Array, data: Uint8Array, iv: Uint8Array) => (cfb ? aesCfb8(cipher, key, data, iv) : aes(cipher, key, data, iv));

/** Standard encryption (MS-OFFCRYPTO 2.3.4.5): AES-ECB with a key from SHA-1 over 50,000 rounds. */
function standard(b: Uint8Array): PackageCipher {
  const r = new ByteReader(b);
  try {
    const headerSize = r.u32();
    const header = new ByteReader(r.bytes(headerSize));
    header.skip(8);
    const algorithm = header.u32();
    const hashAlgorithm = header.u32();
    const bits = header.u32();
    // CALG_AES_128, _192 and _256, which name their key size, with SHA-1.
    const cipher = AES.get(bits);
    if (!cipher || algorithm !== STANDARD_AES.get(bits) || (hashAlgorithm !== 0x8004 && hashAlgorithm !== 0))
      throw unsupported(`Standard encryption with algorithm 0x${algorithm.toString(16)}, hash 0x${hashAlgorithm.toString(16)} and a ${bits}-bit key`);
    r.skip(4);
    const salt = r.bytes(16);
    const verifier = r.bytes(16);
    r.skip(4);
    const verifierHash = r.bytes(32);
    const keyOf = (password: string) => {
      let h = hash('sha1', salt, utf16le(password));
      for (let i = 0; i < 50000; i++) h = hash('sha1', le32(i), h);
      const final = hash('sha1', h, le32(0));
      const x = (fill: number) => {
        const buf = Buffer.alloc(64, fill);
        for (let i = 0; i < final.length; i++) buf[i] ^= final[i];
        return hash('sha1', buf);
      };
      return Buffer.concat([x(0x36), x(0x5c)]).subarray(0, bits / 8);
    };
    return {
      cipher,
      key(password) {
        const key = keyOf(password);
        const v = aes(cipher, key, verifier);
        const vh = aes(cipher, key, verifierHash);
        return hash('sha1', v).equals(Buffer.from(vh.subarray(0, 20))) ? key : undefined;
      },
      source: (read, size, key) => blockSource(size, 16, async (index, count) => aes(cipher, key, await read(index * 16, count * 16))),
    };
  } catch (e) {
    if (e instanceof Truncated) throw required('a Standard EncryptionInfo cut short');
    throw e;
  }
}

/** Agile encryption (MS-OFFCRYPTO 2.3.4.10): an XML descriptor, a password key encryptor, and AES-CBC in 4,096-byte segments. */
async function agile(ctx: RunContext, xml: AsyncIterable<Uint8Array>): Promise<PackageCipher> {
  const found: { keyData?: Record<string, string>; encryptedKey?: Record<string, string> } = {};
  const ENCRYPTION = 'http://schemas.microsoft.com/office/2006/encryption';
  const PASSWORD = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
  await ctx.parseXml('EncryptionInfo', xml, {
    start(el) {
      const all = Object.fromEntries(el.attributes.filter(a => a.ns === '').map(a => [a.local, a.value]));
      if (el.ns === ENCRYPTION && el.local === 'keyData') found.keyData ??= all;
      else if (el.ns === PASSWORD && el.local === 'encryptedKey') found.encryptedKey ??= all;
    },
  });
  const kd = found.keyData;
  const ek = found.encryptedKey;
  if (!kd || !ek) throw required('an Agile descriptor without key data or a password key');
  /** The hash, cipher and chaining a key data or key encryptor element names, when all are on the allowlist. */
  const algorithms = (a: Record<string, string>) => {
    const h = HASHES.get(a.hashAlgorithm);
    const c = AES.get(Number(a.keyBits));
    const chaining = a.cipherChaining;
    if (!h || !c || a.cipherAlgorithm !== 'AES' || (chaining !== 'ChainingModeCBC' && chaining !== 'ChainingModeCFB') || Number(a.blockSize) !== 16)
      throw unsupported(`Agile encryption with ${shown(a.cipherAlgorithm)} in ${shown(a.cipherChaining)}, a ${shown(a.keyBits)}-bit key and ${shown(a.hashAlgorithm)}`);
    return { h, c, cfb: chaining === 'ChainingModeCFB' };
  };
  const { h: dh, c: dc, cfb: dcfb } = algorithms(kd);
  const { h: eh, c: ec, cfb: ecfb } = algorithms(ek);
  const spin = Number(ek.spinCount);
  if (!Number.isInteger(spin) || spin < 0 || spin > MAX_SPIN) throw required('an Agile spin count past 10,000,000');
  const b64 = (s: string | undefined) => Buffer.from(s ?? '', 'base64');
  const pad = (b: Uint8Array, n: number, fill: number) => (b.length >= n ? b.subarray(0, n) : Buffer.concat([b, Buffer.alloc(n - b.length, fill)]));
  const salt = b64(ek.saltValue);
  const BLOCKS = { input: Buffer.from('fea7d2763b4b9e79', 'hex'), value: Buffer.from('d7aa0f6d3061344e', 'hex'), key: Buffer.from('146e0be7abacd0d6', 'hex') };
  const dsalt = b64(kd.saltValue);
  return {
    cipher: dc,
    key(password) {
      let h = hash(eh, salt, utf16le(password));
      for (let i = 0; i < spin; i++) {
        if (i % 65536 === 0) checkTime(ctx.settings.deadline);
        h = hash(eh, le32(i), h);
      }
      const derive = (block: Uint8Array) => pad(hash(eh, h, block), ec.bits / 8, 0x36);
      const iv = pad(salt, 16, 0x36);
      const input = agileDecrypt(ec, ecfb, derive(BLOCKS.input), b64(ek.encryptedVerifierHashInput), iv).subarray(0, Number(ek.saltSize) || 16);
      const expected = agileDecrypt(ec, ecfb, derive(BLOCKS.value), b64(ek.encryptedVerifierHashValue), iv);
      const got = hash(eh, input);
      if (!got.equals(Buffer.from(expected.subarray(0, got.length)))) return undefined;
      return agileDecrypt(ec, ecfb, derive(BLOCKS.key), b64(ek.encryptedKeyValue), iv).subarray(0, dc.bits / 8);
    },
    source: (read, size, key) =>
      blockSource(size, 4096, async (index, count) => {
        const raw = await read(index * 4096, count * 4096);
        const out: Uint8Array[] = [];
        for (let k = 0; k * 4096 < raw.length; k++) out.push(agileDecrypt(dc, dcfb, key, raw.subarray(k * 4096, (k + 1) * 4096), pad(hash(dh, dsalt, le32(index + k)), 16, 0x36)));
        return Buffer.concat(out);
      }),
  };
}

/** A source over decrypted data, decrypting the whole blocks each read covers, and keeping the last ones decrypted. */
function blockSource(size: number, unit: number, decrypt: (index: number, count: number) => Promise<Uint8Array>): ByteSource {
  let cached: { index: number; bytes: Uint8Array } | undefined;
  return {
    size: async () => size,
    async read(offset, length) {
      const end = Math.min(size, offset + length);
      const first = Math.floor(offset / unit);
      const last = Math.floor((end - 1) / unit);
      let bytes: Uint8Array;
      if (cached && cached.index <= first && cached.index + cached.bytes.length / unit > last) bytes = cached.bytes.subarray((first - cached.index) * unit);
      else {
        // A small read decrypts 64 KiB from where it starts, since the ZIP reader reads on from there.
        const count = Math.max(last - first + 1, Math.ceil(65536 / unit));
        bytes = await decrypt(first, count);
        cached = { index: first, bytes };
      }
      return bytes.subarray(offset - first * unit, end - first * unit);
    },
  };
}
