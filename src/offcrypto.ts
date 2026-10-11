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

/** The passwords tried, in order: the caller's, then Excel's default. */
function passwords(ctx: RunContext): Array<{ password: string; supplied: boolean }> {
  const out = ctx.settings.password !== undefined ? [{ password: ctx.settings.password, supplied: true }] : [];
  return [...out, { password: DEFAULT_PASSWORD, supplied: false }];
}

/** Which password opened the file, and the cipher, as findings. */
function opened(ctx: RunContext, supplied: boolean, cipher: D, location: string, bits: number): void {
  ctx.report(C.Encrypted, supplied ? D.UserPassword : D.DefaultPassword, location, supplied ? {} : { password: DEFAULT_PASSWORD });
  ctx.report(C.Encrypted, cipher, location, { bits });
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
 * Opens a workbook that its FILEPASS record (MS-XLS 2.4.117) encrypts with RC4 or RC4 CryptoAPI, with the caller's
 * password or VelvetSweatshop. The records after FILEPASS are encrypted in 1024-byte blocks of the stream, each with its
 * own key. Their headers are not, but take keystream bytes, and some records are never encrypted.
 */
export function unlockWorkbook(ctx: RunContext, filePass: BiffRecord, stream: string): Decrypt {
  const r = new ByteReader(filePass.body);
  let key: BlockKey | undefined;
  let cipher: D = D.Rc4_40;
  let bits = 40;
  try {
    const type = r.u16();
    if (type !== 1) throw unsupported('XOR obfuscation', { stream });
    const major = r.u16();
    const minor = r.u16();
    if (major === 1 && minor === 1) {
      const salt = r.bytes(16);
      const verifier = r.bytes(16);
      const verifierHash = r.bytes(16);
      for (const p of passwords(ctx)) {
        const k = rc4Key(p.password, salt);
        const plain = new Rc4(k(0)).crypt(Buffer.concat([verifier, verifierHash]));
        if (hash('md5', plain.subarray(0, 16)).equals(Buffer.from(plain.subarray(16, 32)))) {
          key = k;
          opened(ctx, p.supplied, D.Rc4_40, stream, 40);
          break;
        }
      }
    } else if ((major === 2 || major === 3 || major === 4) && minor === 2) {
      r.skip(4);
      const headerSize = r.u32();
      const header = new ByteReader(r.bytes(headerSize));
      header.skip(8);
      const algorithm = header.u32();
      header.skip(4);
      bits = header.u32() || 40;
      if (algorithm !== 0x6801 || bits < 40 || bits > 128 || bits % 8) throw unsupported(`CryptoAPI algorithm 0x${algorithm.toString(16)} with a ${bits}-bit key`, { stream });
      cipher = bits === 40 ? D.Rc4_40 : bits === 128 ? D.Rc4_128 : D.Rc4Other;
      const saltSize = r.u32();
      const salt = r.bytes(saltSize);
      const verifier = r.bytes(16);
      const hashSize = r.u32();
      const verifierHash = r.bytes(Math.min(hashSize, r.left));
      for (const p of passwords(ctx)) {
        const k = cryptoApiKey(p.password, salt, bits);
        const plain = new Rc4(k(0)).crypt(Buffer.concat([verifier, verifierHash]));
        if (hash('sha1', plain.subarray(0, 16)).equals(Buffer.from(plain.subarray(16, 36)))) {
          key = k;
          opened(ctx, p.supplied, cipher, stream, bits);
          break;
        }
      }
    } else throw unsupported(`encryption version ${major}.${minor}`, { stream });
  } catch (e) {
    if (e instanceof Truncated) throw required('a FILEPASS record cut short', { stream });
    throw e;
  }
  if (!key) throw required('the workbook is encrypted with a password that is not VelvetSweatshop', { stream });
  return biffDecrypt(key);
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
    opened(ctx, p.supplied, cipher.cipher.detail, pkg.name, cipher.cipher.bits);
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
  /** The hash and cipher a key data or key encryptor element names, when both are on the allowlist. */
  const algorithms = (a: Record<string, string>) => {
    const h = HASHES.get(a.hashAlgorithm);
    const c = AES.get(Number(a.keyBits));
    if (!h || !c || a.cipherAlgorithm !== 'AES' || a.cipherChaining !== 'ChainingModeCBC' || Number(a.blockSize) !== 16)
      throw unsupported(`Agile encryption with ${shown(a.cipherAlgorithm)} in ${shown(a.cipherChaining)}, a ${shown(a.keyBits)}-bit key and ${shown(a.hashAlgorithm)}`);
    return { h, c };
  };
  const { h: dh, c: dc } = algorithms(kd);
  const { h: eh, c: ec } = algorithms(ek);
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
      const input = aes(ec, derive(BLOCKS.input), b64(ek.encryptedVerifierHashInput), iv).subarray(0, Number(ek.saltSize) || 16);
      const expected = aes(ec, derive(BLOCKS.value), b64(ek.encryptedVerifierHashValue), iv);
      const got = hash(eh, input);
      if (!got.equals(Buffer.from(expected.subarray(0, got.length)))) return undefined;
      return aes(ec, derive(BLOCKS.key), b64(ek.encryptedKeyValue), iv).subarray(0, dc.bits / 8);
    },
    source: (read, size, key) =>
      blockSource(size, 4096, async (index, count) => {
        const raw = await read(index * 4096, count * 4096);
        const out: Uint8Array[] = [];
        for (let k = 0; k * 4096 < raw.length; k++) out.push(aes(dc, key, raw.subarray(k * 4096, (k + 1) * 4096), pad(hash(dh, dsalt, le32(index + k)), 16, 0x36)));
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
