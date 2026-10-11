import * as zlib from 'node:zlib';
import { crc32 } from './crc32';
import type { ByteSink } from './types';

const MAX16 = 0xffff;
const MAX32 = 0xffffffff;
/** 1980-01-01 00:00, the earliest date a ZIP header holds, so new entries do not depend on the clock. */
const DOS_DATE = (0 << 9) | (1 << 5) | 1;

/** What the central directory says of an entry whose compressed bytes are copied unchanged. */
export interface CopiedEntry {
  name: string;
  /** 0 stored, 8 deflate. */
  method: number;
  crc: number;
  compressedSize: number;
  size: number;
  dosTime?: number;
  dosDate?: number;
}

class Builder {
  private readonly parts: number[] = [];
  u16(v: number): this {
    this.parts.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }
  u32(v: number): this {
    return this.u16(v & 0xffff).u16(Math.floor(v / 0x10000) & 0xffff);
  }
  u64(v: number): this {
    return this.u32(v % 0x100000000).u32(Math.floor(v / 0x100000000));
  }
  bytes(b: Uint8Array): this {
    for (const x of b) this.parts.push(x);
    return this;
  }
  build(): Uint8Array {
    return Uint8Array.from(this.parts);
  }
}

/**
 * Writes a ZIP archive to an append-only sink. Local headers carry the CRC and sizes, so no entry needs a data
 * descriptor. ZIP64 records appear only where a size, an offset or the entry count needs them.
 */
export class ZipWriter {
  private offset = 0;
  private readonly central: Uint8Array[] = [];
  private readonly names = new Set<string>();

  /** `forceZip64` writes ZIP64 records for every entry and the end of the archive, as a test of readers. */
  constructor(
    private readonly sink: ByteSink,
    private readonly options: { forceZip64?: boolean } = {},
  ) {}

  private async put(b: Uint8Array): Promise<void> {
    await this.sink.write(b);
    this.offset += b.length;
  }

  /** Adds an entry from bytes in memory, deflated unless `compress` is false. A name ending in "/" is a directory. */
  async add(name: string, data: Uint8Array, options: { compress?: boolean; dosTime?: number; dosDate?: number } = {}): Promise<void> {
    const compress = (options.compress ?? true) && data.length > 0;
    const body = compress ? zlib.deflateRawSync(data) : data;
    const info: CopiedEntry = { name, method: compress ? 8 : 0, crc: crc32(data), compressedSize: body.length, size: data.length, dosTime: options.dosTime, dosDate: options.dosDate };
    await this.copy(
      info,
      (async function* () {
        yield body;
      })(),
    );
  }

  /** Adds an entry whose compressed bytes come from `chunks` unchanged, such as an entry of the archive being rewritten. */
  async copy(info: CopiedEntry, chunks: AsyncIterable<Uint8Array>): Promise<void> {
    if (this.names.has(info.name)) throw new Error(`Entry ${info.name} is already in the archive`);
    this.names.add(info.name);
    const name = Buffer.from(info.name, 'utf8');
    const flags = /^[\x20-\x7e]*$/.test(info.name) ? 0 : 0x0800;
    const big = this.options.forceZip64 === true || info.size >= MAX32 || info.compressedSize >= MAX32;
    const localOffset = this.offset;
    const time = info.dosTime ?? 0;
    const date = info.dosDate ?? DOS_DATE;
    const local = new Builder()
      .u32(0x04034b50)
      .u16(big ? 45 : 20)
      .u16(flags)
      .u16(info.method)
      .u16(time)
      .u16(date)
      .u32(info.crc)
      .u32(big ? MAX32 : info.compressedSize)
      .u32(big ? MAX32 : info.size)
      .u16(name.length)
      .u16(big ? 20 : 0)
      .bytes(name);
    if (big) local.u16(0x0001).u16(16).u64(info.size).u64(info.compressedSize);
    await this.put(local.build());
    let written = 0;
    for await (const c of chunks) {
      written += c.length;
      await this.put(c);
    }
    if (written !== info.compressedSize) throw new Error(`Entry ${info.name} has ${written} compressed bytes, not ${info.compressedSize}`);
    const farOffset = this.options.forceZip64 === true || localOffset >= MAX32;
    const extra = new Builder();
    if (big) extra.u64(info.size).u64(info.compressedSize);
    if (farOffset) extra.u64(localOffset);
    const extraBytes = extra.build();
    const version = big || farOffset ? 45 : 20;
    const central = new Builder()
      .u32(0x02014b50)
      .u16(version)
      .u16(version)
      .u16(flags)
      .u16(info.method)
      .u16(time)
      .u16(date)
      .u32(info.crc)
      .u32(big ? MAX32 : info.compressedSize)
      .u32(big ? MAX32 : info.size)
      .u16(name.length)
      .u16(extraBytes.length ? extraBytes.length + 4 : 0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u32(0)
      .u32(farOffset ? MAX32 : localOffset)
      .bytes(name);
    if (extraBytes.length) central.u16(0x0001).u16(extraBytes.length).bytes(extraBytes);
    this.central.push(central.build());
  }

  /** Writes the central directory and the end records. The caller closes the sink. */
  async finish(): Promise<void> {
    const dirOffset = this.offset;
    for (const c of this.central) await this.put(c);
    const dirSize = this.offset - dirOffset;
    const count = this.central.length;
    const big = this.options.forceZip64 === true || count >= MAX16 || dirOffset >= MAX32 || dirSize >= MAX32;
    if (big) {
      const end64 = this.offset;
      await this.put(new Builder().u32(0x06064b50).u64(44).u16(45).u16(45).u32(0).u32(0).u64(count).u64(count).u64(dirSize).u64(dirOffset).build());
      await this.put(new Builder().u32(0x07064b50).u32(0).u64(end64).u32(1).build());
    }
    await this.put(
      new Builder()
        .u32(0x06054b50)
        .u16(0)
        .u16(0)
        .u16(big ? MAX16 : count)
        .u16(big ? MAX16 : count)
        .u32(big ? MAX32 : dirSize)
        .u32(big ? MAX32 : dirOffset)
        .u16(0)
        .build(),
    );
  }
}
