import * as fsp from 'node:fs/promises';
import type { ByteSource } from './types';

/** A source over bytes already in memory. Reads return views, not copies. */
export function bufferSource(bytes: Uint8Array): ByteSource {
  return {
    size: async () => bytes.length,
    read: async (offset, length) => bytes.subarray(Math.max(0, offset), Math.min(bytes.length, offset + length)),
  };
}

/** A source over a local file, read at offsets through one file handle. */
export function fileSource(filePath: string): ByteSource {
  // The pending open is shared, so concurrent first calls get one handle and close() closes it.
  let opening: Promise<fsp.FileHandle> | undefined;
  let size: number | undefined;
  const open = () => {
    opening ??= fsp.open(filePath, 'r');
    return opening;
  };
  const getSize = async () => {
    size ??= (await (await open()).stat()).size;
    return size;
  };
  return {
    size: getSize,
    async read(offset, length) {
      const h = await open();
      const total = await getSize();
      const n = Math.max(0, Math.min(length, total - offset));
      const buf = Buffer.allocUnsafe(n);
      let done = 0;
      while (done < n) {
        const { bytesRead } = await h.read(buf, done, n - done, offset + done);
        if (bytesRead === 0) break;
        done += bytesRead;
      }
      return buf.subarray(0, done);
    },
    async close() {
      const p = opening;
      opening = undefined;
      await p?.then(
        h => h.close(),
        () => {
          // An open that failed left nothing to close.
        },
      );
    },
  };
}

/**
 * Random-access reader with a small block cache, so the ZIP reader can ask for small overlapping windows cheaply. Its
 * reads span at most a few blocks: a header, a name and its extra field, or the end of the file.
 */
export class Reader {
  /** Blocks by index, oldest first. */
  private readonly cache = new Map<number, Uint8Array>();
  constructor(
    readonly source: ByteSource,
    readonly size: number,
    private readonly blockSize = 65536,
    private readonly maxBlocks = 16,
  ) {}

  private async block(index: number): Promise<Uint8Array> {
    const hit = this.cache.get(index);
    if (hit) return hit;
    const data = await this.source.read(index * this.blockSize, Math.min(this.blockSize, this.size - index * this.blockSize));
    if (this.cache.size >= this.maxBlocks) {
      for (const oldest of this.cache.keys()) {
        this.cache.delete(oldest);
        break;
      }
    }
    this.cache.set(index, data);
    return data;
  }

  /** Reads up to `length` bytes at `offset`; shorter only at the end of the source. */
  async read(offset: number, length: number): Promise<Uint8Array> {
    const end = Math.min(this.size, offset + length);
    if (end <= offset) return new Uint8Array(0);
    const first = Math.floor(offset / this.blockSize);
    const last = Math.floor((end - 1) / this.blockSize);
    const out = new Uint8Array(end - offset);
    for (let i = first; i <= last; i++) {
      const b = await this.block(i);
      const from = i === first ? offset - first * this.blockSize : 0;
      const to = i === last ? end - i * this.blockSize : b.length;
      out.set(b.subarray(from, to), i === first ? 0 : i * this.blockSize - offset);
    }
    return out;
  }

  /** Yields the range in chunks without caching them. */
  async *chunks(offset: number, length: number, chunkSize = 262144): AsyncGenerator<Uint8Array> {
    const end = Math.min(this.size, offset + length);
    for (let pos = offset; pos < end; pos += chunkSize) yield this.source.read(pos, Math.min(chunkSize, end - pos));
  }
}
