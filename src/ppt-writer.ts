import { checkTime } from './errors';
import type { SpillSink } from './io';
import type { ByteSource } from './types';

/** What the rewrite of a PowerPoint Document stream changes, each record by its offset in the stream as read. */
export interface PresentationEdits {
  /** Records that go, with everything inside them. */
  remove: Set<number>;
  /** Records whose body is replaced, the header written for the new length. */
  replace: Map<number, Uint8Array>;
  /** Persist objects, by persist id, that the new persist directory leaves out. */
  dropPersist: Set<number>;
  /** The presentation is rewritten only to decrypt it. A rewrite of an encrypted presentation always decrypts it. */
  decrypt: boolean;
}

export const noPresentationEdits = (): PresentationEdits => ({ remove: new Set(), replace: new Map(), dropPersist: new Set(), decrypt: false });
export const hasPresentationEdits = (e: PresentationEdits) => e.decrypt || e.remove.size > 0 || e.replace.size > 0 || e.dropPersist.size > 0;

/** The current edit as its UserEditAtom (MS-PPT 2.3.3) gives it. */
export interface CurrentEdit {
  /** The atom's body, kept but for the offsets the rewrite writes anew. */
  body: Uint8Array;
}

const USER_EDIT = 0x0ff5;
const PERSIST_DIRECTORY = 0x1772;
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
const le32 = (n: number) => Uint8Array.of(n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff);

function head(verInst: number, type: number, length: number): Uint8Array {
  return Uint8Array.of(verInst & 0xff, verInst >> 8, type & 0xff, type >> 8, ...le32(length));
}

/**
 * The bytes of one record, from `offset` in `source`, with the edits inside it made. A container is written from its
 * children, so its length follows what they become; a record without edits inside is copied as it is.
 */
async function recordBytes(source: ByteSource, offset: number, end: number, edits: PresentationEdits, touched: (from: number, to: number) => boolean, deadline?: number): Promise<Uint8Array[]> {
  checkTime(deadline);
  const h = await source.read(offset, 8);
  const verInst = h[0] | (h[1] << 8);
  const type = h[2] | (h[3] << 8);
  const length = Math.min(u32(h, 4), end - offset - 8);
  if (edits.remove.has(offset)) return [];
  const replaced = edits.replace.get(offset);
  if (replaced) return [head(verInst, type, replaced.length), replaced];
  const bodyEnd = offset + 8 + length;
  if ((verInst & 0x0f) !== 0x0f || !touched(offset + 8, bodyEnd)) return [head(verInst, type, length), await source.read(offset + 8, length)];
  const parts: Uint8Array[] = [];
  let total = 0;
  for (let pos = offset + 8; pos + 8 <= bodyEnd; ) {
    const kh = await source.read(pos, 8);
    const kidEnd = Math.min(bodyEnd, pos + 8 + u32(kh, 4));
    for (const p of await recordBytes(source, pos, kidEnd, edits, touched, deadline)) {
      parts.push(p);
      total += p.length;
    }
    pos = kidEnd;
  }
  return [head(verInst, type, total), ...parts];
}

/**
 * Writes a PowerPoint Document stream again as a full save (MS-PPT 2.1.2): each persist object the current edit
 * reaches, in persist id order, then one persist directory for them and one UserEditAtom with no edit before it.
 * `source` is the stream as read, decrypted when it was encrypted, and `persist` maps each persist id to the offset of
 * its object. Returns the offset of the new UserEditAtom, which the Current User stream points at.
 */
export async function rewritePresentation(
  source: ByteSource,
  size: number,
  persist: ReadonlyMap<number, number>,
  current: CurrentEdit,
  edits: PresentationEdits,
  sink: SpillSink,
  deadline?: number,
): Promise<number> {
  const marks = [...edits.remove, ...edits.replace.keys()].sort((a, b) => a - b);
  const touched = (from: number, to: number) => {
    let lo = 0;
    let hi = marks.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (marks[mid] < from) lo = mid + 1;
      else hi = mid;
    }
    return lo < marks.length && marks[lo] < to;
  };
  const written = new Map<number, number>();
  let pos = 0;
  for (const [id, offset] of [...persist].sort((a, b) => a[0] - b[0])) {
    if (edits.dropPersist.has(id) || offset + 8 > size) continue;
    const parts = await recordBytes(source, offset, size, edits, touched, deadline);
    if (!parts.length) continue;
    written.set(id, pos);
    for (const p of parts) {
      await sink.write(p);
      pos += p.length;
    }
  }
  // The directory: runs of consecutive ids, each a 20-bit first id and a 12-bit count, then their offsets.
  const ids = [...written.keys()].sort((a, b) => a - b);
  const dir: number[] = [];
  for (let i = 0; i < ids.length; ) {
    let n = 1;
    while (i + n < ids.length && ids[i + n] === ids[i] + n && n < 0xfff) n++;
    dir.push(...le32((ids[i] & 0xfffff) | (n << 20)));
    for (let k = 0; k < n; k++) dir.push(...le32(written.get(ids[i + k]) ?? 0));
    i += n;
  }
  const dirAt = pos;
  const dirRecord = Uint8Array.from([...head(0, PERSIST_DIRECTORY, dir.length), ...dir]);
  await sink.write(dirRecord);
  pos += dirRecord.length;
  // The UserEditAtom: no edit before it, the new directory, and no encryption session once decrypted.
  const body = Uint8Array.from(current.body.subarray(0, edits.decrypt ? 28 : current.body.length));
  body.set(le32(0), 8);
  body.set(le32(dirAt), 12);
  const editAt = pos;
  await sink.write(Uint8Array.from([...head(0, USER_EDIT, body.length), ...body]));
  return editAt;
}
