import { type BiffRecord, objSubrecords, R, RecordReader } from './biff';
import type { CfbEntry, CompoundFile } from './cfb';
import type { ByteSink } from './types';

/** What the rewrite of a BIFF8 workbook stream changes, each by the offset of the record it changes. */
export interface WorkbookEdits {
  /** BoundSheet8 records whose sheet becomes an empty worksheet of the same name and visibility. */
  emptySheets: Set<number>;
  /** Lbl records renamed to the name given, no longer a built-in name. */
  names: Map<number, string>;
  /** Obj records whose FtMacro, the macro a click runs, goes. */
  objMacros: Set<number>;
  /** SupBook records whose path becomes a reference to the workbook itself. */
  selfBooks: Set<number>;
  /** HLink records that go, with the HLinkTooltip record after one. */
  hyperlinks: Set<number>;
  /** The workbook is rewritten only to decrypt it. Any rewrite of an encrypted workbook decrypts it. */
  decrypt: boolean;
}

export const noEdits = (): WorkbookEdits => ({ emptySheets: new Set(), names: new Map(), objMacros: new Set(), selfBooks: new Set(), hyperlinks: new Set(), decrypt: false });

export const hasEdits = (e: WorkbookEdits) => e.decrypt || e.emptySheets.size > 0 || e.names.size > 0 || e.objMacros.size > 0 || e.selfBooks.size > 0 || e.hyperlinks.size > 0;

/** A rewrite that cannot keep the stream's offsets right, which leaves the stream as it was. */
export class RewriteError extends Error {}

const T = { index: 0x020b, dbcell: 0x00d7, extsst: 0x00ff, defColWidth: 0x0055, dimensions: 0x0200, window2: 0x023e, hlinkTooltip: 0x0800 };
/** Records that hold stream positions: BoundSheet8, Index, DBCell and ExtSST. */
const POINTERS = new Set([R.boundSheet, T.index, T.dbcell, T.extsst]);
/** BIFF8's largest record body. */
const MAX_BODY = 8224;

const u16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
const put32 = (b: Uint8Array, at: number, v: number) => {
  b[at] = v & 0xff;
  b[at + 1] = (v >>> 8) & 0xff;
  b[at + 2] = (v >>> 16) & 0xff;
  b[at + 3] = (v >>> 24) & 0xff;
};
const record = (type: number, body: Uint8Array) => ({ type, body, old: -1 });

/** A record of the output: the stream record it copies, by index, with the body that replaces its own when it changed. */
interface Out {
  type: number;
  /** The record's index in the stream, or -1 for a record the rewrite writes itself. */
  source: number;
  /** Where the record starts in the stream as read, for the records whose position others hold. */
  old: number;
  length: number;
  body?: Uint8Array;
}

/**
 * Writes a workbook stream again into `sink`, record by record in the order the stream holds them, with the edits
 * made. Records change size, so the stream positions other records hold are written anew: BoundSheet8's lbPlyPos,
 * Index's ibXF and rgibRw, DBCell's dbRtrw and ExtSST's ib. A position that does not fall where a record of the stream
 * starts, or inside one that keeps its bytes, throws RewriteError before anything is written. `decrypt` turns each
 * record after FILEPASS into its plain bytes. Bytes after the last whole record are not written.
 *
 * The stream is read twice: once for the record headers and the few bodies the edits change, and once to copy the
 * rest, so the rewrite holds the records' headers, not their bytes.
 */
export async function rewriteWorkbook(
  cfb: CompoundFile,
  stream: CfbEntry,
  edits: WorkbookEdits,
  biff8: boolean,
  decrypt: ((r: BiffRecord) => Uint8Array) | undefined,
  sink: ByteSink,
  deadline?: number,
): Promise<void> {
  const heads: Array<{ type: number; offset: number; length: number }> = [];
  /** The plain bodies the first pass needs: each BOF, the records that hold stream positions, and those the edits change. */
  const bodies = new Map<number, Uint8Array>();
  let filePass = -1;
  const plainOf = (r: BiffRecord, i: number) => Uint8Array.from(filePass >= 0 && i > filePass && decrypt ? decrypt(r) : r.body);
  const reader = new RecordReader(cfb, stream, 0, deadline);
  for (let r = await reader.next(); r; r = await reader.next()) {
    const i = heads.length;
    heads.push({ type: r.type, offset: r.offset, length: r.body.length });
    // office-defuse does not encrypt, so an encrypted workbook is written without FILEPASS, every record decrypted.
    if (r.type === R.filePass && filePass < 0) {
      if (!decrypt) throw new RewriteError('an encrypted workbook without its key');
      filePass = i;
      continue;
    }
    if (r.type === R.bof || POINTERS.has(r.type) || edits.names.has(r.offset) || edits.objMacros.has(r.offset) || edits.selfBooks.has(r.offset)) bodies.set(r.offset, plainOf(r, i));
  }
  if (heads[0]?.type !== R.bof) throw new RewriteError('the stream does not start with a BOF record');
  const indexOf = new Map(heads.map((h, i) => [h.offset, i]));

  // Each substream, from a BOF at the top level to the EOF that closes it, by the offset of its BOF.
  const ends = new Map<number, number>();
  let depth = 0;
  let open = -1;
  heads.forEach((h, i) => {
    if (h.type === R.bof) {
      if (depth === 0) open = i;
      depth++;
    } else if (h.type === R.eof && depth > 0 && --depth === 0) ends.set(heads[open].offset, i);
  });
  const empty = new Map<number, number>();
  for (const at of edits.emptySheets) {
    const body = bodies.get(at);
    if (!body || heads[indexOf.get(at) ?? -1]?.type !== R.boundSheet) throw new RewriteError('an edit names no BoundSheet record');
    const start = u32(body, 0);
    const end = ends.get(start);
    if (end === undefined) throw new RewriteError('a sheet whose position is not where a substream starts');
    empty.set(start, end);
  }

  const out: Out[] = [];
  for (let i = 0; i < heads.length; i++) {
    const h = heads[i];
    if (i === filePass) continue;
    if (edits.hyperlinks.has(h.offset)) {
      if (heads[i + 1]?.type === T.hlinkTooltip) i++;
      continue;
    }
    const end = empty.get(h.offset);
    if (end !== undefined) {
      emptyWorksheet(bodies.get(h.offset) ?? new Uint8Array(0), biff8).forEach((x, k) => {
        out.push({ type: x.type, source: -1, old: k === 0 ? h.offset : -1, length: x.body.length, body: x.body });
      });
      i = end;
      continue;
    }
    let body = bodies.get(h.offset);
    let changed = false;
    if (body && h.type === R.boundSheet && edits.emptySheets.has(h.offset)) {
      body = Uint8Array.from(body);
      body[5] = 0;
      changed = true;
    }
    const name = edits.names.get(h.offset);
    if (body && name !== undefined) [body, changed] = [renamed(body, name, biff8), true];
    if (body && edits.objMacros.has(h.offset)) [body, changed] = [withoutMacro(body), true];
    if (body && edits.selfBooks.has(h.offset)) [body, changed] = [selfBook(body), true];
    const length = changed && body ? body.length : h.length;
    if (length > MAX_BODY) throw new RewriteError('a record past 8,224 bytes');
    out.push({ type: h.type, source: i, old: h.offset, length, body: changed ? body : undefined });
  }

  // New offsets, and the old offset of each record kept as it was, so positions can be followed.
  const moved = new Map<number, number>();
  const kept: Array<{ old: number; now: number; length: number }> = [];
  let pos = 0;
  for (const x of out) {
    if (x.old >= 0) {
      moved.set(x.old, pos);
      if (x.source >= 0 && !x.body) kept.push({ old: x.old, now: pos, length: 4 + x.length });
    }
    pos += 4 + x.length;
  }
  const start = (old: number) => {
    const now = moved.get(old);
    if (now === undefined) throw new RewriteError(`a stream position, ${old}, where no record starts`);
    return now;
  };
  // Kept records are in stream order, so the one around a position is found by halving.
  const inside = (old: number) => {
    let lo = 0;
    let hi = kept.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (kept[mid].old <= old) lo = mid;
      else hi = mid - 1;
    }
    const k = kept[lo];
    if (!k || old < k.old || old >= k.old + k.length) throw new RewriteError(`a stream position, ${old}, inside no record the rewrite keeps as it was`);
    return k.now + (old - k.old);
  };
  const patched = (x: Out, b: Uint8Array, at: number) => {
    if (x.type === R.boundSheet && b.length >= 4) put32(b, 0, start(u32(b, 0)));
    else if (x.type === T.index && b.length >= (biff8 ? 16 : 12)) {
      // BIFF8 writes rwMic and rwMac in 4 bytes each, BIFF5 in 2.
      const ib = biff8 ? 12 : 8;
      if (u32(b, ib)) put32(b, ib, start(u32(b, ib)));
      for (let k = ib + 4; k + 4 <= b.length; k += 4) put32(b, k, start(u32(b, k)));
    } else if (x.type === T.dbcell && b.length >= 4 && x.old >= 0) put32(b, 0, at - start(x.old - u32(b, 0)));
    else if (x.type === T.extsst) for (let k = 2; k + 8 <= b.length; k += 8) put32(b, k, inside(u32(b, k)));
    return b;
  };
  // The records that hold positions are written anew here, so a position that does not resolve fails before the
  // second pass writes anything.
  for (const x of out) {
    if (x.source < 0 || !POINTERS.has(x.type)) continue;
    const base = x.body ?? bodies.get(x.old);
    if (!base) throw new RewriteError('a record that holds stream positions was not read');
    x.body = patched(x, Uint8Array.from(base), moved.get(x.old) ?? 0);
  }

  const again = new RecordReader(cfb, stream, 0, deadline);
  let index = -1;
  let current: BiffRecord | undefined;
  for (const x of out) {
    let b: Uint8Array;
    if (x.body) b = Uint8Array.from(x.body);
    else {
      while (index < x.source) {
        current = await again.next();
        index++;
        if (!current) throw new RewriteError('the stream changed while it was read');
      }
      b = plainOf(current as BiffRecord, index);
    }
    await sink.write(Uint8Array.of(x.type & 0xff, x.type >> 8, b.length & 0xff, b.length >> 8));
    await sink.write(b);
  }
}

/**
 * The substream of an empty worksheet, from a sheet's BOF: the BOF saying worksheet, a default column width, empty
 * dimensions, a window and EOF.
 */
function emptyWorksheet(bof: Uint8Array, biff8: boolean): Array<{ type: number; body: Uint8Array; old: number }> {
  const head = Uint8Array.from(bof);
  if (head.length >= 4) {
    head[2] = 0x10;
    head[3] = 0x00;
  }
  // BIFF5's Dimensions and Window2 are shorter: 2-byte rows, and a window without its zoom fields.
  const window = new Uint8Array(biff8 ? 18 : 10);
  window[0] = 0xb6;
  window[6] = 0x40;
  return [record(R.bof, head), record(T.defColWidth, Uint8Array.of(8, 0)), record(T.dimensions, new Uint8Array(biff8 ? 14 : 10)), record(T.window2, window), record(R.eof, new Uint8Array(0))];
}

/**
 * An Lbl record (MS-XLS 2.4.150) under another name, which is not a built-in one. BIFF8 writes the name after a flag
 * byte, and BIFF5 writes its characters alone.
 */
function renamed(b: Uint8Array, name: string, biff8: boolean): Uint8Array {
  const cch = b[3];
  const wide = biff8 && (b[14] & 1) === 1;
  const after = 14 + (biff8 ? 1 : 0) + cch * (wide ? 2 : 1);
  if (after > b.length) throw new RewriteError('an Lbl record cut short');
  const head = Uint8Array.from(b.subarray(0, 14));
  head[0] &= ~0x20;
  head[3] = name.length;
  return Uint8Array.from([...head, ...(biff8 ? [0] : []), ...Buffer.from(name, 'latin1'), ...b.subarray(after)]);
}

/** An Obj record (MS-XLS 2.4.181) without its FtMacro subrecord. */
function withoutMacro(b: Uint8Array): Uint8Array {
  for (const { ft, at, data } of objSubrecords(b)) if (ft === 0x04) return Buffer.concat([b.subarray(0, at), b.subarray(at + 4 + data.length)]);
  throw new RewriteError('an Obj record without the FtMacro to remove');
}

/** A SupBook record (MS-XLS 2.4.271) whose VirtualPath is U+0002, the workbook itself, with its sheet names kept. */
function selfBook(b: Uint8Array): Uint8Array {
  const cch = u16(b, 2);
  const wide = (b[4] & 1) === 1;
  const after = 4 + 1 + cch * (wide ? 2 : 1);
  if (cch === 0 || cch >= 0x0401 || after > b.length) throw new RewriteError('a SupBook record without a path');
  return Uint8Array.from([b[0], b[1], 1, 0, 0, 2, ...b.subarray(after)]);
}
