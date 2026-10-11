import type { CfbEntry, CompoundFile } from './cfb';
import { clip, type RunContext } from './context';
import { checkTime, StopError } from './errors';
import { latin1, utf16 } from './oleobj';
import { OfficeCategory as C, OfficeDetail as D } from './types';

/*
 * Visio 2000 (version 6) and Visio 2002 to 2010 (version 11) binary drawings. Microsoft has not published this format.
 * Each fact below names the public source it comes from:
 *
 * - libvisio, src/lib/VSDParser.cpp: the version byte at 0x1A (VisioDocument.cpp), the trailer pointer at 0x24, the
 *   pointer and pointer list layout, the stream formats, the chunk header and trailer rules (getChunkHeader; v6 in
 *   VSD6Parser.cpp), the compression (VSDInternalStream.cpp), and the PolylineTo and NURBSTo cell blocks.
 * - Apache POI, poi-scratchpad/src/main/resources/org/apache/poi/hdgf/chunks_parse_cmds.tbl: where the cell blocks
 *   start in each chunk type, for Visio 2000 (type 11 lines) and 2003 (type 21 lines), counted from the chunk header.
 * - OLEToy, oletoy/vsdblock.py (re-lab): the cell block header, the formula tokens and their lengths, and the
 *   function numbers (fnames7ab). oletoy/vsdchunks.py gives the TextField and FrgnType block starts.
 * - VBA-Docs, api/Visio.viscellindices.md: the cell numbers of the Events, Actions and User-defined Cells sections.
 *
 * Checked against the .vsd test files of libvisio and POI: every chunk of the types below ends its cell blocks exactly
 * where the chunk ends, every formula block tokenizes to its exact length, and the cells that call these functions
 * hold what Visio's own stencils put there, such as RUNADDONWARGS("AEC","/INITIALIZE=") in EventDrop.
 */

const STREAM = 'VisioDocument';
const u16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;

/** The functions the .vsdx check looks for, by their number in a formula (OLEToy fnames7ab). */
const CALLS = new Map([
  [0x4f, 'DOOLEVERB'],
  [0x51, 'RUNADDON'],
  [0x79, 'RUNADDONWARGS'],
  [0x7b, 'OPENFILE'],
  [0x9b, 'CALLTHIS'],
  [0xc7, 'RUNMACRO'],
]);
/** The arguments the fixed-argument form of each takes, for showing the formula. */
const ARITY = new Map([
  [0x4f, 1],
  [0x51, 1],
  [0x79, 2],
  [0x7b, 1],
]);

/**
 * Where the cell blocks start in each chunk type, counted from the end of the 19-byte chunk header. From POI's table,
 * less 19, unless noted. POI lists no Visio 2003 start for some types, and the test files then show the Visio 2000
 * start for them too (marked "files").
 */
const V6 = new Map([
  [0x84, 17],
  [0x85, 35],
  [0x86, 25],
  [0x87, 71],
  [0x88, 7],
  [0x89, 3], // files
  [0x8a, 20],
  [0x8b, 20],
  [0x8c, 29],
  [0x8d, 38],
  [0x8f, 56],
  [0x90, 56],
  [0x92, 131],
  [0x93, 7],
  [0x94, 35],
  [0x95, 73],
  [0x96, 29],
  [0x97, 120],
  [0x98, 62], // vsdchunks.py FrgnType
  [0x99, 48],
  [0x9b, 69],
  [0x9c, 69],
  [0x9d, 38],
  [0x9e, 56],
  [0x9f, 56],
  [0xa0, 24],
  [0xa1, 36], // vsdchunks.py TextField
  [0xa2, 41],
  [0xa3, 30],
  [0xa4, 23],
  [0xa5, 46],
  [0xa6, 28],
  [0xa7, 6],
  [0xa8, 33],
  [0xa9, 28],
  [0xaa, 47],
  [0xb4, 15],
  [0xb5, 689],
  [0xb6, 36],
  [0xb7, 79],
  [0xb9, 48],
  [0xba, 60],
  [0xbc, 43], // files
  [0xbd, 45],
  [0xbe, 25],
  [0xbf, 38],
  [0xc0, 144],
  [0xc1, 29], // VSDParser.cpp readPolylineTo
  [0xc3, 61],
  [0xc4, 46],
  [0xc5, 38],
  [0xc8, 72],
]);
const V11 = new Map([
  [0x84, 17],
  [0x85, 35],
  [0x86, 61],
  [0x87, 92],
  [0x88, 7],
  [0x89, 3],
  [0x8a, 20],
  [0x8b, 20],
  [0x8c, 29], // files
  [0x8f, 56], // files
  [0x90, 56], // files
  [0x92, 132], // files, one more than Visio 2000
  [0x93, 7],
  [0x94, 88],
  [0x95, 123],
  [0x96, 29], // files
  [0x98, 62], // vsdchunks.py FrgnType, and files
  [0x99, 48],
  [0x9b, 69],
  [0x9c, 69],
  [0x9d, 38],
  [0x9e, 56], // files
  [0x9f, 60],
  [0xa0, 24],
  [0xa1, 54], // vsdchunks.py TextField, and files
  [0xa2, 41],
  [0xa3, 30],
  [0xa4, 45],
  [0xa5, 46], // files
  [0xa6, 28], // files
  [0xa7, 6],
  [0xa8, 33],
  [0xa9, 57],
  [0xaa, 47],
  [0xb4, 15],
  [0xb6, 36], // files
  [0xb7, 79],
  [0xb9, 48],
  [0xbc, 43], // files
  [0xbd, 45],
  [0xbe, 25],
  [0xbf, 38],
  [0xc0, 144],
  [0xc1, 29], // VSDParser.cpp readPolylineTo
  [0xc3, 61],
  [0xc4, 46], // files
  [0xc6, 61], // files
  [0xc5, 38],
  [0xc8, 72],
]);

/** Names for the findings: the ShapeSheet section of each chunk type that VSDDocumentStructure.h and POI identify. */
const SECTIONS = new Map([
  [0x84, 'Events'],
  [0x85, 'Line'],
  [0x86, 'Fill'],
  [0x87, 'Text Block'],
  [0x89, 'Geometry'],
  [0x8a, 'MoveTo'],
  [0x8b, 'LineTo'],
  [0x8c, 'ArcTo'],
  [0x8d, 'InfiniteLine'],
  [0x8f, 'Ellipse'],
  [0x90, 'EllipticalArcTo'],
  [0x92, 'Page Properties'],
  [0x93, 'Style Properties'],
  [0x94, 'Character'],
  [0x95, 'Paragraph'],
  [0x98, 'Foreign Image Info'],
  [0x99, 'Connection Points'],
  [0x9b, 'Shape Transform'],
  [0x9c, 'Text Transform'],
  [0x9d, '1-D Endpoints'],
  [0x9e, 'Scratch'],
  [0x9f, 'Alignment'],
  [0xa0, 'Protection'],
  [0xa1, 'Text Fields'],
  [0xa2, 'Controls'],
  [0xa3, 'Help'],
  [0xa4, 'Miscellaneous'],
  [0xa5, 'SplineStart'],
  [0xa6, 'SplineKnot'],
  [0xa7, 'Layer Membership'],
  [0xa8, 'Layers'],
  [0xa9, 'Actions'],
  [0xaa, 'Controls'],
  [0xb4, 'User-defined Cells'],
  [0xb6, 'Shape Data'],
  [0xb7, 'Ruler & Grid'],
  [0xb9, 'Connection Points'],
  [0xba, 'Connection Points'],
  [0xbc, 'Document Properties'],
  [0xbd, 'Image Properties'],
  [0xbe, 'Group Properties'],
  [0xbf, 'Shape Layout'],
  [0xc0, 'Page Layout'],
  [0xc1, 'PolylineTo'],
  [0xc3, 'NURBSTo'],
  [0xc4, 'Hyperlinks'],
  [0xc5, 'Reviewer'],
  [0xc6, 'Annotation'],
  [0xc7, 'Smart Tags'],
  [0xc8, 'Print Properties'],
]);

/** Cell names by section and cell number (VBA-Docs, VisCellIndices). */
const CELLS = new Map<number, Map<number, string>>([
  [
    0x84,
    new Map([
      [1, 'TheText'],
      [2, 'EventDblClick'],
      [3, 'EventXFMod'],
      [4, 'EventDrop'],
      [22, 'EventMultiDrop'],
    ]),
  ],
  [
    0xa9,
    new Map([
      [0, 'Menu'],
      [3, 'Action'],
      [4, 'Checked'],
      [5, 'Disabled'],
      [6, 'ReadOnly'],
      [7, 'Invisible'],
      [8, 'BeginGroup'],
      [14, 'TagName'],
      [15, 'ButtonFace'],
      [16, 'SortKey'],
    ]),
  ],
  [
    0xb4,
    new Map([
      [0, 'Value'],
      [1, 'Prompt'],
    ]),
  ],
]);

/** Operators, by token (OLEToy sl_opnames), for showing a formula. */
const BINARY = new Map([
  [0x03, '+'],
  [0x04, '-'],
  [0x05, '*'],
  [0x06, '/'],
  [0x07, '^'],
  [0x08, '&'],
  [0x09, '<'],
  [0x0a, '<='],
  [0x0b, '='],
  [0x0c, '>='],
  [0x0d, '>'],
  [0x0e, '<>'],
]);

const COLORS = 0x16;

interface Pointer {
  type: number;
  offset: number;
  length: number;
  format: number;
}

interface Token {
  kind: 'number' | 'string' | 'operator' | 'call' | 'other' | 'end';
  value?: string | number;
  /** For a call: its function number, its argument count when the token gives one, and where the token starts. */
  fn?: number;
  args?: number;
  at?: number;
}

/** What one stream holds: where its call tokens start, and the span and type of each chunk. */
interface Seen {
  calls: Set<number>;
  chunks: Array<[number, number, number]>;
}

/**
 * Reads the ShapeSheet cells of a Visio 2000 to 2010 drawing for formulas that call RUNADDON, RUNADDONWARGS,
 * CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB. Other versions of the stream go unread. A stream whose structure does
 * not read stops the run as PACKAGE/UNREADABLE.
 */
export async function visioRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry): Promise<void> {
  const doc = await cfb.bytes(stream);
  const version = doc.length > 0x1a ? doc[0x1a] : 0;
  if (version !== 6 && version !== 11) return;
  new VisioReader(ctx, doc, version).run();
}

class VisioReader {
  private readonly blocks: Map<number, number>;
  private readonly seen = new Set<string>();
  /** Bytes read or decompressed so far, against limits.totalDecompressedBytes. The same bytes read twice count twice. */
  private work = 0;

  constructor(
    private readonly ctx: RunContext,
    private readonly doc: Uint8Array,
    private readonly version: number,
  ) {
    this.blocks = version === 11 ? V11 : V6;
  }

  private fail(reason: string): never {
    throw new StopError(C.Package, D.Unreadable, `The ${STREAM} stream does not read: ${reason}`, { stream: STREAM, reason });
  }

  run(): void {
    if (this.doc.length < 0x24 + 18) this.fail('its header is cut short');
    // The trailer stream holds the pointers to every other stream, and is a pointer list itself.
    const lists: Array<{ data: Uint8Array; shift: number }> = [this.open(pointerAt(this.doc, 0x24))];
    const streams: Pointer[] = [];
    // A work list, not recursion: a stream of pointers adds the streams it names.
    while (lists.length > 0 || streams.length > 0) {
      checkTime(this.ctx.settings.deadline);
      const list = lists.pop();
      if (list) {
        this.pointers(list.data, list.shift, streams);
        continue;
      }
      const p = streams.pop();
      if (!p) break;
      const kind = p.format >> 4;
      const blob = kind === 0 || kind === 4 || kind === 5;
      if (!blob && kind !== 8 && kind !== 0xc && kind !== 0xd) continue;
      // The same bytes read the same way once. A blob reads by its type, so the type is part of the key.
      const key = `${p.offset}:${p.length}:${p.format}:${blob ? p.type : ''}`;
      if (this.seen.has(key)) continue;
      this.seen.add(key);
      const s = this.open(p);
      const seen: Seen = { calls: new Set(), chunks: [] };
      if (blob) {
        const start = Math.min(s.shift, s.data.length);
        seen.chunks.push([start, s.data.length, p.type]);
        this.section(p.type, s.data, start, s.data.length, seen);
        if (kind === 5 && p.type !== COLORS) lists.push(s);
      } else this.chunks(s.data, seen);
      this.sweep(s.data, seen);
    }
  }

  /**
   * A stream's bytes, decompressed when its format says so. A compressed blob or pointer list starts 4 bytes in, and a
   * compressed chunk list at its first byte (VSDParser.cpp handleStream).
   */
  private open(p: Pointer): { data: Uint8Array; shift: number } {
    if (p.offset + p.length > this.doc.length) this.fail(`a pointer names bytes ${p.offset} to ${p.offset + p.length}, past the end of the stream`);
    const raw = this.doc.subarray(p.offset, p.offset + p.length);
    if ((p.format & 2) === 0) {
      this.count(raw.length);
      return { data: raw, shift: 0 };
    }
    const data = inflate(raw, this.ctx.settings.decompressedBytes);
    this.count(raw.length + data.length);
    return { data, shift: 4 };
  }

  private count(n: number): void {
    this.work += n;
    const total = this.ctx.settings.totalDecompressedBytes;
    if (this.work > total) throw new StopError(C.Limit, D.TotalDecompressedSize, `The ${STREAM} stream decompresses past ${total} bytes`, { limit: total });
  }

  /** A pointer list: an offset to its header, which counts the 18-byte pointers that follow it. */
  private pointers(data: Uint8Array, shift: number, out: Pointer[]): void {
    if (shift + 4 > data.length) this.fail('a pointer list is cut short');
    const at = u32(data, shift) + shift - 4;
    if (at < 0 || at + 12 > data.length) this.fail('a pointer list header lies outside its stream');
    const n = u32(data, at + 4);
    if (n > (data.length - at - 12) / 18) this.fail(`a pointer list counts ${n} pointers, more than its stream holds`);
    for (let i = 0; i < n; i++) {
      if (i % 65536 === 0) checkTime(this.ctx.settings.deadline);
      const p = pointerAt(data, at + 12 + i * 18);
      if (p.type !== 0) out.push(p);
    }
  }

  /** A list of chunks, each a 19-byte header, its data, and a trailer whose length its type and flags give. */
  private chunks(data: Uint8Array, seen: Seen): void {
    let pos = 0;
    while (pos < data.length) {
      checkTime(this.ctx.settings.deadline);
      // Readers skip the zeros between chunks.
      while (pos < data.length && data[pos] === 0) pos++;
      if (pos >= data.length) return;
      if (pos + 19 > data.length) this.fail('a chunk header is cut short');
      const type = u32(data, pos);
      const list = u32(data, pos + 8);
      const length = u32(data, pos + 12);
      const level = u16(data, pos + 16);
      const flags = data[pos + 18];
      const start = pos + 19;
      if (length > data.length - start) this.fail(`a chunk of type 0x${hex(type)} runs past its stream`);
      seen.chunks.push([start, start + length, type]);
      this.section(type, data, start, start + length, seen);
      pos = start + length + (this.version === 11 ? trailer11(type, list, level, flags) : trailer6(type, list));
    }
  }

  /**
   * One chunk's data, bytes `from` to `to` of its stream. In a chunk type with ShapeSheet cells, the cell blocks follow
   * the fixed fields and run to the end of the data, or to four or more zeros.
   */
  private section(type: number, data: Uint8Array, from: number, to: number, seen: Seen): void {
    const start = this.blocks.get(type);
    if (start === undefined) return;
    if (to - from < start) this.fail(`a chunk of type 0x${hex(type)} is shorter than its fixed fields`);
    let pos = from + start;
    while (pos < to) {
      if (to - pos >= 4 && u32(data, pos) === 0) {
        if (data.subarray(pos, to).some(x => x !== 0)) this.fail(`a chunk of type 0x${hex(type)} holds data past the end of its cell blocks`);
        return;
      }
      if (to - pos < 6) this.fail(`a cell block in a chunk of type 0x${hex(type)} is cut short`);
      const length = u32(data, pos);
      if (length < 6 || length > to - pos) this.fail(`a cell block in a chunk of type 0x${hex(type)} runs past its chunk`);
      const kind = data[pos + 4];
      // 2 is a formula. The others hold dependencies and flags.
      if (kind === 2) this.formula(type, data[pos + 5], data, pos + 6, pos + length, seen);
      else if (kind < 1 || kind > 5) this.fail(`a cell block in a chunk of type 0x${hex(type)} is of the unknown kind ${kind}`);
      pos += length;
    }
  }

  private formula(type: number, cell: number, data: Uint8Array, from: number, to: number, seen: Seen): void {
    const tokens = this.tokens(type, data.subarray(from, to));
    const calls = tokens.filter(t => t.kind === 'call' && t.fn !== undefined && CALLS.has(t.fn));
    if (calls.length === 0) return;
    for (const t of calls) seen.calls.add(from + (t.at ?? 0));
    const section = SECTIONS.get(type) ?? `chunk type 0x${hex(type)}`;
    const name = CELLS.get(type)?.get(cell) ?? `cell ${cell}`;
    const formula = clip(show(tokens) ?? callNames(calls));
    if (type === 0x84 && name.startsWith('Event')) this.ctx.report(C.Macro, D.Event, STREAM, { event: name, formula });
    else this.ctx.report(C.Macro, D.Link, STREAM, { section, cell: name, formula });
  }

  /** A formula's tokens, in the order Visio evaluates them (OLEToy vsdblock.py get_slice). */
  private tokens(type: number, b: Uint8Array): Token[] {
    const out: Token[] = [];
    const bad = (what: string): never => this.fail(`a formula in a chunk of type 0x${hex(type)} ${what}`);
    let i = 0;
    const need = (n: number) => {
      if (i + n > b.length) bad('is cut short');
    };
    while (i < b.length) {
      const c = b[i];
      if (c === 0xfe) {
        out.push({ kind: 'end' });
        i += 1;
      } else if (c < 0x20 || c === 0xe4) {
        // An operator. 0xE4 is a pair of parentheses.
        out.push({ kind: 'operator', value: c });
        i += 1;
      } else if (c < 0x60) {
        // A number, in the unit the token names.
        need(9);
        out.push({ kind: 'number', value: Buffer.from(b.buffer, b.byteOffset + i + 1, 8).readDoubleLE(0) });
        i += 9;
      } else if (c === 0x60) {
        // A string: a count of characters, the characters, and a terminating NUL. Visio 2000 writes 8-bit characters.
        need(2);
        const n = b[i + 1];
        const width = this.version === 11 ? 2 : 1;
        need(2 + (n + 1) * width);
        const s = b.subarray(i + 2, i + 2 + n * width);
        out.push({ kind: 'string', value: width === 2 ? utf16(s) : latin1(s) });
        i += 2 + (n + 1) * width;
      } else if (c === 0x69 && this.version === 11) {
        // A string of more than 255 characters, counted in 4 bytes. Seen in a test file.
        need(5);
        const n = u32(b, i + 1);
        if (n > (b.length - i - 7) / 2) bad('holds a string longer than itself');
        out.push({ kind: 'string', value: utf16(b.subarray(i + 5, i + 5 + n * 2)) });
        i += 5 + (n + 1) * 2;
      } else if (c === 0x61 || c === 0x62) {
        // An integer of 1 or 2 bytes.
        const n = c === 0x61 ? 1 : 2;
        need(1 + n);
        out.push({ kind: 'number', value: n === 1 ? b[i + 1] : u16(b, i + 1) });
        i += 1 + n;
      } else if (REFERENCES.has(c)) {
        // A reference to a cell, of a length by its kind.
        const n = REFERENCES.get(c) ?? 0;
        need(1 + n);
        out.push({ kind: 'other' });
        i += 1 + n;
      } else if (c === 0x7a || c === 0x80) {
        // A function of fixed arguments: its number in 2 bytes, then 2 bytes of flags.
        need(5);
        out.push({ kind: 'call', fn: u16(b, i + 1), at: i });
        i += 5;
      } else if (c === 0x7b || c === 0x81) {
        // A function of variable arguments: the argument count in 4 bytes, its number in 2, and 2 bytes of flags.
        need(9);
        out.push({ kind: 'call', fn: u16(b, i + 5), args: u32(b, i + 1), at: i });
        i += 9;
      } else if (c === 0x8a) {
        // NURBS data (VSDParser.cpp readNURBSTo): 16 bytes, then 32 for each point. OLEToy counts 24 bytes more, which
        // the next token reads instead, so no token goes unread.
        need(17);
        const points = u32(b, i + 13);
        if (points > (b.length - i - 17) / 32) bad('holds NURBS data longer than itself');
        out.push({ kind: 'other' });
        i += 17 + points * 32;
      } else if (c >= 0xa0 && c <= 0xa3) {
        // AND, OR, IF and a branch, with the length of what they span in 4 bytes. What they span is read as tokens.
        need(5);
        out.push({ kind: 'other' });
        i += 5;
      } else if ((c >= 0x90 && c <= 0x93) || c === 0xf0) {
        // The marks that end an AND, OR or IF and their branches, and the filler that replaces a deleted reference.
        out.push({ kind: 'other' });
        i += 1;
      } else bad(`holds the unknown token 0x${hex(c)}`);
    }
    return out;
  }

  /**
   * Every byte of a stream searched for the token of a call to the functions above, with the flags Visio writes after
   * one, outside the formulas read. That covers chunk types whose cells this does not know where to find, and bytes a
   * chunk walk that went wrong would skip. In the test files, every such token is one a formula read found.
   */
  private sweep(b: Uint8Array, seen: Seen): void {
    const found = new Map<string, Token[]>();
    let chunk = 0;
    for (let i = 0; i + 5 <= b.length; i++) {
      const c = b[i];
      let fn: number | undefined;
      if ((c === 0x7a || c === 0x80) && b[i + 3] === 0 && b[i + 4] === 0x40) fn = u16(b, i + 1);
      else if ((c === 0x7b || c === 0x81) && i + 9 <= b.length && b[i + 7] === 0 && b[i + 8] === 0x40) fn = u16(b, i + 5);
      if (fn === undefined || !CALLS.has(fn) || seen.calls.has(i)) continue;
      // The chunks are in order, so the one that holds this byte is at or after the last one found.
      while (chunk < seen.chunks.length && seen.chunks[chunk][1] <= i) chunk++;
      const span = seen.chunks[chunk];
      const where = span && span[0] <= i ? (SECTIONS.get(span[2]) ?? `chunk type 0x${hex(span[2])}`) : 'between chunks';
      const list = found.get(where) ?? [];
      list.push({ kind: 'call', fn });
      found.set(where, list);
    }
    for (const [section, calls] of found) this.ctx.report(C.Macro, D.Link, STREAM, { section, cell: 'unknown', formula: clip(callNames(calls)) });
  }
}

/** The lengths of the reference tokens after their first byte (OLEToy sl_names70 to sl_names76, Visio 2000 and later). */
const REFERENCES = new Map([
  [0x70, 11],
  [0x72, 7],
  [0x74, 8],
  [0x75, 4],
  [0x76, 17],
]);

const hex = (n: number) => n.toString(16).padStart(2, '0');

function pointerAt(b: Uint8Array, at: number): Pointer {
  return { type: u32(b, at), offset: u32(b, at + 8), length: u32(b, at + 12), format: u16(b, at + 16) };
}

/** Chunk trailers of Visio 2002 to 2010 (VSDParser.cpp getChunkHeader). */
function trailer11(type: number, list: number, level: number, flags: number): number {
  if (type === 0x1f || type === 0xc9 || type === 0x2d || type === 0xd1) return 0;
  let n = 0;
  if (list !== 0 || [0x71, 0x70, 0x6b, 0x6a, 0x69, 0x66, 0x65, 0x2c].includes(type)) n += 8;
  if (list !== 0 || (level === 2 && flags === 0x55) || (level === 2 && flags === 0x54 && type === 0xaa) || (level === 3 && flags !== 0x50 && flags !== 0x54)) n += 4;
  if ([0x64, 0x65, 0x66, 0x69, 0x6a, 0x6b, 0x6f, 0x71, 0x92, 0xa9, 0xb4, 0xb6, 0xb9, 0xc7].includes(type) && n !== 12 && n !== 4) n += 4;
  return n;
}

/** Chunk trailers of Visio 2000 (VSD6Parser.cpp getChunkHeader). */
function trailer6(type: number, list: number): number {
  if (type === 0x1f || type === 0xc9) return 0;
  return list !== 0 || [0x76, 0x73, 0x72, 0x71, 0x70, 0x6f, 0x6e, 0x6d, 0x6c, 0x6b, 0x6a, 0x69, 0x68, 0x67, 0x66, 0x65, 0x64, 0x2c, 0x0d].includes(type) ? 8 : 0;
}

/**
 * The LZ77 compression of Visio streams (VSDInternalStream.cpp): a flag byte for each eight items, a set bit for a
 * literal byte and a clear one for a 2-byte back reference into a 4,096-byte window. Output past `limit` stops the run.
 */
export function inflate(src: Uint8Array, limit: number): Uint8Array {
  if (src.length < 2) return new Uint8Array(0);
  // An item of 2 bytes yields at most 18, so the output is at most 9 times the input.
  const out = new Uint8Array(Math.min(limit + 18, src.length * 9));
  const window = new Uint8Array(4096);
  const over = () => {
    throw new StopError(C.Limit, D.DecompressedSize, `A stream in ${STREAM} decompresses past ${limit} bytes`, { name: STREAM, limit });
  };
  let o = 0;
  let pos = 0;
  let i = 0;
  while (i < src.length) {
    const flag = src[i++];
    if (i >= src.length) break;
    for (let bit = 0; bit < 8 && i < src.length; bit++) {
      if (flag & (1 << bit)) {
        if (o >= limit) over();
        window[pos & 4095] = src[i];
        out[o++] = src[i++];
        pos++;
      } else {
        if (i + 2 > src.length) return out.subarray(0, o);
        const a = src[i++];
        const b = src[i++];
        const length = (b & 15) + 3;
        let from = ((b & 0xf0) << 4) | a;
        from = from > 4078 ? from - 4078 : from + 18;
        if (o + length > limit) over();
        for (let j = 0; j < length; j++) {
          const x = window[(from + j) & 4095];
          window[(pos + j) & 4095] = x;
          out[o++] = x;
        }
        pos += length;
      }
    }
  }
  return out.subarray(0, o);
}

const callNames = (calls: Token[]) => calls.map(t => `${CALLS.get(t.fn ?? -1) ?? ''}(...)`).join(', ');

/**
 * A formula written out, when it holds only constants, operators and the calls above. Visio stores formulas in
 * reverse Polish order. Undefined when anything else is in it, such as a reference or another function.
 */
function show(tokens: Token[]): string | undefined {
  const stack: string[] = [];
  for (const t of tokens) {
    if (t.kind === 'end') break;
    if (t.kind === 'number') stack.push(String(t.value));
    else if (t.kind === 'string') stack.push(`"${String(t.value).replace(/"/g, '""')}"`);
    else if (t.kind === 'operator') {
      const op = BINARY.get(Number(t.value));
      if (op !== undefined) {
        const right = stack.pop();
        const left = stack.pop();
        if (left === undefined || right === undefined) return undefined;
        stack.push(`${left}${op}${right}`);
      } else if (t.value === 0x13 || t.value === 0x14 || t.value === 0xe4) {
        const x = stack.pop();
        if (x === undefined) return undefined;
        stack.push(t.value === 0x13 ? `-${x}` : t.value === 0x14 ? `${x}%` : `(${x})`);
      } else return undefined;
    } else if (t.kind === 'call') {
      const name = CALLS.get(t.fn ?? -1);
      const n = t.args ?? ARITY.get(t.fn ?? -1);
      if (name === undefined || n === undefined || n > stack.length) return undefined;
      stack.push(`${name}(${stack.splice(stack.length - n, n).join(',')})`);
    } else return undefined;
  }
  return stack.length === 1 ? stack[0] : undefined;
}
