import type { CfbEntry, CompoundFile } from './cfb';
import { checkTime } from './errors';
import { Reader } from './io';
import { latin1, utf16 } from './oleobj';

/** Record types this reader acts on (MS-XLS 2.3). */
export const R = {
  bof: 0x0809,
  eof: 0x000a,
  filePass: 0x002f,
  continue: 0x003c,
  lbl: 0x0018,
  externName: 0x0023,
  boundSheet: 0x0085,
  template: 0x0060,
  addIn: 0x0087,
  supBook: 0x01ae,
  externSheet: 0x0017,
  obj: 0x005d,
  hlink: 0x01b8,
};

export interface BiffRecord {
  type: number;
  /** Where the record header starts in the stream. */
  offset: number;
  body: Uint8Array;
}

/** A sheet as BoundSheet8 (or BIFF5's BOUNDSHEET) lists it. */
export interface Sheet {
  position: number;
  /** 0 visible, 1 hidden, 2 very hidden. */
  state: number;
  /** 0 worksheet or dialog, 1 macro sheet, 2 chart, 6 VBA module. */
  type: number;
  name: string;
}

/**
 * Reads the records of a workbook stream one at a time, from `start` to the end of the stream. A record's body may
 * run past the stream's end only by being cut short, which stops the reading as Excel stops.
 */
export class RecordReader {
  private readonly reader: Reader;
  private pos: number;
  constructor(
    cfb: CompoundFile,
    readonly stream: CfbEntry,
    start: number,
    private readonly deadline?: number,
  ) {
    this.reader = new Reader(cfb.source(stream), stream.size, 16384, 8);
    this.pos = start;
  }

  async next(): Promise<BiffRecord | undefined> {
    if (this.pos + 4 > this.stream.size) return undefined;
    const h = await this.reader.read(this.pos, 4);
    const type = h[0] | (h[1] << 8);
    const length = h[2] | (h[3] << 8);
    if (this.pos + 4 + length > this.stream.size) return undefined;
    const offset = this.pos;
    const body = await this.reader.read(this.pos + 4, length);
    this.pos += 4 + length;
    if ((offset & 0xffff) < 4) checkTime(this.deadline);
    return { type, offset, body };
  }
}

/** BIFF8 writes strings with a flag byte: 8-bit characters, or 16-bit ones when its low bit is set. */
export function xlString(b: Uint8Array, at: number, chars: number): { text: string; next: number } {
  if (at >= b.length) return { text: '', next: at };
  const wide = (b[at] & 1) === 1;
  const start = at + 1;
  const end = Math.min(b.length, start + chars * (wide ? 2 : 1));
  return { text: wide ? utf16(b.subarray(start, end)) : latin1(b.subarray(start, end)), next: end };
}

/** A BoundSheet record, in BIFF8 or BIFF5. */
export function parseBoundSheet(body: Uint8Array, biff8: boolean): Sheet {
  const position = (body[0] | (body[1] << 8) | (body[2] << 16) | (body[3] << 24)) >>> 0;
  const state = body[4] & 3;
  const type = body[5];
  const cch = body[6] ?? 0;
  const name = biff8 ? xlString(body, 7, cch).text : latin1(body.subarray(7, 7 + cch));
  return { position, state, type, name };
}

/** Built-in names Excel runs as macros (MS-XLS 2.5.112): Auto_Open and Auto_Activate as the workbook opens, Auto_Close and Auto_Deactivate as it closes. */
const BUILTIN_NAMES: Record<number, string> = { 1: 'Auto_Open', 2: 'Auto_Close', 10: 'Auto_Activate', 11: 'Auto_Deactivate' };

/** A defined name, in BIFF8 or BIFF5, with a built-in name spelled out. */
export function parseLbl(body: Uint8Array, biff8: boolean): { name: string; builtin: boolean; hidden: boolean } {
  const flags = body[0] | (body[1] << 8);
  const cch = body[3] ?? 0;
  const builtin = (flags & 0x20) !== 0;
  const name = biff8 ? xlString(body, 14, cch).text : latin1(body.subarray(14, 14 + cch));
  const shown = builtin && name.length > 0 ? (BUILTIN_NAMES[name.charCodeAt(0)] ?? `built-in name ${name.charCodeAt(0)}`) : name;
  return { name: shown, builtin, hidden: (flags & 1) !== 0 };
}

/** The event a defined name runs a macro on, when Excel runs it on one: any name that starts with Auto_Open and so on. */
export function autoName(name: string): { event: string; onOpen: boolean } | undefined {
  const m = /^(?:_xlnm\.)?(auto_open|auto_close|auto_activate|auto_deactivate)/i.exec(name);
  if (!m) return undefined;
  const event = m[1].toLowerCase();
  return { event: `${event[0].toUpperCase()}${event.slice(1, 5)}${event[5].toUpperCase()}${event.slice(6)}`, onOpen: event === 'auto_open' || event === 'auto_activate' };
}
