/**
 * Follows the quotes of a formula one character at a time, as its text arrives: double-quoted strings and
 * single-quoted names, where doubling the quote escapes it.
 */
export class Quotes {
  private quote = '';
  private closing = false;

  /** Takes the next character and tells whether it stands outside quotes. */
  outside(c: string): boolean {
    if (this.closing) {
      this.closing = false;
      if (c === this.quote) return false;
      this.quote = '';
    }
    if (this.quote) {
      if (c === this.quote) this.closing = true;
      return false;
    }
    if (c === '"' || c === "'") {
      this.quote = c;
      return false;
    }
    return true;
  }
}

/**
 * Functions that reach outside the file when a formula is calculated: WEBSERVICE and IMAGE load a URL, RTD starts a
 * COM server, CALL and REGISTER call into a DLL, and EXEC runs a program. A prefix such as `_xlfn.` or
 * `COM.MICROSOFT.` names the same function.
 */
const LOAD = ['WEBSERVICE', 'IMAGE'];
const RUN = ['RTD', 'CALL', 'REGISTER', 'REGISTER.ID', 'EXEC'];

/** What a function a formula calls does outside the file, or undefined when it does nothing there. */
export function riskyFunction(name: string): 'load' | 'run' | undefined {
  const n = name.toUpperCase();
  const is = (f: string) => n === f || n.endsWith(`.${f}`);
  if (LOAD.some(is)) return 'load';
  if (RUN.some(is)) return 'run';
  return undefined;
}

/** The functions, outside quotes, that a formula's text calls and that reach outside the file, as its text arrives. */
export class FunctionCalls {
  private readonly quotes = new Quotes();
  private word = '';
  private spaced = false;
  readonly found = new Map<string, 'load' | 'run'>();

  write(t: string): void {
    for (const c of t) {
      if (!this.quotes.outside(c)) {
        this.word = '';
        continue;
      }
      if (/[A-Za-z0-9_.]/.test(c)) {
        if (this.spaced) this.word = '';
        this.spaced = false;
        if (this.word.length < 64) this.word += c;
      } else if (c === ' ') this.spaced = this.word !== '';
      else {
        if (c === '(' && this.word) {
          const kind = riskyFunction(this.word);
          if (kind) this.found.set(this.word.toUpperCase(), kind);
        }
        this.word = '';
        this.spaced = false;
      }
    }
  }
}

/** The functions a formula's whole text calls that reach outside the file. */
export function formulaCalls(formula: string): Map<string, 'load' | 'run'> {
  const f = new FunctionCalls();
  f.write(formula);
  return f.found;
}

/** Built-in functions by their index in Excel's function table (MS-XLS 2.5.198.17): EXEC, REGISTER, CALL, REGISTER.ID and RTD. */
const FTAB = new Map([
  [110, 'EXEC'],
  [149, 'REGISTER'],
  [150, 'CALL'],
  [267, 'REGISTER.ID'],
  [379, 'RTD'],
]);

/** A name a BIFF formula token refers to: a defined name by index, or an external name through an XTI. */
export type BiffNameRef = { kind: 'name'; index: number } | { kind: 'namex'; ixti: number; index: number };

/**
 * The functions a BIFF8 formula (MS-XLS 2.5.198) calls that reach outside the file: built-in ones by index, and
 * add-in and newer ones by the name the formula names them with, such as `_xlfn.WEBSERVICE`. The walk stops at a
 * token it does not know the size of.
 */
export function biffCalls(rgce: Uint8Array, nameOf: (ref: BiffNameRef) => string | undefined): Map<string, 'load' | 'run'> {
  const found = new Map<string, 'load' | 'run'>();
  const u16 = (at: number) => rgce[at] | (rgce[at + 1] << 8);
  const u32 = (at: number) => (rgce[at] | (rgce[at + 1] << 8) | (rgce[at + 2] << 16) | (rgce[at + 3] << 24)) >>> 0;
  const add = (name: string | undefined) => {
    const kind = name === undefined ? undefined : riskyFunction(name);
    if (name !== undefined && kind) found.set(name.toUpperCase(), kind);
  };
  for (let i = 0; i < rgce.length; ) {
    const ptg = rgce[i++];
    if (ptg < 0x20) {
      if (ptg === 0x01 || ptg === 0x02) i += 4;
      else if (ptg === 0x17) i += 2 + (rgce[i] ?? 0) * ((rgce[i + 1] ?? 0) & 1 ? 2 : 1);
      else if (ptg === 0x19) {
        const grbit = rgce[i] ?? 0;
        const w = u16(i + 1);
        i += 3 + (grbit & 0x04 ? (w + 1) * 2 : 0);
      } else if (ptg === 0x1c || ptg === 0x1d) i += 1;
      else if (ptg === 0x1e) i += 2;
      else if (ptg === 0x1f) i += 8;
      else if (ptg < 0x03 || ptg > 0x16) break;
      continue;
    }
    switch (ptg & 0x1f) {
      case 0x00:
        i += 7;
        break;
      case 0x01:
        add(FTAB.get(u16(i)));
        i += 2;
        break;
      case 0x02:
        add(FTAB.get(u16(i + 1) & 0x7fff));
        i += 3;
        break;
      case 0x03:
        add(nameOf({ kind: 'name', index: u32(i) }));
        i += 4;
        break;
      case 0x04:
      case 0x0a:
      case 0x0c:
        i += 4;
        break;
      case 0x05:
      case 0x0b:
      case 0x0d:
        i += 8;
        break;
      case 0x06:
      case 0x07:
      case 0x08:
        i += 6;
        break;
      case 0x09:
        i += 2;
        break;
      case 0x19:
        add(nameOf({ kind: 'namex', ixti: u16(i), index: u32(i + 2) }));
        i += 6;
        break;
      case 0x1a:
      case 0x1c:
        i += 6;
        break;
      case 0x1b:
      case 0x1d:
        i += 10;
        break;
      default:
        return found;
    }
  }
  return found;
}

/** Whether an OpenDocument formula calls the DDE function. */
export function callsDde(formula: string): boolean {
  const q = new Quotes();
  for (let i = 0; i < formula.length; i++) {
    if (q.outside(formula[i]) && /^DDE\s*\(/i.test(formula.slice(i, i + 16)) && !/[\w.]/.test(formula[i - 1] ?? '')) return true;
  }
  return false;
}

/** One word of a Word field code. A switch is an unquoted word that starts with a backslash. */
export interface FieldWord {
  text: string;
  quoted: boolean;
}

/**
 * Splits a Word field code into words: quoted arguments, switches and bare words. Inside quotes a backslash escapes
 * the next character; outside them a doubled backslash stands for one, as Word writes paths.
 */
export function fieldWords(code: string): FieldWord[] {
  const out: FieldWord[] = [];
  const space = (c: string) => c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\xa0';
  let i = 0;
  while (i < code.length) {
    while (i < code.length && space(code[i])) i++;
    if (i >= code.length) break;
    let text = '';
    const quoted = code[i] === '"';
    if (quoted) {
      i++;
      while (i < code.length && code[i] !== '"') {
        if (code[i] === '\\' && i + 1 < code.length) i++;
        text += code[i++];
      }
      i++;
    } else {
      while (i < code.length && !space(code[i])) {
        if (code[i] === '\\' && code[i + 1] === '\\') i++;
        text += code[i++];
      }
    }
    out.push({ text, quoted });
  }
  return out;
}
