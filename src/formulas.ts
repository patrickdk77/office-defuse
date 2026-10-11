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
