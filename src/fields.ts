import { clip, type RunContext } from './context';
import { fieldWords } from './formulas';
import { isFragment } from './links';
import type { FieldMarks } from './plan';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { unknownEntity } from './xml';

/** How much of a field code is kept, counted from its first word, which names the field. */
const MAX_FIELD_CODE = 4096;

/**
 * Reports a Word field by its code, and tells whether any of its findings strips it. A field whose name holds the
 * content of an external entity has a type the file does not say, as one whose name another field builds does. A
 * HYPERLINK field's result is the text the reader sees, and goes with its tooltip into the label check.
 */
export function evaluateField(ctx: RunContext, location: string, code: string, computed: boolean, result = ''): boolean {
  const data = { code: clip(code.trim()) };
  const words = fieldWords(code);
  const first = words[0]?.text ?? '';
  const entity = unknownEntity(first);
  const actions: string[] = [];
  if (computed || entity !== undefined) actions.push(ctx.report(C.Field, D.Computed, location, entity === undefined ? data : { ...data, entity: clip(entity) }).action);
  const name = first.toUpperCase();
  // An argument comes before any switch, and a switch is an unquoted word starting with a backslash.
  const switchAt = words.findIndex(w => !w.quoted && w.text.startsWith('\\'));
  const arg = (n: number) => (words[n] !== undefined && (switchAt < 0 || n < switchAt) ? words[n].text : undefined);
  if (name === 'DDE' || name === 'DDEAUTO') actions.push(ctx.report(C.Dde, D.Field, location, data).action);
  else if (name === 'MACROBUTTON') actions.push(ctx.report(C.Field, D.Macrobutton, location, { ...data, macro: clip(words[1]?.text ?? '') }).action);
  else if (name === 'INCLUDETEXT' || name === 'INCLUDEPICTURE' || name === 'IMPORT' || name === 'LINK') {
    const target = arg(name === 'LINK' ? 2 : 1);
    if (target === undefined) actions.push(ctx.report(C.External, D.Include, location, data).action);
    else actions.push(ctx.external(D.Include, target, location, data, 'field').action);
  } else if (name === 'HYPERLINK') {
    const target = arg(1);
    if (target !== undefined && !isFragment(target)) {
      const r = ctx.link(target, location, { ...data, source: 'HYPERLINK field' });
      actions.push(r.action);
      // The \o switch gives the tooltip.
      const tip = words.findIndex(w => !w.quoted && w.text.toLowerCase() === '\\o');
      const labels = [result, tip >= 0 ? words[tip + 1]?.text : undefined];
      if (r.action !== 'strip') {
        const m = ctx.labelMismatch(target, labels, location, { ...data, source: 'HYPERLINK field' });
        if (m) actions.push(m.action);
      }
    }
  }
  return actions.includes('strip');
}

/**
 * Word field codes, followed through nested fields as the text of a document arrives: a field begins, its code
 * comes, a separator starts its result, and it ends. Each format says where those marks are.
 */
/** The most of a field's result that is kept, for the label check. */
const MAX_RESULT = 4096;

interface Field {
  code: string;
  inCode: boolean;
  computed: boolean;
  marks: FieldMarks;
  /** The text of the field's result, as far as it is kept. */
  result: string;
}

export class FieldStack {
  private readonly stack: Field[] = [];
  /**
   * `onStrip` hears of each field a strip finding names, with the marks the format gave its begin, separator and end,
   * so the rewrite can find the field again.
   */
  constructor(
    private readonly ctx: RunContext,
    private readonly location: string,
    private readonly onStrip?: (marks: FieldMarks) => void,
  ) {}

  /** A field starts. One that starts before its parent's code has a word in it makes that code at run time. */
  begin(mark = -1): void {
    this.nest();
    this.stack.push({ code: '', inCode: true, computed: false, marks: { begin: mark, simple: false }, result: '' });
  }

  /** A field whose code is given whole, such as OOXML's fldSimple, whose result follows until simpleEnd(). */
  simpleBegin(code: string, mark = -1): void {
    this.nest();
    this.stack.push({ code, inCode: false, computed: false, marks: { begin: mark, simple: true }, result: '' });
  }

  simpleEnd(): void {
    const top = this.stack[this.stack.length - 1];
    if (!top?.marks.simple) return;
    this.stack.pop();
    this.decide(top);
  }

  separate(mark = -1): void {
    const top = this.stack[this.stack.length - 1];
    if (!top) return;
    top.inCode = false;
    top.marks.separate ??= mark;
  }

  end(mark = -1): void {
    const f = this.stack.pop();
    if (!f) return;
    f.marks.end = mark;
    this.decide(f);
  }

  private decide(f: Field): void {
    if (evaluateField(this.ctx, this.location, f.code, f.computed, f.result)) this.onStrip?.(f.marks);
  }

  /** Whether a field's code is being read now. */
  get inCode(): boolean {
    return this.stack[this.stack.length - 1]?.inCode === true;
  }

  /** Text of the innermost field's code. White space before the first word counts for nothing, so no amount of it can push the field's name out. */
  text(t: string): void {
    const top = this.stack[this.stack.length - 1];
    if (top?.inCode && top.code.length < MAX_FIELD_CODE) top.code = top.code ? top.code + t : t.trimStart();
  }

  /** Text of the innermost field's result, the text the reader sees. */
  result(t: string): void {
    const top = this.stack[this.stack.length - 1];
    if (top && !top.inCode && top.result.length < MAX_RESULT) top.result += t;
  }

  /** A document can end inside a field. Word still reads the code it has. */
  finish(): void {
    for (let f = this.stack.pop(); f; f = this.stack.pop()) this.decide(f);
  }

  private nest(): void {
    const top = this.stack[this.stack.length - 1];
    if (top?.inCode && top.code === '') top.computed = true;
  }
}
