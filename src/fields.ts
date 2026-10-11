import { clip, type RunContext } from './context';
import { fieldWords } from './formulas';
import { isFragment } from './links';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { unknownEntity } from './xml';

/** How much of a field code is kept, counted from its first word, which names the field. */
const MAX_FIELD_CODE = 4096;

/**
 * Reports a Word field by its code. A field whose name holds the content of an external entity has a type the file
 * does not say, as one whose name another field builds does.
 */
export function evaluateField(ctx: RunContext, location: string, code: string, computed: boolean): void {
  const data = { code: clip(code.trim()) };
  const words = fieldWords(code);
  const first = words[0]?.text ?? '';
  const entity = unknownEntity(first);
  if (computed || entity !== undefined) ctx.report(C.Field, D.Computed, location, entity === undefined ? data : { ...data, entity: clip(entity) });
  const name = first.toUpperCase();
  // An argument comes before any switch, and a switch is an unquoted word starting with a backslash.
  const switchAt = words.findIndex(w => !w.quoted && w.text.startsWith('\\'));
  const arg = (n: number) => (words[n] !== undefined && (switchAt < 0 || n < switchAt) ? words[n].text : undefined);
  if (name === 'DDE' || name === 'DDEAUTO') ctx.report(C.Dde, D.Field, location, data);
  else if (name === 'MACROBUTTON') ctx.report(C.Field, D.Macrobutton, location, { ...data, macro: clip(words[1]?.text ?? '') });
  else if (name === 'INCLUDETEXT' || name === 'INCLUDEPICTURE' || name === 'IMPORT' || name === 'LINK') {
    const target = arg(name === 'LINK' ? 2 : 1);
    if (target === undefined) ctx.report(C.External, D.Include, location, data);
    else ctx.external(D.Include, target, location, data, 'field');
  } else if (name === 'HYPERLINK') {
    const target = arg(1);
    if (target !== undefined && !isFragment(target)) ctx.link(target, location, { ...data, source: 'HYPERLINK field' });
  }
}

/**
 * Word field codes, followed through nested fields as the text of a document arrives: a field begins, its code
 * comes, a separator starts its result, and it ends. Each format says where those marks are.
 */
export class FieldStack {
  private readonly stack: Array<{ code: string; inCode: boolean; computed: boolean }> = [];
  constructor(
    private readonly ctx: RunContext,
    private readonly location: string,
  ) {}

  /** A field starts. One that starts before its parent's code has a word in it makes that code at run time. */
  begin(): void {
    this.nest();
    this.stack.push({ code: '', inCode: true, computed: false });
  }

  /** A field that is all code, such as OOXML's fldSimple. */
  simple(code: string): void {
    this.nest();
    evaluateField(this.ctx, this.location, code, false);
  }

  separate(): void {
    const top = this.stack[this.stack.length - 1];
    if (top) top.inCode = false;
  }

  end(): void {
    const f = this.stack.pop();
    if (f) evaluateField(this.ctx, this.location, f.code, f.computed);
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

  /** A document can end inside a field. Word still reads the code it has. */
  finish(): void {
    for (let f = this.stack.pop(); f; f = this.stack.pop()) evaluateField(this.ctx, this.location, f.code, f.computed);
  }

  private nest(): void {
    const top = this.stack[this.stack.length - 1];
    if (top?.inCode && top.code === '') top.computed = true;
  }
}
