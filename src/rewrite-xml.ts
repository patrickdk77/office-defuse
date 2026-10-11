import { TextDecoder } from 'node:util';
import type { ElementEdit, FieldMarks, PartPlan } from './plan';
import { attributeKey } from './plan';
import { detectEncoding, UNKNOWN, type XmlDoctype, type XmlElement, type XmlEncoding, type XmlExternalReference, type XmlHandler, type XmlLimits, XmlTokenizer } from './xml';

/** A range of a part's decoded text that the rewrite leaves out, or replaces with `text`. */
export interface Edit {
  start: number;
  end: number;
  text?: string;
}

/** What the rewrite of one part does besides the plan: edits that follow from what was dropped elsewhere. */
export interface RewriteRules {
  /** The edit an element gets, from the plan or from a reference to something that goes. */
  element(el: XmlElement): ElementEdit | undefined;
  /** Whether an element goes once every element in it went, as a list that would otherwise be left empty. */
  collapses?(el: XmlElement): boolean;
  /** Text to add before the root element's end tag, asked for when it is reached. */
  append?(): string;
  /** Whether an element belongs to WordprocessingML, whose fields the plan names. */
  word?(el: XmlElement): boolean;
}

export interface PartRewrite {
  edits: Edit[];
  /** The plan removes the root element, so the part goes instead. */
  rootRemoved: boolean;
}

/** Word elements that a field's code holds whole, which go when the field does. */
const IN_CODE = new Set(['r', 'fldSimple', 'hyperlink', 'smartTag']);
/** Word elements that make up a field's code and marks. */
const CODE_MARKS = new Set(['instrText', 'delInstrText', 'fldChar']);

const escapeText = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#13;');
const escapeAttribute = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;').replace(/\t/g, '&#9;').replace(/\n/g, '&#10;').replace(/\r/g, '&#13;');

/** A start tag written out from what the tokenizer read, with the edit's attributes left out or changed. */
export function startTag(el: XmlElement, edit?: ElementEdit): string {
  let s = `<${el.qname}`;
  for (const [prefix, uri] of el.namespaces) s += ` ${prefix ? `xmlns:${prefix}` : 'xmlns'}="${escapeAttribute(uri)}"`;
  for (const a of el.attributes) {
    if (edit?.strip?.has(attributeKey(a.ns, a.local))) continue;
    s += ` ${a.qname}="${escapeAttribute(edit?.set?.get(a.qname) ?? a.value)}"`;
  }
  return s + (el.empty ? '/>' : '>');
}

/** Two edits of one element as one. Removing it covers everything else. */
function merge(a: ElementEdit | undefined, b: ElementEdit | undefined): ElementEdit | undefined {
  if (!a || !b) return a ?? b;
  if (a.remove || b.remove) return { remove: true };
  return {
    unwrap: a.unwrap || b.unwrap,
    strip: new Set([...(a.strip ?? []), ...(b.strip ?? [])]),
    set: new Map([...(a.set ?? []), ...(b.set ?? [])]),
    content: b.content ?? a.content,
  };
}

interface Frame {
  el: XmlElement;
  edit?: ElementEdit;
  /** The element goes with everything in it. */
  dropped: boolean;
  /** The element began inside an entity reference being written out. */
  expanded: boolean;
  /** The removed fields whose code holds the element so far. It goes if any still does at its end. */
  regions?: Set<FieldMarks>;
  kept: number;
  removed: number;
  /** While an entity is written out: the start tag and what the element holds. */
  head?: string;
  body?: string[];
}

/**
 * Finds the edits of one part as the tokenizer reads it: the elements the plan names and the elements that refer to
 * something that goes, the Word fields the plan names, and with an internal subset that goes, the document type
 * declaration, every entity reference expanded and every start tag whose attributes the declarations changed.
 */
class Rewriter implements XmlHandler {
  readonly edits: Edit[] = [];
  rootRemoved = false;
  private readonly stack: Frame[] = [];
  /** The depth of the element being skipped, 0 when none is. */
  private skipping = 0;
  /** The entity reference being written out, at the top level of the part. */
  private expansion?: { start: number; end: number; out: string[]; changed: boolean };
  /** A reference inside a removed element, whose text needs no writing out. */
  private ignored = false;
  private readonly marks = new Map<number, { field: FieldMarks; role: 'begin' | 'close' | 'other' }>();
  private readonly active = new Set<FieldMarks>();
  /**
   * Whether the internal subset goes. References to external entities that stay need their declarations, so then the
   * subset stays, and the check of the output reports it.
   */
  private readonly subset: boolean;

  constructor(
    private readonly plan: PartPlan | undefined,
    private readonly rules: RewriteRules,
  ) {
    this.subset = plan?.subset === true && !(plan.hasExternal && !plan.external);
    for (const f of plan?.fields ?? []) {
      if (f.simple) {
        this.marks.set(f.begin, { field: f, role: 'begin' });
        continue;
      }
      // The code runs from the begin mark to the separator, or to the end mark when there is none.
      const close = f.separate ?? f.end;
      this.marks.set(f.begin, { field: f, role: 'begin' });
      if (f.separate !== undefined) this.marks.set(f.separate, { field: f, role: f.separate === close ? 'close' : 'other' });
      if (f.end !== undefined && !this.marks.has(f.end)) this.marks.set(f.end, { field: f, role: f.end === close ? 'close' : 'other' });
    }
  }

  private get top(): Frame | undefined {
    return this.stack[this.stack.length - 1];
  }

  /** Where text written out from an entity goes: the element it is in, or the reference itself. */
  private sink(): string[] | undefined {
    const top = this.top;
    if (top?.expanded) return top.body;
    return this.expansion?.out;
  }

  private fieldEdit(el: XmlElement, edit: ElementEdit | undefined): ElementEdit | undefined {
    const mark = this.marks.get(el.index);
    if (mark) {
      if (mark.role === 'close') {
        // The separator ends the code. Elements still open hold it, so they are not wholly code.
        this.active.delete(mark.field);
        for (const f of this.stack) f.regions?.delete(mark.field);
      }
      return mark.field.simple ? { ...edit, unwrap: true } : { remove: true };
    }
    if (this.active.size && this.rules.word?.(el) && CODE_MARKS.has(el.local)) return { remove: true };
    return edit;
  }

  start(el: XmlElement): void {
    const parent = this.top;
    if (this.skipping) {
      this.stack.push({ el, dropped: true, expanded: false, kept: 0, removed: 0 });
      return;
    }
    const edit = this.fieldEdit(el, merge(this.plan?.elements.get(el.index), this.rules.element(el)));
    const expanded = this.expansion !== undefined;
    const frame: Frame = { el, edit, dropped: edit?.remove === true, expanded, kept: 0, removed: 0 };
    if (!frame.dropped && this.active.size && this.rules.word?.(el) && IN_CODE.has(el.local)) frame.regions = new Set(this.active);
    const mark = this.marks.get(el.index);
    if (mark?.role === 'begin' && !mark.field.simple) this.active.add(mark.field);
    if (parent) {
      if (frame.dropped) parent.removed++;
      else parent.kept++;
    }
    this.stack.push(frame);
    if (frame.dropped) {
      if (el.depth === 1 && !expanded) this.rootRemoved = true;
      this.skipping = el.depth;
      if (this.expansion) this.expansion.changed = true;
      return;
    }
    if (this.expansion) {
      frame.head = edit?.unwrap ? '' : startTag(el, edit);
      frame.body = [];
      if (edit) this.expansion.changed = true;
    } else if (edit?.unwrap) this.edits.push({ start: el.start, end: el.end });
    else if (edit?.strip?.size || edit?.set?.size || (this.subset && el.rewritten)) this.edits.push({ start: el.start, end: el.end, text: startTag(el, edit) });
  }

  end(el: XmlElement, start: number, end: number): void {
    const frame = this.stack.pop();
    if (!frame) return;
    const parent = this.top;
    if (frame.dropped) {
      if (this.skipping !== el.depth) return;
      this.skipping = 0;
      if (!frame.expanded) this.edits.push({ start: el.start, end });
      return;
    }
    const cut = (frame.regions?.size ?? 0) > 0 || (frame.kept === 0 && frame.removed > 0 && this.rules.collapses?.(el) === true);
    if (cut) {
      if (parent) {
        parent.kept--;
        parent.removed++;
      }
      if (frame.expanded) {
        if (this.expansion) this.expansion.changed = true;
      } else if (el.depth === 1) this.rootRemoved = true;
      else this.edits.push({ start: el.start, end });
      return;
    }
    if (frame.expanded) {
      const body = frame.edit?.content !== undefined && !frame.edit.unwrap ? escapeText(frame.edit.content) : (frame.body ?? []).join('');
      const text = (frame.head ?? '') + body + (el.empty || frame.edit?.unwrap ? '' : `</${el.qname}>`);
      this.sink()?.push(text);
      return;
    }
    if (el.depth === 1 && !el.empty) {
      const added = this.rules.append?.();
      if (added) this.edits.push({ start, end: start, text: added });
    }
    if (frame.edit?.unwrap && !el.empty) this.edits.push({ start, end });
    // New content goes between the tags. Text in the element, which the edit replaces, is never written out.
    if (frame.edit?.content !== undefined && !el.empty && !frame.edit.unwrap) this.edits.push({ start: el.end, end: start, text: escapeText(frame.edit.content) });
  }

  text(t: string): void {
    if (this.skipping || !this.expansion) return;
    let s = escapeText(t);
    if (s.includes(UNKNOWN)) {
      // A reference to an external entity inside the text being written out.
      const drop = this.plan?.external === true;
      if (drop) this.expansion.changed = true;
      s = s.replace(new RegExp(`${UNKNOWN}([^${UNKNOWN}]*)${UNKNOWN}`, 'g'), (_m, name: string) => (drop ? '' : `&${name};`));
    }
    this.sink()?.push(s);
  }

  doctype(d: XmlDoctype): void {
    if ((this.subset && d.internalSubset) || (this.plan?.doctype && !d.internalSubset)) this.edits.push({ start: d.start, end: d.end });
  }

  externalEntity(r: XmlExternalReference): void {
    if (this.skipping || this.expansion || !this.plan?.external) return;
    this.edits.push({ start: r.start, end: r.end });
  }

  referenceStart(start: number, end: number): void {
    if (this.skipping) {
      this.ignored = true;
      return;
    }
    this.expansion = { start, end, out: [], changed: false };
  }

  referenceEnd(): void {
    if (this.ignored) {
      this.ignored = false;
      return;
    }
    const x = this.expansion;
    this.expansion = undefined;
    // A reference is written out as its text when the subset goes, or when something in its text goes.
    if (x && (this.subset || x.changed)) this.edits.push({ start: x.start, end: x.end, text: x.out.join('') });
  }
}

/**
 * Reads a part again with the tokenizer and returns the edits its rewrite makes. The part was read once already, by
 * the same tokenizer with the same limits, so this reading sees the same elements in the same order.
 */
export async function partEdits(chunks: AsyncIterable<Uint8Array>, part: string, limits: XmlLimits, vml: boolean, plan: PartPlan | undefined, rules: RewriteRules): Promise<PartRewrite> {
  const r = new Rewriter(plan, rules);
  const t = new XmlTokenizer(r, limits, part, {
    vml,
    notes: { doctype: d => r.doctype(d), externalEntity: e => r.externalEntity(e), referenceStart: (s, e) => r.referenceStart(s, e), referenceEnd: () => r.referenceEnd() },
  });
  for await (const c of chunks) t.write(c);
  t.end();
  return { edits: r.edits, rootRemoved: r.rootRemoved };
}

/** Sorts edits and drops each one that lies inside another, so they apply in one pass. */
function normalize(edits: Edit[]): Edit[] {
  const sorted = [...edits].sort((a, b) => a.start - b.start || Number(b.start === b.end) - Number(a.start === a.end) || b.end - a.end);
  const out: Edit[] = [];
  let reach = 0;
  for (const e of sorted) {
    if (e.start < reach) continue;
    out.push(e);
    reach = Math.max(reach, e.end);
  }
  return out;
}

/** Encodes text in the part's own encoding. */
function encode(s: string, encoding: XmlEncoding): Uint8Array {
  if (encoding === 'utf-8') return Buffer.from(s, 'utf8');
  const b = Buffer.from(s, 'utf16le');
  return encoding === 'utf-16be' ? b.swap16() : b;
}

/**
 * Applies edits to a part as its bytes arrive, and yields the result in the part's encoding, its byte order mark
 * included. Text the edits do not touch keeps its characters, so it keeps its bytes.
 */
export async function* splice(chunks: AsyncIterable<Uint8Array>, edits: Edit[]): AsyncGenerator<Uint8Array> {
  const list = normalize(edits);
  let k = 0;
  let skipTo = 0;
  let pos = 0;
  let head: Uint8Array = new Uint8Array(0);
  let decoder: TextDecoder | undefined;
  let encoding: XmlEncoding = 'utf-8';
  // The decoder hands over whole characters, and edits start and end at markup, so no piece ends inside a pair.
  const apply = (text: string, last: boolean): Uint8Array => {
    let out = '';
    let i = 0;
    const base = pos;
    while (i < text.length || (last && k < list.length && list[k].start === base + i)) {
      const at = base + i;
      if (at < skipTo) {
        const to = Math.min(text.length, skipTo - base);
        if (to <= i) break;
        i = to;
        continue;
      }
      const next = list[k];
      if (next && next.start === at) {
        out += next.text ?? '';
        skipTo = next.end;
        k++;
        continue;
      }
      const stop = next ? Math.min(text.length, next.start - base) : text.length;
      if (stop <= i) break;
      out += text.slice(i, stop);
      i = stop;
    }
    pos += text.length;
    return encode(out, encoding);
  };
  const begin = (): Uint8Array[] => {
    const found = detectEncoding(head) ?? { encoding: 'utf-8' as const, bom: 0 };
    encoding = found.encoding;
    decoder = new TextDecoder(encoding, { fatal: true, ignoreBOM: true });
    const bom = head.subarray(0, found.bom);
    const rest = head.subarray(found.bom);
    head = new Uint8Array(0);
    return [bom, apply(decoder.decode(rest, { stream: true }), false)];
  };
  for await (const c of chunks) {
    if (!decoder) {
      head = head.length ? Buffer.concat([head, c]) : c;
      if (head.length < 4) continue;
      for (const b of begin()) if (b.length) yield b;
      continue;
    }
    const b = apply(decoder.decode(c, { stream: true }), false);
    if (b.length) yield b;
  }
  if (!decoder) for (const b of begin()) if (b.length) yield b;
  const tail = apply((decoder as TextDecoder).decode(new Uint8Array(0)), true);
  if (tail.length) yield tail;
}
