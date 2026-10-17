import { autoName } from './biff';
import { CompoundFile, guid } from './cfb';
import { type ContainedDecision, decideContained, ownFinding } from './contained';
import { clip, type Reported, type ZipContext } from './context';
import { StopError } from './errors';
import { FieldStack } from './fields';
import { type Format, OOXML_FORMATS, XPS_FORMATS, XPS_SEQUENCE } from './formats';
import { bagProblem, controlProblem, controlStorageProblem, FORMS_CONTROLS, OOXML_CONTROL_STREAMS } from './forms';
import { FunctionCalls, Quotes } from './formulas';
import { bufferSource, SpillSink } from './io';
import { isAbsoluteTarget, isFragment } from './links';
import { compoundRules, nestedNotes } from './objects';
import { attributeKey, CompoundPlan } from './plan';
import { writeCompound } from './rewrite-cfb';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { XlsbLinks } from './xlsb';
import { attr, UNKNOWN, unknownEntity, type XmlElement, type XmlHandler } from './xml';
import { nameKey, type ZipEntry } from './zip';

export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const T = 'http://schemas.openxmlformats.org/officedocument/2006/relationships/';
const S = 'http://purl.oclc.org/ooxml/officedocument/relationships/';
const MS = 'http://schemas.microsoft.com/office/2006/relationships/';
const both = (name: string) => [T + name, S + name];

/** The relationship types that name an XPS package's FixedDocumentSequence, lowercased. */
export const XPS_MAIN = { xps: 'http://schemas.microsoft.com/xps/2005/06/fixedrepresentation', oxps: 'http://schemas.openxps.org/oxps/v1.0/fixedrepresentation' };
/** The XPS and OpenXPS markup namespaces, lowercased. */
export const XPS_NS = new Set(['http://schemas.microsoft.com/xps/2005/06', 'http://schemas.openxps.org/oxps/v1.0']);

/** Relationship types, lowercased, as OPC compares them without regard to case. */
export const REL = {
  main: new Set([...both('officedocument'), 'http://schemas.microsoft.com/visio/2010/relationships/document', XPS_MAIN.xps, XPS_MAIN.oxps]),
  hyperlink: new Set(both('hyperlink')),
  template: new Set(both('attachedtemplate')),
  image: new Set([...both('image'), 'http://schemas.microsoft.com/office/2007/relationships/hdphoto']),
  frame: new Set(both('frame')),
  subDocument: new Set(both('subdocument')),
  workbook: new Set([
    ...both('externallinkpath'),
    `${MS}xlexternallinkpath/xlpathmissing`,
    `${MS}xlexternallinkpath/xlstartup`,
    `${MS}xlexternallinkpath/xlalternatestartup`,
    `${MS}xlexternallinkpath/xllibrary`,
  ]),
};

/** Relationship types that say what the part they reach is. Each is a bit of Part.reachedAs. */
const AS = { vba: 1, macrosheet: 2, ole: 4, package: 8, control: 16, controlBinary: 32 };
const REACHED_AS = new Map<string, number>([
  [`${MS}vbaproject`, AS.vba],
  [`${MS}xlmacrosheet`, AS.macrosheet],
  [`${MS}xlintlmacrosheet`, AS.macrosheet],
  ...both('oleobject').map((t): [string, number] => [t, AS.ole]),
  ...both('package').map((t): [string, number] => [t, AS.package]),
  ...both('control').map((t): [string, number] => [t, AS.control]),
  [`${MS}activexcontrolbinary`, AS.controlBinary],
]);

/** Content types, lowercased. */
export const TYPE = {
  vba: new Set(['application/vnd.ms-office.vbaproject']),
  macrosheet: new Set(['application/vnd.ms-excel.macrosheet+xml', 'application/vnd.ms-excel.intlmacrosheet+xml', 'application/vnd.ms-excel.macrosheet', 'application/vnd.ms-excel.intlmacrosheet']),
  ole: new Set(['application/vnd.openxmlformats-officedocument.oleobject']),
  control: new Set(['application/vnd.ms-office.activex+xml']),
  controlBinary: new Set(['application/vnd.ms-office.activex']),
  vml: 'application/vnd.openxmlformats-officedocument.vmldrawing',
  xlsbExternalLink: 'application/vnd.ms-excel.externallink',
};

/** Whether the rules read a part of this content type as XML. */
export const isXmlType = (ct: string | undefined) => ct !== undefined && (ct.endsWith('+xml') || ct === 'application/xml' || ct === 'text/xml' || ct === TYPE.vml);

/**
 * Binary parts that belong to the package's own format rather than being files of their own: the records of a binary
 * workbook, printer settings, obfuscated fonts, VBA project signatures and the origin of digital signatures.
 */
const formatBinary = (ct: string) =>
  ct.startsWith('application/vnd.ms-excel.') ||
  ct.endsWith('.printersettings') ||
  ct === 'application/vnd.openxmlformats-officedocument.obfuscatedfont' ||
  ct === 'application/vnd.ms-package.obfuscated-opentype' ||
  ct.startsWith('application/vnd.ms-office.vbaprojectsignature') ||
  ct === 'application/vnd.openxmlformats-package.digital-signature-origin';

/** Relationship types that leave the package, by the finding they make. Any other is EXTERNAL/OTHER. */
const EXTERNAL_KINDS: Array<[Set<string>, D]> = [
  [REL.template, D.Template],
  [new Set([...both('oleobject'), ...both('package')]), D.Object],
  [REL.image, D.Image],
  [REL.frame, D.Frame],
  [REL.subDocument, D.Include],
  [REL.workbook, D.Workbook],
];

export const WORD_NS = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);
export const SHEET_NS = new Set(['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main']);
const XDR_NS = new Set(['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetdrawing', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetdrawing']);
export const R_NS = new Set(['http://schemas.openxmlformats.org/officedocument/2006/relationships', 'http://purl.oclc.org/ooxml/officedocument/relationships']);
const CUSTOM_UI_NS = new Set(['http://schemas.microsoft.com/office/2006/01/customui', 'http://schemas.microsoft.com/office/2009/07/customui']);
const VML_EXCEL_NS = 'urn:schemas-microsoft-com:office:excel';
const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const ACTIVEX_NS = 'http://schemas.microsoft.com/office/2006/activex';
const EXTENDED_NS = new Set(['http://schemas.openxmlformats.org/officedocument/2006/extended-properties', 'http://purl.oclc.org/ooxml/officedocument/extendedproperties']);
/**
 * The attributes of a SpreadsheetML data connection that name where it loads from (ECMA-376 18.13): an ODC file or an
 * offline cube on the connection, a connection string and command, a web query's URL, or a text file.
 */
const CONNECTION_SOURCES: Record<string, string[]> = { connection: ['odcFile', 'sourceFile'], dbPr: ['connection', 'command'], webPr: ['url'], textPr: ['sourceFile'] };
const VISIO_CODE = /\b(RUNADDON|RUNADDONWARGS|CALLTHIS|RUNMACRO|OPENFILE|DOOLEVERB)\s*\(/i;

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export interface Part {
  entry: ZipEntry;
  contentType?: string;
  /** The AS bits of the relationship types that reach the part. */
  reachedAs: number;
  /** A relationship of an ActiveX control part reaches it, so it is that control's binary. */
  claimed: boolean;
  /** The parts its own relationships reach. */
  targets: Set<Part>;
  ole: boolean;
  /** For a VBA project part read before the XML parts, what became of it. */
  vba?: ContainedDecision;
  /** For an activeX part, the control its ax:ocx element describes. */
  activeX?: ActiveXInfo;
  /** For an activeX part, what the check of its control found. */
  forms?: FormsVerdict;
  /** A binary part that holds the data of a Forms 2.0 control that stays. */
  formsData?: boolean;
}

/** What an activeX part's ax:ocx element says: the class, how the control persists itself, and a property bag. */
export interface ActiveXInfo {
  classId: string;
  persistence?: string;
  props: Array<{ name: string; value: string }>;
  /** An element below ax:ocxPr, such as a font or picture, or one ax:ocx does not hold. */
  nested: boolean;
}

/** An activeX part's control as checked: its class, and why it may not be a Forms 2.0 control as Office writes one. */
interface FormsVerdict {
  classId?: string;
  problem?: string;
}

/** What the rewrite needs of an OOXML package: its parts, which part each relationships part belongs to, and the roots. */
export interface OoxmlGraph {
  contentTypes: ZipEntry;
  parts: Map<string, Part>;
  relsOf: Map<string, ZipEntry>;
  /** The parts the package relationships reach. */
  roots: Set<Part>;
  /** The package uses the relationship types of Strict OOXML. */
  strict: boolean;
  /** The package is XPS or OpenXPS, whose parts name the parts they use in attributes. */
  xps: boolean;
}

/** A Word attribute, qualified as Word writes it or unqualified. */
const wordAttr = (el: XmlElement, local: string) => el.attributes.find(a => a.local === local && (a.ns === '' || WORD_NS.has(a.ns.toLowerCase())))?.value;

/** The key of the part a relationships part belongs to, '' for the package, or undefined for any other part. */
export const relsSource = (relsName: string): string | undefined => {
  const m = /^(.*?)_rels\/([^/]*)\.rels$/i.exec(relsName);
  return m ? nameKey(m[1] + m[2]) : undefined;
};

/**
 * Resolves an internal target against the folder of the part it belongs to, as RFC 3986 resolves a relative
 * reference: a ".." at the root of the package stays at the root. LibreOffice writes "../customXml/item1.xml" in the
 * package relationships of workbooks it saves.
 */
export function resolveTarget(relsName: string, target: string): string {
  const folder = relsName.slice(0, relsName.toLowerCase().indexOf('_rels/'));
  const path = (target.startsWith('/') ? target.slice(1) : folder + target).split('#')[0];
  const out: string[] = [];
  for (const seg of path.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return nameKey(out.join('/'));
}

/** A reference in an XPS part to another part: an absolute part name, or a path relative to the part's folder. */
export function resolvePartRef(part: string, ref: string): string {
  const folder = part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : '';
  const path = (ref.startsWith('/') ? ref.slice(1) : folder + ref).split('#')[0];
  const out: string[] = [];
  for (const seg of path.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return nameKey(out.join('/'));
}

/** The URIs an XPS ImageSource names: the value itself, or a ColorConvertedBitmap's image and color profile. */
function imageSourceUris(value: string): string[] {
  const t = value.trim();
  if (!t.startsWith('{')) return [t];
  const words = t
    .replace(/^\{|\}$/g, '')
    .trim()
    .split(/\s+/);
  return words[0] === 'ColorConvertedBitmap' ? words.slice(1) : [];
}

/**
 * The resources an XPS element names in its attributes, as the attribute and the URI: a document, page or resource
 * dictionary's Source, a font's FontUri, an image brush's image and profile, and a ContextColor's profile.
 */
export function xpsResources(el: XmlElement): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  for (const a of el.attributes) {
    if (a.ns !== '') continue;
    if (a.local === 'Source' || a.local === 'FontUri') out.push([a.local, a.value.trim()]);
    else if (a.local === 'ImageSource') for (const uri of imageSourceUris(a.value)) out.push([a.local, uri]);
    else {
      const m = /^\s*ContextColor\s+(\S+)/.exec(a.value);
      if (m) out.push([a.local, m[1]]);
    }
  }
  return out.filter(([, uri]) => uri !== '');
}

const duplicate = (ctx: ZipContext, location: string, kind: string, name: string) => ctx.report(C.Package, D.DuplicateEntry, location, { kind, name: clip(name) });

/**
 * The defaults and overrides of [Content_Types].xml. Only those that type an entry of the archive are kept, so a
 * part full of declarations costs no memory. Two that type the same entry make PACKAGE/DUPLICATE_ENTRY.
 */
async function readContentTypes(ctx: ZipContext, entry: ZipEntry): Promise<{ defaults: Map<string, string>; overrides: Map<string, string> }> {
  const keys = new Set(ctx.zip.entries.map(e => nameKey(e.name)));
  const extensions = new Set(ctx.zip.entries.map(e => /\.([^./]*)$/.exec(e.name)?.[1]?.toLowerCase() ?? ''));
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  await ctx.parse(entry, {
    start(el) {
      if (el.ns.toLowerCase() !== CONTENT_TYPES_NS) return;
      const type = attr(el, '', 'ContentType')?.trim().toLowerCase();
      if (type === undefined) return;
      if (el.local === 'Default') {
        const ext = attr(el, '', 'Extension')?.toLowerCase();
        if (ext === undefined || !extensions.has(ext)) return;
        if (defaults.has(ext)) duplicate(ctx, entry.name, 'content type default', ext);
        defaults.set(ext, type);
      } else if (el.local === 'Override') {
        const name = attr(el, '', 'PartName');
        if (name === undefined) return;
        const key = nameKey(name.replace(/^\//, ''));
        if (!keys.has(key)) return;
        if (overrides.has(key)) duplicate(ctx, entry.name, 'content type override', name);
        overrides.set(key, type);
      }
    },
  });
  return { defaults, overrides };
}

/** Reads a relationships part and hands each relationship to `each` as it is read. Ids given twice are reported. */
export async function readRelationships(ctx: ZipContext, entry: ZipEntry, each: (r: Relationship) => void): Promise<void> {
  const ids = new Set<string>();
  await ctx.parse(entry, {
    start(el) {
      if (el.ns.toLowerCase() !== RELS_NS || el.local !== 'Relationship') return;
      const id = attr(el, '', 'Id') ?? '';
      if (ids.has(id)) duplicate(ctx, entry.name, 'relationship Id', id);
      ids.add(id);
      const target = attr(el, '', 'Target');
      if (target === undefined) return;
      const mode = (attr(el, '', 'TargetMode') ?? 'Internal').trim().toLowerCase();
      each({ id, type: (attr(el, '', 'Type') ?? '').trim().toLowerCase(), target, external: mode !== 'internal' || isAbsoluteTarget(target) });
    },
  });
}

/**
 * Word field codes in WordprocessingML: fldSimple, and fldChar marks around instrText. A field a strip finding names
 * goes into the plan by the element numbers of its marks.
 */
class Fields {
  private readonly stack: FieldStack;
  private capture = false;
  /** Inside a w:t, whose text is part of a field's result when a field's code has ended. */
  private shown = false;
  constructor(ctx: ZipContext, part: string) {
    this.stack = new FieldStack(ctx, part, marks => ctx.plan.part(part).fields.push(marks));
  }

  start(el: XmlElement): void {
    if (el.local === 'fldSimple') this.stack.simpleBegin(wordAttr(el, 'instr') ?? '', el.index);
    else if (el.local === 'fldChar') {
      const kind = wordAttr(el, 'fldCharType');
      if (kind === 'begin') this.stack.begin(el.index);
      else if (kind === 'separate') this.stack.separate(el.index);
      else if (kind === 'end') this.stack.end(el.index);
    } else if (el.local === 'instrText' || el.local === 'delInstrText') this.capture = true;
    else if (el.local === 't') this.shown = true;
  }

  end(el: XmlElement): void {
    if (el.local === 'instrText' || el.local === 'delInstrText') this.capture = false;
    else if (el.local === 't') this.shown = false;
    else if (el.local === 'fldSimple') this.stack.simpleEnd();
  }

  text(t: string): void {
    if (this.capture) this.stack.text(t);
    else if (this.shown) this.stack.result(t);
  }

  finish(): void {
    this.stack.finish();
  }
}

/** Text that may run long. Detection needs only whether it holds more than white space; the report shows its start. */
class Shown {
  text = '';
  content = false;
  add(t: string): void {
    if (/\S/.test(t)) this.content = true;
    if (this.text.length < 200) this.text = this.text ? this.text + t : t.trimStart();
  }
}

/** The rules that read inside one XML part of an OOXML package. */
class PartRules implements XmlHandler {
  private readonly fields: Fields;
  /**
   * The formula being read, with a check for DDE that sees all of it however long it is, and the first external
   * entity whose content stands in it, which could hold the DDE reference or move the quotes around one.
   */
  private formula?: { shown: Shown; quotes: Quotes; ddeBar: boolean; entity?: string; index: number; calls: FunctionCalls };
  private macroText?: { shown: Shown; index: number };
  /** The part is an Excel 4.0 macro sheet, whose formulas are its macros and which its own finding covers. */
  private macroSheet = false;
  /** Word's hyperlink base, in the extended properties, which relative hyperlinks resolve against. */
  private hyperlinkBase?: { shown: Shown; index: number };
  /** The parts this XPS part names in its markup, by key, which count as reached from it. */
  readonly refs = new Set<string>();
  /** The control an activeX part describes in its ax:ocx element. */
  activeX?: ActiveXInfo;
  /**
   * The labels of the hyperlinks the part names by relationship Id: their text, tooltips and display text. The
   * relationships part comes after the part, so the targets are compared with these as it is read.
   */
  readonly labels = new Map<string, string[]>();
  /** A hyperlink whose text is being read, with the depth of the element whose end ends it. */
  private label?: { id: string; depth: number; text: string; also: string[] };
  /** The last namespace seen and its lowercase form. Neighboring elements share one, and this spares a copy each. */
  private ns = '';
  private lower = '';

  constructor(
    private readonly ctx: ZipContext,
    private readonly part: string,
  ) {
    this.fields = new Fields(ctx, part);
  }

  /** Records that the rewrite removes the element, when the finding strips. */
  private remove(r: Reported, el: { index: number }): void {
    if (r.action === 'strip') this.ctx.plan.element(this.part, el.index, { remove: true });
  }

  private lowerNs(el: XmlElement): string {
    if (el.ns !== this.ns) {
      this.ns = el.ns;
      this.lower = el.ns.toLowerCase();
    }
    return this.lower;
  }

  /** Adds labels for the hyperlink a relationship Id names. */
  private addLabels(id: string, labels: Array<string | undefined>): void {
    const list = this.labels.get(id) ?? [];
    for (const l of labels) if (l?.trim()) list.push(l);
    if (list.length) this.labels.set(id, list);
  }

  /** The labels a hyperlink element gives: the text of `textDepth`'s element, its tooltip, and an Excel display name. */
  private hyperlink(el: XmlElement, textDepth: number | undefined, also: Array<string | undefined>): void {
    const id = el.attributes.find(a => a.local === 'id' && R_NS.has(a.ns.toLowerCase()))?.value;
    if (id === undefined) return;
    if (textDepth === undefined || this.label) this.addLabels(id, also);
    else this.label = { id, depth: textDepth, text: '', also: also.filter((l): l is string => l !== undefined) };
  }

  start(el: XmlElement): void {
    const ns = this.lowerNs(el);
    // Below ax:ocx only ax:ocxPr may stand, and nothing below it: anything else marks the bag as more than plain values.
    if (this.activeX && (el.depth > 2 || (el.depth === 2 && !(ns === ACTIVEX_NS && el.local === 'ocxPr')))) this.activeX.nested = true;
    if (ns === ACTIVEX_NS) {
      const value = (local: string) => el.attributes.find(a => a.ns.toLowerCase() === ACTIVEX_NS && a.local === local)?.value;
      if (el.depth === 1 && el.local === 'ocx') this.activeX = { classId: value('classid') ?? '', persistence: value('persistence'), props: [], nested: false };
      else if (this.activeX && el.depth === 2 && el.local === 'ocxPr') this.activeX.props.push({ name: value('name') ?? '', value: value('value') ?? '' });
      else if (this.activeX) this.activeX.nested = true;
      return;
    }
    if (CUSTOM_UI_NS.has(ns)) {
      if (el.depth === 1) this.ctx.report(C.Ui, D.Ribbon, this.part);
      for (const a of el.attributes) {
        if (a.ns !== '' || !a.value.trim() || !/^(on[A-Z]|get[A-Z]|loadImage$)/.test(a.local)) continue;
        // The callback goes and the control stays, doing nothing, unless it calls a macro of a project that stays.
        if (this.ctx.macroTrigger(C.Ui, D.Callback, this.part, { callback: `${a.local}=${clip(a.value)}` }, a.value).action === 'strip') {
          this.ctx.plan.element(this.part, el.index, { strip: new Set([attributeKey('', a.local)]) });
        }
      }
      return;
    }
    if (el.depth === 1 && el.local === 'macrosheet') this.macroSheet = true;
    if (WORD_NS.has(ns)) {
      this.fields.start(el);
      if (el.local === 'hyperlink') this.hyperlink(el, el.depth, [wordAttr(el, 'tooltip')]);
    } else if (SHEET_NS.has(ns)) {
      if (el.local === 'hyperlink') this.hyperlink(el, undefined, [attr(el, '', 'display'), attr(el, '', 'tooltip')]);
      if (el.local === 'f' || el.local === 'definedName') this.formula = { shown: new Shown(), quotes: new Quotes(), ddeBar: false, index: el.index, calls: new FunctionCalls() };
      if (el.local === 'definedName') {
        // Excel runs the macro a name like _xlnm.Auto_Open refers to, as it does in a binary workbook.
        const name = attr(el, '', 'name') ?? '';
        const auto = autoName(name);
        if (auto) this.remove(this.ctx.report(C.Macro, auto.onOpen ? D.EventOnOpen : D.Event, this.part, { event: auto.event, name: clip(name) }), el);
      } else if (el.local === 'ddeLink')
        this.remove(this.ctx.report(C.Dde, D.Link, this.part, { application: clip(attr(el, '', 'ddeService') ?? ''), topic: clip(attr(el, '', 'ddeTopic') ?? '') }), el);
      else if (Object.hasOwn(CONNECTION_SOURCES, el.local)) this.connection(el);
    } else if (XDR_NS.has(ns)) {
      const macro = attr(el, '', 'macro');
      // The shape stays, and clicking it runs nothing.
      if (macro?.trim() && this.ctx.macroTrigger(C.Macro, D.Link, this.part, { macro: clip(macro) }, macro).action === 'strip') {
        this.ctx.plan.element(this.part, el.index, { strip: new Set([attributeKey('', 'macro')]) });
      }
    } else if (ns === VML_EXCEL_NS && el.local === 'FmlaMacro') this.macroText = { shown: new Shown(), index: el.index };
    else if (ns === VISIO_NS && el.local === 'Cell') this.visioCell(el);
    else if (XPS_NS.has(ns)) this.xps(el);
    else if (EXTENDED_NS.has(ns) && el.local === 'HyperlinkBase') this.hyperlinkBase = { shown: new Shown(), index: el.index };
    if (el.local === 'hlinkClick' || el.local === 'hlinkHover' || el.local === 'hlinkMouseOver') {
      this.drawingAction(el);
      // In a text run the link sits in the run's properties, and the run's text is what the reader sees.
      this.hyperlink(el, el.depth >= 2 ? el.depth - 2 : undefined, [attr(el, '', 'tooltip')]);
    }
  }

  text(t: string): void {
    if (this.label && this.label.text.length < 4096) this.label.text += t;
    this.fields.text(t);
    const f = this.formula;
    if (f) {
      f.shown.add(t);
      // Excel writes a DDE reference as application|topic!item, and "|" means nothing else outside quotes.
      for (const c of t) if (f.quotes.outside(c) && c === '|') f.ddeBar = true;
      f.calls.write(t);
      if (f.entity === undefined && t.includes(UNKNOWN)) f.entity = unknownEntity(t);
    }
    this.macroText?.shown.add(t);
    this.hyperlinkBase?.shown.add(t);
  }

  end(el: XmlElement): void {
    const ns = this.lowerNs(el);
    if (this.label && el.depth === this.label.depth) {
      this.addLabels(this.label.id, [this.label.text, ...this.label.also]);
      this.label = undefined;
    }
    if (WORD_NS.has(ns)) this.fields.end(el);
    if (SHEET_NS.has(ns) && (el.local === 'f' || el.local === 'definedName') && this.formula) {
      // A cell keeps its cached value without the formula. A defined name goes whole.
      const { ddeBar, entity, shown, calls } = this.formula;
      if (ddeBar) this.remove(this.ctx.report(C.Dde, D.Formula, this.part, { formula: clip(shown.text) }), this.formula);
      else if (entity !== undefined) this.remove(this.ctx.report(C.Dde, D.Formula, this.part, { formula: clip(shown.text), entity: clip(entity) }), this.formula);
      if (!this.macroSheet)
        for (const [fn, kind] of calls.found)
          this.remove(this.ctx.report(kind === 'fetch' ? C.External : C.Macro, D.Formula, this.part, { function: clip(fn), formula: clip(shown.text) }), this.formula);
      this.formula = undefined;
    }
    if (ns === VML_EXCEL_NS && el.local === 'FmlaMacro' && this.macroText) {
      const macro = this.macroText.shown.text.trim();
      if (this.macroText.shown.content) this.remove(this.ctx.macroTrigger(C.Macro, D.Link, this.part, { macro: clip(macro) }, macro), this.macroText);
      this.macroText = undefined;
    }
    if (this.hyperlinkBase && EXTENDED_NS.has(ns) && el.local === 'HyperlinkBase') {
      const base = this.hyperlinkBase.shown.text.trim();
      if (base && !isFragment(base)) this.remove(this.ctx.link(base, this.part, { source: 'hyperlink base' }), this.hyperlinkBase);
      this.hyperlinkBase = undefined;
    }
    if (el.depth === 1) this.fields.finish();
  }

  /**
   * A data connection's source. Each source a connection names is a finding, and the rewrite empties it, so the
   * connection stays for the query tables and pivot caches that name it and loads nothing.
   */
  private connection(el: XmlElement): void {
    const set = new Map<string, string>();
    for (const name of CONNECTION_SOURCES[el.local]) {
      const a = el.attributes.find(x => x.ns === '' && x.local === name);
      if (!a?.value.trim()) continue;
      const data = { element: clip(el.qname), attribute: name, target: clip(a.value, 2000) };
      if (this.ctx.report(C.External, D.DataConnection, this.part, data).action === 'strip') set.set(a.qname, '');
    }
    if (set.size) this.ctx.plan.element(this.part, el.index, { set });
  }

  /** PowerPoint action settings, in a DrawingML hyperlink's action attribute. The relationship it names is reported on its own. */
  private drawingAction(el: XmlElement): void {
    const written = (attr(el, '', 'action') ?? '').trim();
    const action = written.toLowerCase();
    if (!action) return;
    const id = el.attributes.find(a => a.local === 'id' && R_NS.has(a.ns.toLowerCase()))?.value;
    const data = { trigger: el.local, action: clip(written), ...(id === undefined ? {} : { relationship: clip(id) }) };
    if (action.startsWith('ppaction://program')) this.remove(this.ctx.report(C.Action, D.RunProgram, this.part, data), el);
    else if (action.startsWith('ppaction://macro')) {
      const name = /[?&]name=([^&]*)/i.exec(written)?.[1] ?? '';
      this.remove(this.ctx.macroTrigger(C.Macro, D.Link, this.part, data, safeDecode(name)), el);
    } else if (action.startsWith('ppaction://ole')) this.remove(this.ctx.report(C.Action, D.OleVerb, this.part, data), el);
  }

  private visioCell(el: XmlElement): void {
    const n = attr(el, '', 'N') ?? '';
    const formula = attr(el, '', 'F') ?? '';
    if (VISIO_CODE.test(formula)) {
      if (n.startsWith('Event')) this.remove(this.ctx.report(C.Macro, D.Event, this.part, { event: clip(n), formula: clip(formula) }), el);
      else this.remove(this.ctx.report(C.Macro, D.Link, this.part, { cell: clip(n), formula: clip(formula) }), el);
    }
    const v = attr(el, '', 'V');
    if (n === 'Address' && v?.trim() && !isFragment(v)) this.remove(this.ctx.link(v, this.part, { source: 'Visio hyperlink' }), el);
  }

  /**
   * XPS markup, which names the parts it uses in attributes rather than in relationships. A resource inside the
   * package counts as reached from this part. One outside it loads from outside, and the element that names it goes.
   * A link with a scheme or a network path gets the link checks and loses its target. A relative link moves inside
   * the document.
   */
  private xps(el: XmlElement): void {
    for (const a of el.attributes) {
      if (a.ns !== '' || (a.local !== 'FixedPage.NavigateUri' && a.local !== 'OutlineTarget')) continue;
      if (isFragment(a.value) || !isAbsoluteTarget(a.value)) continue;
      const data = { element: clip(el.qname) };
      let r = this.ctx.link(a.value, this.part, data);
      // The characters of a Glyphs element are what the reader sees.
      if (r.action !== 'strip' && el.local === 'Glyphs') r = this.ctx.labelMismatch(a.value, [attr(el, '', 'UnicodeString')], this.part, data) ?? r;
      if (r.action === 'strip') this.ctx.plan.element(this.part, el.index, { strip: new Set([attributeKey('', a.local)]) });
    }
    for (const [name, uri] of xpsResources(el)) {
      if (!isAbsoluteTarget(uri)) {
        this.refs.add(resolvePartRef(this.part, uri));
        continue;
      }
      const detail = name === 'ImageSource' ? D.Image : name === 'Source' ? D.Include : D.Other;
      this.remove(this.ctx.external(detail, uri, this.part, { element: clip(el.qname), attribute: name }, name), el);
    }
  }
}

/** An OOXML package's parts, typed by its [Content_Types].xml, and its relationships parts by the key of the part each belongs to. */
export interface OoxmlParts {
  parts: Map<string, Part>;
  relsOf: Map<string, ZipEntry>;
}

/** Reads [Content_Types].xml and types each part by it. Directory entries are neither parts nor relationships parts. */
export async function readParts(ctx: ZipContext, contentTypes: ZipEntry): Promise<OoxmlParts> {
  const { defaults, overrides } = await readContentTypes(ctx, contentTypes);
  const parts = new Map<string, Part>();
  const relsOf = new Map<string, ZipEntry>();
  for (const entry of ctx.zip.entries) {
    if (entry === contentTypes || entry.directory) continue;
    const source = relsSource(entry.name);
    if (source !== undefined) {
      relsOf.set(source, entry);
      continue;
    }
    const key = nameKey(entry.name);
    const ext = /\.([^./]*)$/.exec(entry.name)?.[1]?.toLowerCase();
    parts.set(key, { entry, contentType: overrides.get(key) ?? (ext === undefined ? undefined : defaults.get(ext)), reachedAs: 0, claimed: false, targets: new Set(), ole: false });
  }
  return { parts, relsOf };
}

/** The format the package relationships name, and the relationship types they use. */
export interface OoxmlMain {
  format: Format;
  /** The package uses the relationship types of Strict OOXML. */
  strict: boolean;
  /** The package is XPS or OpenXPS. */
  xps: boolean;
}

/**
 * Reads the package relationships and names the format by the one main part they name: by its content type, or for
 * XPS by the relationship type too, since XPS and OpenXPS share their content types. `each` sees every package
 * relationship as it is read. Throws when no main part, or more than one, is named, or its format is unknown.
 */
export async function readMain(ctx: ZipContext, { parts, relsOf }: OoxmlParts, each?: (rels: ZipEntry, r: Relationship) => void): Promise<OoxmlMain> {
  const packageRels = relsOf.get('');
  const mains: Array<Part | undefined> = [];
  let strict = false;
  let mainType = '';
  if (packageRels) {
    await readRelationships(ctx, packageRels, r => {
      if (REL.main.has(r.type)) {
        mains.push(r.external ? undefined : parts.get(resolveTarget(packageRels.name, r.target)));
        strict = r.type.startsWith(S);
        mainType = r.type;
      }
      each?.(packageRels, r);
    });
  }
  if (mains.length > 1) throw new StopError(C.Package, D.AmbiguousFormat, 'The package names more than one main part', { reason: 'more than one main part' });
  const main = mains[0];
  if (!main) throw new StopError(C.Package, D.Unrecognized, 'The package has no main part', { reason: 'no main document part' });
  const xps = mainType === XPS_MAIN.xps || mainType === XPS_MAIN.oxps;
  const format = xps ? (main.contentType === XPS_SEQUENCE ? XPS_FORMATS[mainType === XPS_MAIN.oxps ? 'oxps' : 'xps'] : undefined) : OOXML_FORMATS.get(main.contentType ?? '');
  if (!format) throw new StopError(C.Package, D.Unrecognized, 'The main part has an unknown content type', { contentType: clip(main.contentType ?? '') });
  return { format, strict, xps };
}

/** Names an OOXML package's format as its inspection does, reading only [Content_Types].xml and the package relationships. */
export async function nameOoxml(ctx: ZipContext, contentTypes: ZipEntry): Promise<Format> {
  return (await readMain(ctx, await readParts(ctx, contentTypes))).format;
}

/** The content type of an add-in's presentation part, which PowerPoint writes empty. */
const PPAM_MAIN = 'application/vnd.ms-powerpoint.addin.macroenabled.main+xml';

/**
 * Inspects an OOXML package. The package relationships come first, since they name the main part and so the format.
 * Then each part is read in archive order, followed by its relationships, which are reported as they are read and
 * leave behind only which parts they reach and how. The run names `outer` as the format when given, as for a Flat
 * OPC package unpacked into a ZIP archive, and the package's own format is returned.
 */
export async function inspectOoxml(ctx: ZipContext, contentTypes: ZipEntry, outer?: Format): Promise<Format> {
  const typed = await readParts(ctx, contentTypes);
  const { parts, relsOf } = typed;
  const plan = ctx.plan;
  const roots = new Set<Part>();
  /** The labels each part gives its hyperlinks, by part key and relationship Id. */
  const labelsOf = new Map<string, Map<string, string[]>>();
  const relate = (rels: ZipEntry, from: Part | undefined, r: Relationship) => {
    const location = `${rels.name}, ${clip(r.id)}`;
    const part = r.external ? undefined : parts.get(resolveTarget(rels.name, r.target));
    if (part) {
      part.reachedAs |= REACHED_AS.get(r.type) ?? 0;
      if (TYPE.control.has(from?.contentType ?? '')) part.claimed = true;
      (from?.targets ?? roots).add(part);
    }
    // A hyperlink to a part of the package, or to a place in the document, goes nowhere else. A relationship a strip
    // finding names goes, and so does every element that refers to it.
    let reported: Reported | undefined;
    if (REL.hyperlink.has(r.type)) {
      if (!part && !isFragment(r.target)) {
        reported = ctx.link(r.target, location);
        const labels = from ? labelsOf.get(nameKey(from.entry.name))?.get(r.id) : undefined;
        if (labels && reported.action !== 'strip') reported = ctx.labelMismatch(r.target, labels, location) ?? reported;
      }
    } else if (r.external) {
      const kind = EXTERNAL_KINDS.find(([types]) => types.has(r.type))?.[1] ?? D.Other;
      reported = ctx.external(kind, r.target, location, { type: clip(r.type) }, clip(r.type.slice(r.type.lastIndexOf('/') + 1)));
    }
    if (reported?.action === 'strip') plan.dropRelationship(rels.name, r.id);
  };

  const { format, strict, xps } = await readMain(ctx, typed, (rels, r) => relate(rels, undefined, r));
  ctx.format = outer ?? format;
  // A VBA project goes to the script plugins before the XML parts are read, whose buttons and callbacks keep running
  // its macros only when it stays.
  if (ctx.run.options.scriptPlugins !== undefined) for (const p of parts.values()) if (TYPE.vba.has(p.contentType ?? '')) p.vba = await vbaPart(ctx, p.entry, p.entry.name);

  // OPC has no folders. A directory entry is still inflated once, so its bytes are checked like any other.
  for (const entry of ctx.zip.entries) if (entry.directory) await ctx.drain(entry);
  for (const [key, p] of parts) {
    const ct = p.contentType;
    if (ct === undefined) {
      if (ctx.report(C.Package, D.UntypedPart, p.entry.name, { part: p.entry.name }).action === 'strip') plan.drop.add(key);
      p.ole = await ctx.drain(p.entry);
    } else if (format.format === 'ppam' && p.entry.size === 0 && ct.toLowerCase() === PPAM_MAIN) {
      // PowerPoint writes an add-in's presentation part empty, since an add-in has no slides.
      await ctx.drain(p.entry);
    } else if (isXmlType(ct)) {
      const rules = new PartRules(ctx, p.entry.name);
      await ctx.parse(p.entry, rules, ct === TYPE.vml);
      if (rules.labels.size) labelsOf.set(key, rules.labels);
      if (rules.activeX) p.activeX = rules.activeX;
      for (const k of rules.refs) {
        const t = parts.get(k);
        if (t && t !== p) p.targets.add(t);
      }
    } else if (ct === TYPE.xlsbExternalLink) {
      // The binary form of an Excel ddeLink.
      const links = new XlsbLinks();
      p.ole = await ctx.drain(p.entry, c => links.write(c));
      // The binary part is the link, so it goes whole.
      for (const l of links.dde) if (ctx.report(C.Dde, D.Link, p.entry.name, { application: clip(l.application), topic: clip(l.topic) }).action === 'strip') plan.drop.add(key);
    } else p.ole = await ctx.drain(p.entry);
    const rels = relsOf.get(key);
    if (rels) await readRelationships(ctx, rels, r => relate(rels, p, r));
  }
  // Relationships of parts that are not there still say what the parts they reach are.
  for (const [source, rels] of relsOf) if (source !== '' && !parts.has(source)) await readRelationships(ctx, rels, r => relate(rels, undefined, r));

  // Controls are checked before any part's own rules, since a Forms 2.0 control that stays keeps its binary part.
  for (const p of parts.values()) if (TYPE.control.has(p.contentType ?? '') || (p.reachedAs & AS.control) !== 0) p.forms = await formsControl(ctx, p);
  for (const p of parts.values()) await partRules(ctx, p);
  unreferenced(ctx, parts, relsOf, roots);
  plan.graph = { contentTypes, parts, relsOf, roots, strict, xps };
  return format;
}

/**
 * Findings about a whole part, from its content type and the relationships that reach it. The files inside the
 * package, its OLE parts and embedded packages, go to the plugins first. A plugin's verdict decides whether such a
 * part stays, and when no plugin takes it, the findings of this function do.
 */
async function partRules(ctx: ZipContext, p: Part): Promise<void> {
  const ct = p.contentType ?? '';
  const loc = p.entry.name;
  const key = nameKey(loc);
  const data = { part: loc, contentType: clip(ct) };
  const as = (bit: number) => (p.reachedAs & bit) !== 0;
  const vba = TYPE.vba.has(ct) || as(AS.vba);
  const ole = TYPE.ole.has(ct) || as(AS.ole);
  const control = TYPE.control.has(ct) || as(AS.control);
  const controlBinary = TYPE.controlBinary.has(ct) || as(AS.controlBinary);
  const pkg = as(AS.package);
  // Besides OLE parts and packages, every part that is a whole file in a format of its own, such as an image, goes
  // to the plugins. Such a part has no rule of its own, so one no plugin takes stays.
  const wholeFile = !isXmlType(p.contentType) && !formatBinary(ct);
  // The data of a Forms 2.0 control that stays was read with the control, and stays with it.
  if (p.formsData) return;
  const contained = vba || ole || pkg || controlBinary || p.ole || wholeFile;
  // With script plugins, a VBA project part is read here, so its modules go to them, rather than to the file plugins.
  const vbaDecision = p.vba ?? (vba && ctx.run.options.scriptPlugins !== undefined ? await vbaPart(ctx, p.entry, loc) : undefined);
  const decision = vbaDecision ?? (contained ? await decideContained(ctx, { name: loc, declaredType: ct || undefined, location: loc, chunks: () => ctx.zip.read(p.entry) }) : undefined);
  const found: Reported[] = [];
  // A plugin's answer is final, so a file it answered for gets none of office-defuse's own findings.
  if (vba && !vbaDecision) found.push(ownFinding(ctx, decision, C.Macro, D.VbaProject, loc, data));
  // A macro sheet becomes an empty worksheet, so the indexes of the sheets after it stay right.
  if ((TYPE.macrosheet.has(ct) || as(AS.macrosheet)) && ctx.report(C.Macro, D.XlmMacrosheet, loc, data).action === 'strip') ctx.plan.sheets.add(key);
  if (ole) found.push(ownFinding(ctx, decision, C.Embedded, D.OleObject, loc, data));
  if (pkg) found.push(ownFinding(ctx, decision, C.Embedded, D.Package, loc, data));
  if (control && p.forms && !p.forms.problem) found.push(ctx.report(C.Control, D.FormsControl, loc, { ...data, clsid: p.forms.classId ?? '' }));
  else if (control || (controlBinary && !p.claimed)) {
    const why = control && p.forms ? { ...(p.forms.classId ? { clsid: p.forms.classId } : {}), ...(p.forms.problem ? { reason: clip(p.forms.problem) } : {}) } : {};
    found.push(control ? ctx.report(C.Control, D.ActiveX, loc, { ...data, ...why }) : ownFinding(ctx, decision, C.Control, D.ActiveX, loc, { ...data, ...why }));
  }
  if (p.ole && !vba && !ole && !control && !controlBinary && !pkg) found.push(ownFinding(ctx, decision, C.Embedded, D.OleFile, loc, data));
  if (decision ? !decision.keep : found.some(f => f.action === 'strip')) ctx.plan.drop.add(key);
  else if (decision?.replacement) ctx.plan.replace.set(key, decision.replacement);
}

/** Bytes a Forms 2.0 control's binary part holds at most, pictures left out. */
const MAX_CONTROL_DATA = 65536;

/**
 * Checks the control an activeX part describes as a Forms 2.0 control that Office persisted: by its class, and by
 * its properties, read from the property bag in the part or from the binary part its relationship names. A control
 * persisted as a stream starts with its class id; one persisted as a storage holds its CompObj and contents. A
 * control that passes marks its binary part, which then stays with it.
 */
async function formsControl(ctx: ZipContext, p: Part): Promise<FormsVerdict> {
  const ax = p.activeX;
  if (!ax) return { problem: 'its part has no ax:ocx element' };
  const classId = ax.classId.replace(/[{}]/g, '').trim().toUpperCase();
  if (!FORMS_CONTROLS.has(classId)) return { classId, problem: 'its class is not a Forms 2.0 control office-defuse reads' };
  const binaries = [...p.targets].filter(t => t.claimed);
  const persistence = ax.persistence ?? 'persistStreamInit';
  if (persistence === 'persistPropertyBag') {
    if (binaries.length) return { classId, problem: 'a property bag with a binary part' };
    return { classId, problem: bagProblem(classId, ax.props, ax.nested) };
  }
  if (ax.props.length || ax.nested) return { classId, problem: `properties in its part beside ${persistence}` };
  if (binaries.length !== 1) return { classId, problem: 'it has no binary part, or more than one' };
  const bin = binaries[0];
  if (bin.entry.size > MAX_CONTROL_DATA) return { classId, problem: 'its binary part is larger than a Forms 2.0 control without pictures' };
  const chunks: Uint8Array[] = [];
  for await (const c of ctx.zip.read(bin.entry)) chunks.push(c);
  const bytes = Buffer.concat(chunks);
  let problem: string | undefined;
  if (persistence === 'persistStreamInit' || persistence === 'persistStream') {
    problem = bytes.length < 16 || guid(bytes, 0) !== classId ? 'its binary part names another class' : controlProblem(classId, bytes.subarray(16));
  } else if (persistence === 'persistStorage') {
    let cfb: CompoundFile | undefined;
    try {
      cfb = await CompoundFile.open(bufferSource(bytes), ctx.settings, ctx.run.deadline);
    } catch (e) {
      if (!(e instanceof StopError) || e.category !== C.Container) throw e;
      problem = `its binary part is not a compound file: ${e.message}`;
    }
    if (cfb) problem = await controlStorageProblem(cfb, cfb.root, OOXML_CONTROL_STREAMS, classId);
  } else problem = `its persistence is ${persistence}`;
  if (!problem) bin.formsData = true;
  return { classId, problem };
}

/**
 * A VBA project part, read as the compound file it is: the rules every compound file gets, which show the project to
 * the script plugins and the storages of its forms to the file plugins. The part stays, as the rewrite writes it, when
 * the project does.
 */
async function vbaPart(ctx: ZipContext, entry: ZipEntry, loc: string): Promise<ContainedDecision> {
  const run = ctx.run;
  const copy = new SpillSink(run.temp, run.threshold);
  run.held.push(copy);
  for await (const c of ctx.zip.read(entry)) await copy.write(c);
  await copy.close();
  const source = copy.source();
  run.sources.push(source);
  let cfb: CompoundFile;
  try {
    cfb = await CompoundFile.open(source, ctx.settings, run.deadline);
  } catch (e) {
    if (!(e instanceof StopError) || e.category !== C.Container) throw e;
    ctx.report(C.Macro, D.VbaProject, loc, { part: loc, reason: clip(`its compound file does not read: ${e.message}`) });
    return { keep: false };
  }
  nestedNotes(ctx, cfb, loc);
  const plan = new CompoundPlan();
  await compoundRules(ctx, cfb, loc, false, plan);
  const vba = cfb.child(cfb.root, 'VBA');
  if (!vba) ctx.report(C.Macro, D.VbaProject, loc, { part: loc, reason: 'it holds no VBA storage' });
  if (!vba || plan.drop.has(vba)) return { keep: false };
  const out = new SpillSink(run.temp, run.threshold);
  run.held.push(out);
  await writeCompound(run, cfb, plan, out);
  await out.close();
  return { keep: true, replacement: out };
}

/** A percent-encoded value, or the value as written when it does not decode. */
function safeDecode(v: string): string {
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
}

/** Parts that no chain of internal relationships reaches from the package. */
function unreferenced(ctx: ZipContext, parts: Map<string, Part>, relsOf: Map<string, ZipEntry>, roots: Set<Part>): void {
  const reached = new Set<Part>(roots);
  const queue = [...roots];
  for (let p = queue.pop(); p; p = queue.pop()) {
    for (const t of p.targets) {
      if (reached.has(t)) continue;
      reached.add(t);
      queue.push(t);
    }
  }
  for (const p of parts.values()) {
    if (!reached.has(p) && ctx.report(C.Package, D.UnreferencedPart, p.entry.name, { part: p.entry.name }).action === 'strip') ctx.plan.drop.add(nameKey(p.entry.name));
  }
  for (const [source, rels] of relsOf) {
    const owner = parts.get(source);
    if (source !== '' && !(owner && reached.has(owner)) && ctx.report(C.Package, D.UnreferencedPart, rels.name, { part: rels.name }).action === 'strip') ctx.plan.drop.add(nameKey(rels.name));
  }
}
