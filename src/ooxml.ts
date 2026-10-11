import { autoName } from './biff';
import { clip, type ZipContext } from './context';
import { StopError } from './errors';
import { FieldStack } from './fields';
import { OOXML_FORMATS } from './formats';
import { Quotes } from './formulas';
import { isAbsoluteTarget, isFragment } from './links';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { XlsbLinks } from './xlsb';
import { attr, UNKNOWN, unknownEntity, type XmlElement, type XmlHandler } from './xml';
import { nameKey, type ZipEntry } from './zip';

const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const T = 'http://schemas.openxmlformats.org/officedocument/2006/relationships/';
const S = 'http://purl.oclc.org/ooxml/officedocument/relationships/';
const MS = 'http://schemas.microsoft.com/office/2006/relationships/';
const both = (name: string) => [T + name, S + name];

/** Relationship types, lowercased, as OPC compares them without regard to case. */
const REL = {
  main: new Set([...both('officedocument'), 'http://schemas.microsoft.com/visio/2010/relationships/document']),
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
const TYPE = {
  vba: new Set(['application/vnd.ms-office.vbaproject']),
  macrosheet: new Set(['application/vnd.ms-excel.macrosheet+xml', 'application/vnd.ms-excel.intlmacrosheet+xml', 'application/vnd.ms-excel.macrosheet', 'application/vnd.ms-excel.intlmacrosheet']),
  ole: new Set(['application/vnd.openxmlformats-officedocument.oleobject']),
  control: new Set(['application/vnd.ms-office.activex+xml']),
  controlBinary: new Set(['application/vnd.ms-office.activex']),
  vml: 'application/vnd.openxmlformats-officedocument.vmldrawing',
  xlsbExternalLink: 'application/vnd.ms-excel.externallink',
};

/** Relationship types that leave the package, by the finding they make. Any other is EXTERNAL/OTHER. */
const EXTERNAL_KINDS: Array<[Set<string>, D]> = [
  [REL.template, D.Template],
  [new Set([...both('oleobject'), ...both('package')]), D.Object],
  [REL.image, D.Image],
  [REL.frame, D.Frame],
  [REL.subDocument, D.Include],
  [REL.workbook, D.Workbook],
];

const WORD_NS = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);
const SHEET_NS = new Set(['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main']);
const XDR_NS = new Set(['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetdrawing', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetdrawing']);
const R_NS = new Set(['http://schemas.openxmlformats.org/officedocument/2006/relationships', 'http://purl.oclc.org/ooxml/officedocument/relationships']);
const CUSTOM_UI_NS = new Set(['http://schemas.microsoft.com/office/2006/01/customui', 'http://schemas.microsoft.com/office/2009/07/customui']);
const VML_EXCEL_NS = 'urn:schemas-microsoft-com:office:excel';
const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const VISIO_CODE = /\b(RUNADDON|RUNADDONWARGS|CALLTHIS|RUNMACRO|OPENFILE|DOOLEVERB)\s*\(/i;

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

interface Part {
  entry: ZipEntry;
  contentType?: string;
  /** The AS bits of the relationship types that reach the part. */
  reachedAs: number;
  /** A relationship of an ActiveX control part reaches it, so it is that control's binary. */
  claimed: boolean;
  /** The parts its own relationships reach. */
  targets: Set<Part>;
  ole: boolean;
}

/** A Word attribute, qualified as Word writes it or unqualified. */
const wordAttr = (el: XmlElement, local: string) => el.attributes.find(a => a.local === local && (a.ns === '' || WORD_NS.has(a.ns.toLowerCase())))?.value;

/** The key of the part a relationships part belongs to, '' for the package, or undefined for any other part. */
const relsSource = (relsName: string): string | undefined => {
  const m = /^(.*?)_rels\/([^/]*)\.rels$/i.exec(relsName);
  return m ? nameKey(m[1] + m[2]) : undefined;
};

/**
 * Resolves an internal target against the folder of the part it belongs to, as RFC 3986 resolves a relative
 * reference: a ".." at the root of the package stays at the root. LibreOffice writes "../customXml/item1.xml" in the
 * package relationships of workbooks it saves.
 */
function resolveTarget(relsName: string, target: string): string {
  const folder = relsName.slice(0, relsName.toLowerCase().indexOf('_rels/'));
  const path = (target.startsWith('/') ? target.slice(1) : folder + target).split('#')[0];
  const out: string[] = [];
  for (const seg of path.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return nameKey(out.join('/'));
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
async function readRelationships(ctx: ZipContext, entry: ZipEntry, each: (r: Relationship) => void): Promise<void> {
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

/** Word field codes in WordprocessingML: fldSimple, and fldChar marks around instrText. */
class Fields {
  private readonly stack: FieldStack;
  private capture = false;
  constructor(ctx: ZipContext, part: string) {
    this.stack = new FieldStack(ctx, part);
  }

  start(el: XmlElement): void {
    if (el.local === 'fldSimple') this.stack.simple(wordAttr(el, 'instr') ?? '');
    else if (el.local === 'fldChar') {
      const kind = wordAttr(el, 'fldCharType');
      if (kind === 'begin') this.stack.begin();
      else if (kind === 'separate') this.stack.separate();
      else if (kind === 'end') this.stack.end();
    } else if (el.local === 'instrText' || el.local === 'delInstrText') this.capture = true;
  }

  end(el: XmlElement): void {
    if (el.local === 'instrText' || el.local === 'delInstrText') this.capture = false;
  }

  text(t: string): void {
    if (this.capture) this.stack.text(t);
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
  private formula?: { shown: Shown; quotes: Quotes; ddeBar: boolean; entity?: string };
  private macroText?: Shown;
  /** The last namespace seen and its lowercase form. Neighboring elements share one, and this spares a copy each. */
  private ns = '';
  private lower = '';

  constructor(
    private readonly ctx: ZipContext,
    private readonly part: string,
  ) {
    this.fields = new Fields(ctx, part);
  }

  private lowerNs(el: XmlElement): string {
    if (el.ns !== this.ns) {
      this.ns = el.ns;
      this.lower = el.ns.toLowerCase();
    }
    return this.lower;
  }

  start(el: XmlElement): void {
    const ns = this.lowerNs(el);
    if (CUSTOM_UI_NS.has(ns)) {
      if (el.depth === 1) this.ctx.report(C.Ui, D.Ribbon, this.part);
      for (const a of el.attributes) {
        if (a.ns === '' && a.value.trim() && /^(on[A-Z]|get[A-Z]|loadImage$)/.test(a.local)) this.ctx.report(C.Ui, D.Callback, this.part, { callback: `${a.local}=${clip(a.value)}` });
      }
      return;
    }
    if (WORD_NS.has(ns)) this.fields.start(el);
    else if (SHEET_NS.has(ns)) {
      if (el.local === 'f' || el.local === 'definedName') this.formula = { shown: new Shown(), quotes: new Quotes(), ddeBar: false };
      if (el.local === 'definedName') {
        // Excel runs the macro a name like _xlnm.Auto_Open refers to, as it does in a binary workbook.
        const name = attr(el, '', 'name') ?? '';
        const auto = autoName(name);
        if (auto) this.ctx.report(C.Macro, auto.onOpen ? D.EventOnOpen : D.Event, this.part, { event: auto.event, name: clip(name) });
      } else if (el.local === 'ddeLink') this.ctx.report(C.Dde, D.Link, this.part, { application: clip(attr(el, '', 'ddeService') ?? ''), topic: clip(attr(el, '', 'ddeTopic') ?? '') });
    } else if (XDR_NS.has(ns)) {
      const macro = attr(el, '', 'macro');
      if (macro?.trim()) this.ctx.report(C.Macro, D.Link, this.part, { macro: clip(macro) });
    } else if (ns === VML_EXCEL_NS && el.local === 'FmlaMacro') this.macroText = new Shown();
    else if (ns === VISIO_NS && el.local === 'Cell') this.visioCell(el);
    if (el.local === 'hlinkClick' || el.local === 'hlinkHover' || el.local === 'hlinkMouseOver') this.drawingAction(el);
  }

  text(t: string): void {
    this.fields.text(t);
    const f = this.formula;
    if (f) {
      f.shown.add(t);
      // Excel writes a DDE reference as application|topic!item, and "|" means nothing else outside quotes.
      for (const c of t) if (f.quotes.outside(c) && c === '|') f.ddeBar = true;
      if (f.entity === undefined && t.includes(UNKNOWN)) f.entity = unknownEntity(t);
    }
    this.macroText?.add(t);
  }

  end(el: XmlElement): void {
    const ns = this.lowerNs(el);
    if (WORD_NS.has(ns)) this.fields.end(el);
    if (SHEET_NS.has(ns) && (el.local === 'f' || el.local === 'definedName') && this.formula) {
      const { ddeBar, entity, shown } = this.formula;
      if (ddeBar) this.ctx.report(C.Dde, D.Formula, this.part, { formula: clip(shown.text) });
      else if (entity !== undefined) this.ctx.report(C.Dde, D.Formula, this.part, { formula: clip(shown.text), entity: clip(entity) });
      this.formula = undefined;
    }
    if (ns === VML_EXCEL_NS && el.local === 'FmlaMacro' && this.macroText) {
      if (this.macroText.content) this.ctx.report(C.Macro, D.Link, this.part, { macro: clip(this.macroText.text.trim()) });
      this.macroText = undefined;
    }
    if (el.depth === 1) this.fields.finish();
  }

  /** PowerPoint action settings, in a DrawingML hyperlink's action attribute. The relationship it names is reported on its own. */
  private drawingAction(el: XmlElement): void {
    const written = (attr(el, '', 'action') ?? '').trim();
    const action = written.toLowerCase();
    if (!action) return;
    const id = el.attributes.find(a => a.local === 'id' && R_NS.has(a.ns.toLowerCase()))?.value;
    const data = { trigger: el.local, action: clip(written), ...(id === undefined ? {} : { relationship: clip(id) }) };
    if (action.startsWith('ppaction://program')) this.ctx.report(C.Action, D.RunProgram, this.part, data);
    else if (action.startsWith('ppaction://macro')) this.ctx.report(C.Macro, D.Link, this.part, data);
    else if (action.startsWith('ppaction://ole')) this.ctx.report(C.Action, D.OleVerb, this.part, data);
  }

  private visioCell(el: XmlElement): void {
    const n = attr(el, '', 'N') ?? '';
    const formula = attr(el, '', 'F') ?? '';
    if (VISIO_CODE.test(formula)) {
      if (n.startsWith('Event')) this.ctx.report(C.Macro, D.Event, this.part, { event: clip(n), formula: clip(formula) });
      else this.ctx.report(C.Macro, D.Link, this.part, { cell: clip(n), formula: clip(formula) });
    }
    const v = attr(el, '', 'V');
    if (n === 'Address' && v?.trim() && !isFragment(v)) this.ctx.link(v, this.part, { source: 'Visio hyperlink' });
  }
}

/**
 * Inspects an OOXML package. The package relationships come first, since they name the main part and so the format.
 * Then each part is read in archive order, followed by its relationships, which are reported as they are read and
 * leave behind only which parts they reach and how.
 */
export async function inspectOoxml(ctx: ZipContext, contentTypes: ZipEntry): Promise<void> {
  const { defaults, overrides } = await readContentTypes(ctx, contentTypes);
  const parts = new Map<string, Part>();
  const relsOf = new Map<string, ZipEntry>();
  for (const entry of ctx.zip.entries) {
    if (entry === contentTypes) continue;
    // OPC has no folders. A directory entry is still inflated once, so its bytes are checked like any other.
    if (entry.directory) {
      await ctx.drain(entry);
      continue;
    }
    const source = relsSource(entry.name);
    if (source !== undefined) {
      relsOf.set(source, entry);
      continue;
    }
    const key = nameKey(entry.name);
    const ext = /\.([^./]*)$/.exec(entry.name)?.[1]?.toLowerCase();
    parts.set(key, { entry, contentType: overrides.get(key) ?? (ext === undefined ? undefined : defaults.get(ext)), reachedAs: 0, claimed: false, targets: new Set(), ole: false });
  }

  const roots = new Set<Part>();
  const mains: Array<Part | undefined> = [];
  const relate = (rels: ZipEntry, from: Part | undefined, r: Relationship) => {
    const location = `${rels.name}, ${clip(r.id)}`;
    const part = r.external ? undefined : parts.get(resolveTarget(rels.name, r.target));
    if (part) {
      part.reachedAs |= REACHED_AS.get(r.type) ?? 0;
      if (TYPE.control.has(from?.contentType ?? '')) part.claimed = true;
      (from?.targets ?? roots).add(part);
    }
    // A hyperlink to a part of the package, or to a place in the document, goes nowhere else.
    if (REL.hyperlink.has(r.type)) {
      if (!part && !isFragment(r.target)) ctx.link(r.target, location);
    } else if (r.external) {
      const kind = EXTERNAL_KINDS.find(([types]) => types.has(r.type))?.[1] ?? D.Other;
      ctx.external(kind, r.target, location, { type: clip(r.type) }, clip(r.type.slice(r.type.lastIndexOf('/') + 1)));
    }
  };

  const packageRels = relsOf.get('');
  if (packageRels) {
    await readRelationships(ctx, packageRels, r => {
      if (REL.main.has(r.type)) mains.push(r.external ? undefined : parts.get(resolveTarget(packageRels.name, r.target)));
      relate(packageRels, undefined, r);
    });
  }
  if (mains.length > 1) throw new StopError(C.Package, D.AmbiguousFormat, 'The package names more than one main part', { reason: 'more than one main part' });
  const main = mains[0];
  if (!main) throw new StopError(C.Package, D.Unrecognized, 'The package has no main part', { reason: 'no main document part' });
  const format = OOXML_FORMATS.get(main.contentType ?? '');
  if (!format) throw new StopError(C.Package, D.Unrecognized, 'The main part has an unknown content type', { contentType: clip(main.contentType ?? '') });
  ctx.format = format;

  for (const [key, p] of parts) {
    const ct = p.contentType;
    if (ct === undefined) {
      ctx.report(C.Package, D.UntypedPart, p.entry.name, { part: p.entry.name });
      p.ole = await ctx.drain(p.entry);
    } else if (ct.endsWith('+xml') || ct === 'application/xml' || ct === 'text/xml' || ct === TYPE.vml) {
      await ctx.parse(p.entry, new PartRules(ctx, p.entry.name), ct === TYPE.vml);
    } else if (ct === TYPE.xlsbExternalLink) {
      // The binary form of an Excel ddeLink.
      const links = new XlsbLinks();
      p.ole = await ctx.drain(p.entry, c => links.write(c));
      for (const l of links.dde) ctx.report(C.Dde, D.Link, p.entry.name, { application: clip(l.application), topic: clip(l.topic) });
    } else p.ole = await ctx.drain(p.entry);
    const rels = relsOf.get(key);
    if (rels) await readRelationships(ctx, rels, r => relate(rels, p, r));
  }
  // Relationships of parts that are not there still say what the parts they reach are.
  for (const [source, rels] of relsOf) if (source !== '' && !parts.has(source)) await readRelationships(ctx, rels, r => relate(rels, undefined, r));

  for (const p of parts.values()) partRules(ctx, p);
  unreferenced(ctx, parts, relsOf, roots);
}

/** Findings about a whole part, from its content type and the relationships that reach it. */
function partRules(ctx: ZipContext, p: Part): void {
  const ct = p.contentType ?? '';
  const loc = p.entry.name;
  const data = { part: loc, contentType: clip(ct) };
  const as = (bit: number) => (p.reachedAs & bit) !== 0;
  const vba = TYPE.vba.has(ct) || as(AS.vba);
  const ole = TYPE.ole.has(ct) || as(AS.ole);
  const control = TYPE.control.has(ct) || as(AS.control);
  const controlBinary = TYPE.controlBinary.has(ct) || as(AS.controlBinary);
  const pkg = as(AS.package);
  if (vba) ctx.report(C.Macro, D.VbaProject, loc, data);
  if (TYPE.macrosheet.has(ct) || as(AS.macrosheet)) ctx.report(C.Macro, D.XlmMacrosheet, loc, data);
  if (ole) ctx.report(C.Embedded, D.OleObject, loc, data);
  if (pkg) ctx.report(C.Embedded, D.Package, loc, data);
  if (control || (controlBinary && !p.claimed)) ctx.report(C.Control, D.ActiveX, loc, data);
  if (p.ole && !vba && !ole && !control && !controlBinary && !pkg) ctx.report(C.Embedded, D.OleFile, loc, data);
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
  for (const p of parts.values()) if (!reached.has(p)) ctx.report(C.Package, D.UnreferencedPart, p.entry.name, { part: p.entry.name });
  for (const [source, rels] of relsOf) {
    const owner = parts.get(source);
    if (source !== '' && !(owner && reached.has(owner))) ctx.report(C.Package, D.UnreferencedPart, rels.name, { part: rels.name });
  }
}
