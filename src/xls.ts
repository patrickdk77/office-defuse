import { autoName, type BiffRecord, objSubrecords, parseBoundSheet, parseLbl, R, RecordReader, type Sheet, xlString } from './biff';
import type { CfbEntry, CompoundFile } from './cfb';
import { clip, type Reported, type RunContext } from './context';
import { StopError } from './errors';
import { type BiffNameRef, biffCalls } from './formulas';
import { SpillSink } from './io';
import { isFragment } from './links';
import { ByteReader, bounded, readMoniker, utf16 } from './oleobj';
import type { CompoundPlan } from './plan';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { hasEdits, noEdits, RewriteError, rewriteWorkbook, type WorkbookEdits } from './xls-writer';

const STATES = ['visible', 'hidden', 'very hidden', 'hidden'];

/** Turns an encrypted record body into its plain bytes. Without encryption the body is as stored. */
export type Decrypt = (r: BiffRecord) => Uint8Array;

/**
 * Opens a workbook encrypted by its FILEPASS record, or says it cannot. Set by the decryption module, so this module
 * reads workbooks without knowing how they are encrypted.
 */
export type Unlock = (ctx: RunContext, filePass: BiffRecord, stream: string) => Decrypt | Promise<Decrypt>;

interface SupBook {
  ctab: number;
  cch: number;
  path: string;
  /** Whether ExternName records after it name OLE objects, and whether they name DDE items. */
  ole: boolean;
  dde: boolean;
  /** The names its ExternName records give, in order, which a formula refers to by index. */
  names: string[];
  /** Where the record starts in the stream. */
  offset: number;
}

/** Record types of cell formulas (MS-XLS 2.4.127, 2.4.260 and 2.4.4): FORMULA, SHRFMLA and ARRAY. */
const FORMULA = 0x0006;
const SHRFMLA = 0x04bc;
const ARRAY = 0x0221;
/** Where the parsed formula starts in each: its cce, then its rgce. */
const FORMULA_AT = new Map([
  [FORMULA, 20],
  [SHRFMLA, 8],
  [ARRAY, 12],
]);

/**
 * Whether a workbook stream holds a template or an add-in: its globals substream, from the BOF that starts the stream
 * to its EOF, holds a Template or an AddIn record. Record headers are never encrypted, so an encrypted workbook tells
 * this without its password. A stream that does not start with a BOF record is neither.
 */
export async function workbookKind(cfb: CompoundFile, stream: CfbEntry, deadline?: number): Promise<{ template: boolean; addIn: boolean }> {
  const kind = { template: false, addIn: false };
  const records = new RecordReader(cfb, stream, 0, deadline);
  const first = await records.next();
  if (first?.type !== R.bof || first.body.length < 4) return kind;
  for (let r = await records.next(); r && r.type !== R.eof; r = await records.next()) {
    if (r.type === R.template) kind.template = true;
    else if (r.type === R.addIn) kind.addIn = true;
  }
  return kind;
}

/**
 * Inspects an Excel 5.0 to 2003 workbook stream. Excel reads the globals substream from the start of the stream,
 * then each sheet from where its BoundSheet record says the sheet starts. The globals give the macro sheets, the
 * names Excel runs as macros, and the DDE, OLE and workbook links. The sheets give shapes that run macros and
 * hyperlinks. An encrypted workbook goes through `unlock`.
 */
/** An ActiveX control a sheet's Obj record keeps in the Ctls stream: its ProgID and where its data is. */
export interface CtlsControl {
  progId: string;
  position: number;
  size: number;
  sheet: string;
}

export async function excelRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry, unlock: Unlock, plan?: CompoundPlan, controls: CtlsControl[] = []): Promise<void> {
  const location = stream.name;
  const records = new RecordReader(cfb, stream, 0, ctx.settings.deadline);
  const first = await records.next();
  if (first?.type !== R.bof || first.body.length < 4)
    throw new StopError(C.Package, D.Unrecognized, `The ${stream.name} stream does not start with a BOF record`, { reason: 'a workbook stream without a BOF record' });
  const biff8 = (first.body[0] | (first.body[1] << 8)) === 0x0600;
  let decrypt: Decrypt = r => r.body;
  const sheets: Sheet[] = [];
  const books: SupBook[] = [];
  const names: string[] = [];
  /** The SupBook each XTI of the ExternSheet record names. */
  const xti: number[] = [];
  const edits = noEdits();
  /** Names that run a macro and go, by their Lbl record, given new names once every name is known. */
  const renames: Array<{ offset: number; event: string }> = [];
  for (let r = await records.next(); r && r.type !== R.eof; r = await records.next()) {
    if (r.type === R.filePass) {
      const before = ctx.findings.length;
      decrypt = await unlock(ctx, r, location);
      // The rewrite writes the workbook decrypted, so one VelvetSweatshop opens is rewritten when that finding strips.
      if (ctx.findings.slice(before).some(f => f.category === C.Encrypted && f.detail === D.DefaultPassword && f.action === 'strip')) edits.decrypt = true;
      continue;
    }
    const body = decrypt(r);
    if (r.type === R.boundSheet) sheets.push({ ...parseBoundSheet(body, biff8), record: r.offset });
    else if (r.type === R.lbl) {
      const name = parseLbl(body, biff8);
      names.push(name.name);
      const auto = autoName(name.name);
      // Removing the name would renumber the names after it, which formulas refer to by number, so it is renamed.
      if (auto && ctx.report(C.Macro, auto.onOpen ? D.EventOnOpen : D.Event, location, { event: auto.event, name: clip(name.name) }).action === 'strip')
        renames.push({ offset: r.offset, event: auto.event });
    } else if (r.type === R.supBook && body.length >= 4) {
      const cch = body[2] | (body[3] << 8);
      books.push({ ctab: body[0] | (body[1] << 8), cch, path: cch > 0 && cch <= 0xff ? xlString(body, 4, cch).text : '', ole: false, dde: false, names: [], offset: r.offset });
    } else if (r.type === R.externSheet && biff8 && body.length >= 2) {
      const n = body[0] | (body[1] << 8);
      for (let k = 0; k < n && 2 + k * 6 + 2 <= body.length; k++) xti.push(body[2 + k * 6] | (body[3 + k * 6] << 8));
    } else if (r.type === R.externName && body.length >= 2 && books.length) {
      if (biff8 && body.length >= 7) books[books.length - 1].names.push(xlString(body, 7, body[6]).text);
      // fOle and fOleLink mark an OLE object's name. Excel writes StdDocumentName with fOle beside the items of a DDE
      // link, so a link with any DDE item is a DDE link.
      if ((body[0] | (body[1] << 8)) & 0x18) books[books.length - 1].ole = true;
      else books[books.length - 1].dde = true;
    }
  }
  // A link that goes keeps its SupBook, which formulas count to by index, with a path to the workbook itself.
  for (const b of books) if (linkRules(ctx, b, location)?.action === 'strip') edits.selfBooks.add(b.offset);
  const nameOf = (ref: BiffNameRef) => (ref.kind === 'name' ? names[ref.index - 1] : books[xti[ref.ixti] ?? -1]?.names[ref.index - 1]);
  // A macro of the workbook's own: an external name of the SupBook that is the workbook itself, as Excel names a VBA
  // macro a button runs. A defined name may stand for anything, so it is not one.
  const ownName = (ref: BiffNameRef) => (ref.kind === 'namex' && books[xti[ref.ixti] ?? -1]?.cch === 0x401 ? nameOf(ref) : undefined);
  for (const s of sheets) await sheetRules(ctx, cfb, stream, s, decrypt, location, edits, controls, biff8 ? nameOf : undefined, biff8 ? ownName : undefined);
  const taken = new Set(names.map(n => n.toLowerCase()));
  for (const { offset, event } of renames) {
    let k = 1;
    while (taken.has(`removed_${event}_${k}`.toLowerCase())) k++;
    const name = `Removed_${event}_${k}`;
    taken.add(name.toLowerCase());
    edits.names.set(offset, name);
  }
  // BIFF5 has no SupBook, HLink or Obj macro records in the shape BIFF8 has, so only its sheets and names change. What
  // else strips stays, and the check of the output says so.
  if (!biff8) {
    edits.objMacros.clear();
    edits.selfBooks.clear();
    edits.hyperlinks.clear();
  }
  if (!plan || !hasEdits(edits)) return;
  const sink = new SpillSink(ctx.run.temp, ctx.run.threshold);
  ctx.run.held.push(sink);
  try {
    await rewriteWorkbook(cfb, stream, edits, biff8, decrypt, sink, ctx.settings.deadline);
    await sink.close();
    plan.replace.set(stream, sink);
  } catch (e) {
    // Left as it is, the stream still holds what the findings name, and the check of the output says so.
    if (!(e instanceof RewriteError)) throw e;
  }
}

/**
 * A SupBook record, read as MS-XLS 2.4.271 and oletools tell its kinds apart. A path that starts with U+0002 or U+0003
 * refers to the workbook itself, as LibreOffice reads it.
 */
function linkRules(ctx: RunContext, b: SupBook, location: string): Reported | undefined {
  if (b.cch === 0x401 || b.cch === 0x3a01 || b.path === ' ' || b.path === '\0' || !b.path || b.path.charCodeAt(0) === 2 || b.path.charCodeAt(0) === 3) return undefined;
  if (b.ctab === 0) {
    // A DDE or OLE data source: the application or class, U+0003, and the topic or file.
    const [application, topic = ''] = b.path.split('\x03');
    if (b.ole && !b.dde) {
      const data = { source: 'SupBook', className: clip(application) };
      return topic ? ctx.external(D.Object, topic, location, data, 'linked object') : ctx.report(C.External, D.Object, location, data);
    }
    return ctx.report(C.Dde, D.Link, location, { application: clip(application), topic: clip(topic) });
  }
  return ctx.external(D.Workbook, virtualPath(b.path), location, { source: 'SupBook' }, 'external workbook');
}

/**
 * A VirtualPath (MS-XLS 2.5.277) written out. One that starts with U+0001 is encoded: its control characters name a
 * volume, the root, a folder, a step up, a URL's volume or Excel's startup folders.
 */
export function virtualPath(p: string): string {
  if (p.charCodeAt(0) !== 1) return p;
  let out = '';
  for (let i = 1; i < p.length; i++) {
    const c = p.charCodeAt(i);
    if (c === 1) {
      // A drive letter, or @ for a UNC host. A U+0003 after it starts the path.
      const v = p[i + 1] ?? '';
      i++;
      out += v === '@' ? '\\\\' : `${v}:`;
    } else if (c === 2 || c === 3) out += '\\';
    else if (c === 4) out += '..\\';
    // A long volume, such as a URL, after a character that counts it.
    else if (c === 5) i++;
    else if (c >= 6 && c <= 8) out += 'XLSTART\\';
    else out += p[i];
  }
  return out;
}

/** One sheet's substream, with any chart substreams inside it: Obj records that run macros, and hyperlinks. */
async function sheetRules(
  ctx: RunContext,
  cfb: CompoundFile,
  stream: CfbEntry,
  s: Sheet,
  decrypt: Decrypt,
  location: string,
  edits: WorkbookEdits,
  controls: CtlsControl[],
  nameOf?: (ref: BiffNameRef) => string | undefined,
  ownName?: (ref: BiffNameRef) => string | undefined,
): Promise<void> {
  const records = new RecordReader(cfb, stream, s.position, ctx.settings.deadline);
  const bof = await records.next();
  if (bof?.type !== R.bof || bof.body.length < 4) return;
  const bofType = bof.body[2] | (bof.body[3] << 8);
  // The sheet's own BOF says what it is, and so does its BoundSheet record. Either one makes a macro sheet.
  const macroSheet = s.type === 1 || bofType === 0x0040;
  // A macro sheet becomes an empty worksheet of the same name, so the sheets after it keep their indexes.
  if (macroSheet && ctx.report(C.Macro, D.XlmMacrosheet, location, { sheet: clip(s.name), state: STATES[s.state] }).action === 'strip' && s.record !== undefined) edits.emptySheets.add(s.record);
  let depth = 1;
  for (let r = await records.next(); r && depth > 0; r = await records.next()) {
    if (r.type === R.bof) depth++;
    else if (r.type === R.eof) depth--;
    else if (r.type === R.obj) {
      const body = decrypt(r);
      const control = ctlsControl(body);
      if (control) controls.push({ ...control, sheet: s.name });
      const macro = objMacro(body);
      if (!macro) continue;
      // A shape that runs one of the workbook's own macros stays when its VBA project does.
      const name = macro.ref && ownName ? ownName(macro.ref) : undefined;
      const data = { sheet: clip(s.name), ...(macro.ref && nameOf ? { macro: clip(nameOf(macro.ref) ?? '') } : {}) };
      if (ctx.macroTrigger(C.Macro, D.Link, location, data, name).action === 'strip') edits.objMacros.add(r.offset);
    } else if (r.type === R.hlink) {
      const target = hyperlink(decrypt(r));
      if (target.trim() && !isFragment(target) && ctx.link(target, location, { sheet: clip(s.name), source: 'HLink record' }).action === 'strip') edits.hyperlinks.add(r.offset);
    } else if (nameOf && !macroSheet && FORMULA_AT.has(r.type)) {
      // A macro sheet's formulas are its macros, which its own finding covers.
      const b = decrypt(r);
      const at = FORMULA_AT.get(r.type) ?? 0;
      if (b.length < at + 2) continue;
      const cce = b[at] | (b[at + 1] << 8);
      for (const [fn, kind] of biffCalls(b.subarray(at + 2, at + 2 + cce), nameOf))
        ctx.report(kind === 'fetch' ? C.External : C.Macro, D.Formula, location, { sheet: clip(s.name), function: clip(fn) });
    }
  }
}

/**
 * The FtMacro subrecord of an Obj record (MS-XLS 2.4.181), the macro a click runs, with the name its formula refers
 * to when the formula is one PtgName or PtgNameX and nothing else. Undefined when the record has none.
 */
function objMacro(b: Uint8Array): { ref?: BiffNameRef } | undefined {
  for (const { ft, data: f } of objSubrecords(b)) {
    if (ft !== 0x04) continue;
    if (f.length === 0) return undefined;
    // As Excel writes it: cce in 15 bits, 4 unused bytes, and the tokens. PtgName and PtgNameX are classed tokens,
    // 0x23, 0x43 or 0x63 and 0x39, 0x59 or 0x79; a base token below 0x20 shares their low bits and is neither.
    const cce = f.length >= 2 ? (f[0] | (f[1] << 8)) & 0x7fff : 0;
    const rgce = f.subarray(6, 6 + cce);
    const ptg = rgce[0] ?? 0;
    const u32 = (k: number) => (rgce[k] | (rgce[k + 1] << 8) | (rgce[k + 2] << 16) | (rgce[k + 3] << 24)) >>> 0;
    if (ptg >= 0x20 && (ptg & 0x1f) === 0x03 && cce === 5 && rgce.length === 5) return { ref: { kind: 'name', index: u32(1) } };
    if (ptg >= 0x20 && (ptg & 0x1f) === 0x19 && cce === 7 && rgce.length === 7) return { ref: { kind: 'namex', ixti: rgce[1] | (rgce[2] << 8), index: u32(3) } };
    return {};
  }
  return undefined;
}

/**
 * The control an Obj record keeps in the Ctls stream: FtPioGrbit says an ActiveX control (fCtl) whose data is in Ctls
 * (fPrstm), and FtPictFmla gives its ProgID, in the formula's embedInfo, and its place in Ctls after the formula.
 */
function ctlsControl(b: Uint8Array): { progId: string; position: number; size: number } | undefined {
  let flags = 0;
  for (const { ft, data: d } of objSubrecords(b)) {
    if (ft === 0x08 && d.length >= 2) flags = d[0] | (d[1] << 8);
    if (ft !== 0x09 || (flags & 0x30) !== 0x30 || d.length < 2) continue;
    // ObjFmla: cbFmla, then cce, 4 unused bytes, the tokens, and PictFmlaEmbedInfo: 0x03, cbClass, a reserved byte
    // and the class name. lPosInCtlStm and cbBufInCtlStm follow.
    const cbFmla = d[0] | (d[1] << 8);
    const f = d.subarray(2, 2 + cbFmla);
    const cce = f.length >= 2 ? (f[0] | (f[1] << 8)) & 0x7fff : 0;
    const info = f.subarray(6 + cce);
    const progId = info[0] === 0x03 && info.length >= 4 ? xlString(info, 3, info[1]).text : '';
    const after = 2 + cbFmla;
    if (after + 8 > d.length) return undefined;
    const u32 = (k: number) => (d[k] | (d[k + 1] << 8) | (d[k + 2] << 16) | (d[k + 3] << 24)) >>> 0;
    return { progId, position: u32(after), size: u32(after + 4) };
  }
  return undefined;
}

/** The target of an HLink record: its Hyperlink Object (MS-OSHARED 2.3.7.1) after the cell range and class id. */
export function hyperlink(b: Uint8Array): string {
  const r = new ByteReader(b);
  return bounded(
    () => {
      r.skip(8 + 16 + 4);
      const flags = r.u32();
      const string = () => utf16(r.bytes(r.u32() * 2)).replace(/\0+$/, '');
      if (flags & 0x10) string();
      if (flags & 0x80) string();
      if (!(flags & 0x01)) return '';
      return flags & 0x100 ? string() : readMoniker(r).target;
    },
    () => '',
  );
}
