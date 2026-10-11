import { autoName, type BiffRecord, parseBoundSheet, parseLbl, R, RecordReader, type Sheet, xlString } from './biff';
import type { CfbEntry, CompoundFile } from './cfb';
import { clip, type RunContext } from './context';
import { StopError } from './errors';
import { isFragment } from './links';
import { ByteReader, bounded, readMoniker, utf16 } from './oleobj';
import { OfficeCategory as C, OfficeDetail as D } from './types';

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
}

/**
 * Inspects an Excel 5.0 to 2003 workbook stream. Excel reads the globals substream from the start of the stream,
 * then each sheet from where its BoundSheet record says the sheet starts. The globals give the macro sheets, the
 * names Excel runs as macros, and the DDE, OLE and workbook links. The sheets give shapes that run macros and
 * hyperlinks. An encrypted workbook goes through `unlock`.
 */
export async function excelRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry, unlock: Unlock): Promise<{ template: boolean; addIn: boolean }> {
  const location = stream.name;
  const records = new RecordReader(cfb, stream, 0, ctx.settings.deadline);
  const first = await records.next();
  if (first?.type !== R.bof || first.body.length < 4)
    throw new StopError(C.Package, D.Unrecognized, `The ${stream.name} stream does not start with a BOF record`, { reason: 'a workbook stream without a BOF record' });
  const biff8 = (first.body[0] | (first.body[1] << 8)) === 0x0600;
  let decrypt: Decrypt = r => r.body;
  const sheets: Sheet[] = [];
  const books: SupBook[] = [];
  const flags = { template: false, addIn: false };
  for (let r = await records.next(); r && r.type !== R.eof; r = await records.next()) {
    if (r.type === R.filePass) {
      decrypt = await unlock(ctx, r, location);
      continue;
    }
    const body = decrypt(r);
    if (r.type === R.boundSheet) sheets.push(parseBoundSheet(body, biff8));
    else if (r.type === R.template) flags.template = true;
    else if (r.type === R.addIn) flags.addIn = true;
    else if (r.type === R.lbl) {
      const name = parseLbl(body, biff8);
      const auto = autoName(name.name);
      if (auto) ctx.report(C.Macro, auto.onOpen ? D.EventOnOpen : D.Event, location, { event: auto.event, name: clip(name.name) });
    } else if (r.type === R.supBook && body.length >= 4) {
      const cch = body[2] | (body[3] << 8);
      books.push({ ctab: body[0] | (body[1] << 8), cch, path: cch > 0 && cch <= 0xff ? xlString(body, 4, cch).text : '', ole: false, dde: false });
    } else if (r.type === R.externName && body.length >= 2 && books.length) {
      // fOle and fOleLink mark an OLE object's name. Excel writes StdDocumentName with fOle beside the items of a DDE
      // link, so a link with any DDE item is a DDE link.
      if ((body[0] | (body[1] << 8)) & 0x18) books[books.length - 1].ole = true;
      else books[books.length - 1].dde = true;
    }
  }
  for (const b of books) linkRules(ctx, b, location);
  for (const s of sheets) await sheetRules(ctx, cfb, stream, s, decrypt, location);
  return flags;
}

/** A SupBook record, read as MS-XLS 2.4.271 and oletools tell its kinds apart. */
function linkRules(ctx: RunContext, b: SupBook, location: string): void {
  if (b.cch === 0x401 || b.cch === 0x3a01 || b.path === ' ' || b.path === '\0' || !b.path) return;
  if (b.ctab === 0) {
    // A DDE or OLE data source: the application or class, U+0003, and the topic or file.
    const [application, topic = ''] = b.path.split('\x03');
    if (b.ole && !b.dde) {
      const data = { source: 'SupBook', className: clip(application) };
      if (topic) ctx.external(D.Object, topic, location, data, 'linked object');
      else ctx.report(C.External, D.Object, location, data);
    } else ctx.report(C.Dde, D.Link, location, { application: clip(application), topic: clip(topic) });
    return;
  }
  ctx.external(D.Workbook, virtualPath(b.path), location, { source: 'SupBook' }, 'external workbook');
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
async function sheetRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry, s: Sheet, decrypt: Decrypt, location: string): Promise<void> {
  const records = new RecordReader(cfb, stream, s.position, ctx.settings.deadline);
  const bof = await records.next();
  if (bof?.type !== R.bof || bof.body.length < 4) return;
  const bofType = bof.body[2] | (bof.body[3] << 8);
  // The sheet's own BOF says what it is, and so does its BoundSheet record. Either one makes a macro sheet.
  if (s.type === 1 || bofType === 0x0040) ctx.report(C.Macro, D.XlmMacrosheet, location, { sheet: clip(s.name), state: STATES[s.state] });
  let depth = 1;
  for (let r = await records.next(); r && depth > 0; r = await records.next()) {
    if (r.type === R.bof) depth++;
    else if (r.type === R.eof) depth--;
    else if (r.type === R.obj) {
      if (objMacro(decrypt(r))) ctx.report(C.Macro, D.Link, location, { sheet: clip(s.name) });
    } else if (r.type === R.hlink) {
      const target = hyperlink(decrypt(r));
      if (target.trim() && !isFragment(target)) ctx.link(target, location, { sheet: clip(s.name), source: 'HLink record' });
    }
  }
}

/** Whether an Obj record (MS-XLS 2.4.181) holds an FtMacro subrecord, the macro a click runs. */
function objMacro(b: Uint8Array): boolean {
  let at = 0;
  while (at + 4 <= b.length) {
    const ft = b[at] | (b[at + 1] << 8);
    const cb = b[at + 2] | (b[at + 3] << 8);
    // FtEnd ends the list, and FtLbsData's length does not count all its data, so the walk stops there.
    if (ft === 0x00 || ft === 0x13) return false;
    if (ft === 0x04) return cb > 0;
    at += 4 + cb;
  }
  return false;
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
