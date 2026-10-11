import { CFB_SIGNATURE, CompoundFile } from './cfb';
import { clip, type RunContext } from './context';
import { checkTime, StopError } from './errors';
import { FieldStack } from './fields';
import { bufferSource } from './io';
import { isFragment } from './links';
import { compoundRules, describeObject, isEquationClass, linkedObject, nestedNotes, packageData } from './objects';
import { latin1, parseOle1, parsePackage } from './oleobj';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';

/**
 * The destinations of the RTF 1.9.1 specification, which rtfobj lists too. Word sends the text of each to its own
 * place, so none of it is object data or a field code. \rtf is not one, though it opens the document.
 */
const DESTINATIONS = new Set(
  (
    'aftncn aftnsep aftnsepc annotation atnauthor atndate atnid atnparent atnref atrfend atrfstart author background bkmkend bkmkstart ' +
    'blipuid buptim category colorschememapping colortbl comment company creatim datafield datastore defchp defpap do doccomm docvar ' +
    'dptxbxtext ebcend ebcstart factoidname falt fchars ffdeftext ffentrymcr ffexitmcr ffformat ffhelptext ffl ffname ffstattext field ' +
    'file filetbl fldinst fldrslt fldtype fontemb fonttbl footer footerf footerl footerr footnote formfield ftncn ftnsep ftnsepc g ' +
    'generator gridtbl header headerf headerl headerr hl hlfr hlinkbase hlloc hlsrc hsv info keywords latentstyles lchars levelnumbers ' +
    'leveltext lfolevel linkval list listlevel listname listoverride listoverridetable listpicture liststylename listtable listtext ' +
    'lsdlockedexcept macc maccPr mailmerge malnScr manager margPr mbar mbarPr mbaseJc mbegChr mborderBox mborderBoxPr mbox mboxPr mchr ' +
    'mcount mctrlPr md mdeg mdegHide mden mdiff mdPr me mendChr meqArr meqArrPr mf mfName mfPr mfunc mfuncPr mgroupChr mgroupChrPr mgrow ' +
    'mhideBot mhideLeft mhideRight mhideTop mlim mlimLoc mlimLow mlimLowPr mlimUpp mlimUppPr mm mmaddfieldname mmathPict mmaxDist mmc ' +
    'mmcJc mmconnectstr mmconnectstrdata mmcPr mmcs mmdatasource mmheadersource mmmailsubject mmodso mmodsofilter mmodsofldmpdata ' +
    'mmodsomappedname mmodsoname mmodsorecipdata mmodsosort mmodsosrc mmodsotable mmodsoudl mmodsoudldata mmodsouniquetag mmPr mmquery ' +
    'mmr mnary mnaryPr mnoBreak mnum mobjDist moMath moMathPara moMathParaPr mopEmu mphant mphantPr mplcHide mpos mr mrad mradPr mrPr ' +
    'msepChr mshow mshp msPre msPrePr msSub msSubPr msSubSup msSubSupPr msSup msSupPr mstrikeBLTR mstrikeH mstrikeTLBR mstrikeV msub ' +
    'msubHide msup msupHide mtransp mtype mvertJc mvfmf mvfml mvtof mvtol mzeroAsc mzeroDesc mzeroWid nesttableprops nonesttables ' +
    'objalias objclass objdata object objname objsect oldcprops oldpprops oldsprops oldtprops oleclsid operator panose password ' +
    'passwordhash pgp pgptbl picprop pict pn pnseclvl pntext pntxta pntxtb printim propname protend protstart protusertbl result revtbl ' +
    'revtim rxe shp shpgrp shpinst shppict shprslt shptxt sn sp staticval stylesheet subject sv svb tc template themedata title txe ud ' +
    'upr userprops wgrffmtfilter windowcaption writereservation writereservhash xe xform xmlattrname xmlattrvalue xmlclose xmlname ' +
    'xmlnstbl xmlopen margSz pnaiu pnaiud'
  ).split(' '),
);
/** Destinations whose content holds no field, object or link, so their groups are skipped whole: tables, styles, pictures, metadata. */
const INERT = new Set(
  (
    'fonttbl colortbl stylesheet info pict themedata colorschememapping datastore latentstyles listtable listoverridetable revtbl ' +
    'generator xmlnstbl filetbl pgptbl protusertbl blipuid datafield fontemb panose falt password passwordhash wgrffmtfilter'
  ).split(' '),
);
/** Object kinds, from the control word that sets them. */
const OBJECT_KINDS = new Set(['objemb', 'objlink', 'objautlink', 'objsub', 'objpub', 'objicemb', 'objhtml', 'objocx']);
/** Word reads at most this many letters of a control word's name and digits of its parameter. */
const MAX_WORD = 250;
const CHUNK = 262144;
/** Collected text other than an object's data is short: a class name, a template path, a link. */
const MAX_TEXT = 4096;

/** A destination the reader is in, which ends when the group it opened in closes, and what it collects. */
type FrameBody =
  | { kind: 'document' | 'skip' | 'field' | 'fldinst' | 'fldrslt' | 'hl' }
  | { kind: 'object'; object: RtfObject }
  | { kind: 'objdata'; object: RtfObject; hex: HexData; alt: HexData; standalone: boolean }
  /** A group inside object data that \* marks and starts with a control word that is not a destination. */
  | { kind: 'objskip'; data: { hex: HexData } }
  | { kind: 'objclass' | 'template' | 'hlsrc'; text: string; object?: RtfObject };
/** A destination with the group depth it opened at and the \uc count of its group. */
type Frame = FrameBody & { depth: number; uc: number };

interface RtfObject {
  offset: number;
  kinds: Set<string>;
  update: boolean;
  className: string;
  /** The decoded data, once its objdata group closes. */
  data?: Uint8Array;
  /** The data read with the groups that \* marks left out, when that differs. */
  alt?: Uint8Array;
}

/** An object's data as Word decodes its hex digits, with the quirks rtfobj documents. */
class HexData {
  private bytes: Uint8Array;
  private length = 0;
  /** A hex digit read and not yet paired, or -1. */
  private nibble = -1;
  readonly offset: number;
  constructor(
    offset: number,
    private readonly limit: number,
  ) {
    this.offset = offset;
    this.bytes = new Uint8Array(1024);
  }

  digit(v: number): void {
    if (this.nibble < 0) {
      this.nibble = v;
      return;
    }
    this.push((this.nibble << 4) | v);
    this.nibble = -1;
  }

  /** Word drops a hex digit left unpaired before a \'hh, and so does rtfobj. */
  quote(): void {
    this.nibble = -1;
  }

  /** \bin data in hex goes in as its bytes, two digits each, so an unpaired digit stays unpaired. */
  binary(b: Uint8Array): void {
    for (const x of b) {
      this.digit(x >> 4);
      this.digit(x & 15);
    }
  }

  private push(b: number): void {
    if (this.length >= this.limit) throw new StopError(C.Limit, D.DecompressedSize, `The data of the object at offset ${this.offset} decodes past ${this.limit} bytes`, { limit: this.limit });
    if (this.length === this.bytes.length) {
      const grown = new Uint8Array(Math.min(this.limit, this.bytes.length * 2));
      grown.set(this.bytes);
      this.bytes = grown;
    }
    this.bytes[this.length++] = b;
  }

  /** The decoded bytes. A last digit left unpaired is dropped, as Word drops it. */
  result(): Uint8Array {
    return this.bytes.subarray(0, this.length);
  }
}

const hexValue = (c: number) => (c >= 0x30 && c <= 0x39 ? c - 0x30 : c >= 0x41 && c <= 0x46 ? c - 0x37 : c >= 0x61 && c <= 0x66 ? c - 0x57 : -1);
const isLetter = (c: number) => (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
const isDigit = (c: number) => c >= 0x30 && c <= 0x39;

/**
 * Inspects an RTF file the way Word reads one. Objects with their data, fields, the attached template and shape
 * hyperlinks are reported. Unknown destinations marked with \* are skipped, unknown control words are ignored, and
 * text after the document's last brace is still read, as rtfobj reads it, and reported as trailing data.
 */
export async function inspectRtf(ctx: RunContext, source: ByteSource, size: number): Promise<void> {
  ctx.format = { format: 'rtf', mimeType: 'application/rtf' };
  const reader = new RtfReader(ctx, size);
  for (let pos = 0; pos < size; pos += CHUNK) {
    checkTime(ctx.settings.deadline);
    const chunk = await source.read(pos, Math.min(CHUNK, size - pos));
    await reader.write(chunk, pos + chunk.length >= size);
  }
  await reader.end();
}

class RtfReader {
  private readonly frames: Frame[];
  private readonly fields: FieldStack;
  private depth = 0;
  /** A partial token carried from one chunk to the next. */
  private carry: Uint8Array = new Uint8Array(0);
  /** The file offset of the first byte of the buffer being read. */
  private base = 0;
  /** \bin bytes still to take, and where they go. */
  private binLeft = 0;
  /** Fallback characters still to skip after a \u. */
  private skipChars = 0;
  /** Set by \*, which makes an unknown destination that follows skipped. */
  private starred = false;
  /** Where the document's outermost group closed, and how much that is not white space comes after it. */
  private closedAt = -1;
  private trailing = 0;
  private decoded = 0;
  /** Objects whose groups closed, read in order once the chunk is done. */
  private readonly pending: RtfObject[] = [];

  constructor(
    private readonly ctx: RunContext,
    private readonly size: number,
  ) {
    this.frames = [{ kind: 'document', depth: 0, uc: 1 }];
    this.fields = new FieldStack(ctx, 'document');
  }

  private get top(): Frame {
    return this.frames[this.frames.length - 1];
  }

  async write(chunk: Uint8Array, last: boolean): Promise<void> {
    const b = this.carry.length ? Buffer.concat([this.carry, chunk]) : chunk;
    let i = 0;
    while (i < b.length) {
      if (this.binLeft > 0) {
        const n = Math.min(this.binLeft, b.length - i);
        const top = this.top;
        if (top.kind === 'objdata') {
          top.hex.binary(b.subarray(i, i + n));
          top.alt.binary(b.subarray(i, i + n));
        } else if (top.kind === 'objskip') top.data.hex.binary(b.subarray(i, i + n));
        this.note(b, i, i + n);
        this.binLeft -= n;
        i += n;
        continue;
      }
      const c = b[i];
      if (c === 0x7b) {
        this.note(b, i, i + 1);
        this.open();
        i++;
      } else if (c === 0x7d) {
        this.note(b, i, i + 1);
        this.close(this.base + i);
        i++;
      } else if (c === 0x5c) {
        const used = this.control(b, i, last);
        if (used === 0) break;
        this.note(b, i, i + used);
        i += used;
      } else {
        let j = i;
        while (j < b.length && b[j] !== 0x7b && b[j] !== 0x7d && b[j] !== 0x5c) j++;
        this.note(b, i, j);
        this.text(b, i, j);
        i = j;
      }
    }
    this.carry = i < b.length ? b.slice(i) : new Uint8Array(0);
    this.base += i;
    await this.objects();
  }

  private async objects(): Promise<void> {
    for (const o of this.pending.splice(0)) await objectRules(this.ctx, o);
  }

  async end(): Promise<void> {
    // A file can end inside groups. Word reads what it has, and so does this.
    while (this.depth > 0) this.close(this.size);
    this.fields.finish();
    await this.objects();
    if (this.trailing > 0) this.ctx.report(C.Container, D.TrailingBytes, undefined, { bytes: this.trailing });
  }

  /** Counts bytes past the document's outermost group that are not white space or NUL. */
  private note(b: Uint8Array, from: number, to: number): void {
    if (this.closedAt < 0) return;
    for (let k = from; k < to; k++) if (b[k] > 0x20) this.trailing++;
  }

  private open(): void {
    this.depth++;
    this.skipChars = 0;
    const limit = this.ctx.settings.rtfDepth;
    if (this.depth > limit) throw new StopError(C.Limit, D.RtfDepth, `Groups nest deeper than ${limit}`, { limit });
  }

  private close(offset: number): void {
    this.skipChars = 0;
    this.starred = false;
    while (this.frames.length > 1 && this.top.depth >= this.depth) this.finishFrame(this.frames.pop() as Frame);
    if (this.depth === 0) return;
    this.depth--;
    if (this.depth === 0 && this.closedAt < 0) this.closedAt = offset;
  }

  /**
   * Opens a destination for the rest of the group. One opened in the same group as the last ends that one, as in
   * rtfobj, unless the last is a field or an object, whose group the destination belongs to.
   */
  private push(body: FrameBody): void {
    const top = this.top;
    if (this.frames.length > 1 && top.depth === this.depth && top.kind !== 'field' && top.kind !== 'object') this.finishFrame(this.frames.pop() as Frame);
    const limit = this.ctx.settings.rtfDepth;
    if (this.frames.length > limit) throw new StopError(C.Limit, D.RtfDepth, `More than ${limit} destinations are open at once`, { limit });
    this.frames.push({ ...body, depth: this.depth, uc: this.top.uc });
  }

  /** Reads a control word or symbol at `i`. Returns the bytes it takes, or 0 when the buffer ends inside it. */
  private control(b: Uint8Array, i: number, last: boolean): number {
    if (i + 1 >= b.length) return last ? 1 : 0;
    const c = b[i + 1];
    if (!isLetter(c)) {
      if (c === 0x27) {
        if (i + 3 >= b.length && !last) return 0;
        this.quoted(b, i + 2);
        return Math.min(4, b.length - i);
      }
      if (c === 0x2a) this.starred = true;
      else if (c === 0x5c || c === 0x7b || c === 0x7d) this.char(String.fromCharCode(c));
      else if (c === 0x0a || c === 0x0d) this.char('\n');
      else if (c === 0x7e) this.char(' ');
      return 2;
    }
    let j = i + 1;
    while (j < b.length && isLetter(b[j]) && j - i - 1 < MAX_WORD) j++;
    let k = j;
    if (k < b.length && b[k] === 0x2d && k + 1 < b.length && isDigit(b[k + 1])) k++;
    while (k < b.length && isDigit(b[k]) && k - j < MAX_WORD) k++;
    if (k >= b.length && !last) return 0;
    const name = latin1(b.subarray(i + 1, j));
    const digits = latin1(b.subarray(j, k));
    const param = digits === '' ? undefined : Number(digits);
    // A space ends the control word and belongs to it.
    const used = k - i + (k < b.length && b[k] === 0x20 ? 1 : 0);
    this.word(name, param, this.base + i);
    return used;
  }

  private word(name: string, param: number | undefined, offset: number): void {
    const starred = this.starred;
    this.starred = false;
    // \bin data is taken whole wherever it is, since its bytes may hold braces.
    if (name === 'bin') {
      // Word reads a negative or missing length as none.
      this.binLeft = param !== undefined && param > 0 ? Math.min(param, Number.MAX_SAFE_INTEGER) : 0;
      return;
    }
    const top = this.top;
    if (top.kind === 'skip') return;
    if (name === 'u' && param !== undefined) {
      this.skipChars = 0;
      this.char(String.fromCharCode(param < 0 ? param + 65536 : param & 0xffff));
      this.skipChars = top.uc;
      return;
    }
    this.skipChars = 0;
    if (name === 'uc') {
      top.uc = param ?? 1;
      return;
    }
    const object = this.object();
    if (INERT.has(name)) this.push({ kind: 'skip' });
    else if (name === 'field') {
      this.fields.begin();
      this.push({ kind: 'field' });
    } else if (name === 'fldinst') this.push({ kind: 'fldinst' });
    else if (name === 'fldrslt') {
      this.fields.separate();
      this.push({ kind: 'fldrslt' });
    } else if (name === 'object') this.push({ kind: 'object', object: { offset, kinds: new Set(), update: false, className: '' } });
    else if (name === 'objupdate') {
      if (object) object.update = true;
    } else if (OBJECT_KINDS.has(name)) object?.kinds.add(name);
    else if (name === 'objdata') {
      const limit = this.ctx.settings.decompressedBytes;
      this.push({
        kind: 'objdata',
        object: object ?? { offset, kinds: new Set(), update: false, className: '' },
        hex: new HexData(offset, limit),
        alt: new HexData(offset, limit),
        standalone: !object,
      });
    } else if (name === 'objclass') this.push({ kind: 'objclass', text: '', object });
    else if (name === 'template' || (name === 'hlsrc' && this.frames.some(f => f.kind === 'hl'))) this.push({ kind: name, text: '' });
    else if (name === 'hl') this.push({ kind: 'hl' });
    // Any other destination takes its own text. What it holds is still read for fields and objects.
    else if (DESTINATIONS.has(name)) this.push({ kind: 'document' });
    // Word skips a group that \* marks when it does not know its control word. Inside object data rtfobj keeps it, and
    // a sample Word opens hides its hex digits in groups such as {\*\par574 0}. Both readings of the data are kept.
    else if (starred && (top.kind === 'objdata' || top.kind === 'objskip')) this.push({ kind: 'objskip', data: top.kind === 'objdata' ? top : top.data });
    else if (starred) this.push({ kind: 'skip' });
  }

  /** The object of the innermost object group the reader is in. */
  private object(): RtfObject | undefined {
    for (let k = this.frames.length - 1; k > 0; k--) {
      const f = this.frames[k];
      if (f.kind === 'object') return f.object;
    }
    return undefined;
  }

  private quoted(b: Uint8Array, at: number): void {
    const top = this.top;
    if (top.kind === 'skip') return;
    this.starred = false;
    if (top.kind === 'objdata') {
      top.hex.quote();
      top.alt.quote();
      return;
    }
    if (top.kind === 'objskip') {
      top.data.hex.quote();
      return;
    }
    // A file can end before the two characters of a \'hh.
    const hi = hexValue(b[at] ?? -1);
    const lo = hexValue(b[at + 1] ?? -1);
    if (hi >= 0 && lo >= 0) this.char(String.fromCharCode((hi << 4) | lo));
  }

  private text(b: Uint8Array, from: number, to: number): void {
    const top = this.top;
    if (top.kind === 'skip') return;
    this.starred = false;
    if (top.kind === 'objdata' || top.kind === 'objskip') {
      const both = top.kind === 'objdata';
      const hex = both ? top.hex : top.data.hex;
      for (let k = from; k < to; k++) {
        const v = hexValue(b[k]);
        if (v < 0) continue;
        hex.digit(v);
        if (both) top.alt.digit(v);
      }
      return;
    }
    // Line ends in RTF text are not text.
    let s = '';
    for (let k = from; k < to; k++) {
      const c = b[k];
      if (c === 0x0a || c === 0x0d) continue;
      if (this.skipChars > 0) {
        this.skipChars--;
        continue;
      }
      s += String.fromCharCode(c);
    }
    if (s) this.char(s, true);
  }

  /** Text decoded from the file, for the destination being read. */
  private char(s: string, counted = false): void {
    if (!counted && this.skipChars > 0) {
      this.skipChars--;
      return;
    }
    const top = this.top;
    if (top.kind === 'fldinst') this.fields.text(s);
    else if ((top.kind === 'objclass' || top.kind === 'template' || top.kind === 'hlsrc') && top.text.length < MAX_TEXT) top.text += s;
  }

  private finishFrame(f: Frame): void {
    if (f.kind === 'field') this.fields.end();
    else if (f.kind === 'objclass') {
      if (f.object) f.object.className = f.text.trim();
    } else if (f.kind === 'template') {
      const path = f.text.trim();
      if (path) this.ctx.external(D.Template, path, 'template', { source: 'RTF template' }, 'template');
    } else if (f.kind === 'hlsrc') {
      const target = f.text.trim();
      if (target && !isFragment(target)) this.ctx.link(target, 'hyperlink', { source: 'RTF shape hyperlink' });
    } else if (f.kind === 'objdata') {
      f.object.data = f.hex.result();
      const alt = f.alt.result();
      if (!Buffer.from(alt).equals(Buffer.from(f.object.data))) f.object.alt = alt;
      this.decoded += f.object.data.length + (f.object.alt ? alt.length : 0);
      const total = this.ctx.settings.totalDecompressedBytes;
      if (this.decoded > total) throw new StopError(C.Limit, D.TotalDecompressedSize, `The objects decode past ${total} bytes`, { limit: total });
      // An objdata outside any object is an object of its own, as rtfobj reads it.
      if (f.standalone) this.pending.push(f.object);
    } else if (f.kind === 'object') this.pending.push(f.object);
  }
}

/** One embedded or linked object: what its OLE 1.0 header names, and what its data holds. */
async function objectRules(ctx: RunContext, o: RtfObject): Promise<void> {
  const location = `object at offset ${o.offset}`;
  const kinds = [...o.kinds];
  const base: Record<string, string | number> = { object: location, ...(kinds.length ? { kind: kinds.join(' ') } : {}) };
  if (o.update) ctx.report(C.Action, D.ObjectUpdate, location, base);
  if (o.kinds.has('objocx')) {
    ctx.report(C.Control, D.ActiveX, location, { ...base, ...(o.className ? { className: clip(o.className) } : {}) });
    return;
  }
  await objectData(ctx, o, o.data, location, base);
  // The data read without the groups \* marks, where that reads differently, is the object a reader that skips them sees.
  if (o.alt) await objectData(ctx, o, o.alt, location, { ...base, reading: 'without the groups \\* marks' });
}

/** An object's data: its OLE 1.0 header, and the Packager data or storage it holds. */
async function objectData(ctx: RunContext, o: RtfObject, bytes: Uint8Array | undefined, location: string, base: Record<string, string | number>): Promise<void> {
  const ole1 = bytes?.length ? parseOle1(bytes) : undefined;
  const className = (ole1?.className || o.className).replace(/\0+$/, '');
  const data: Record<string, string | number> = { ...base, ...(className ? { className: clip(className) } : {}) };
  const link = { object: location, ...(className ? { className: clip(className) } : {}) };
  let equation = isEquationClass(className);
  if (ole1?.formatId === 1 || (!bytes?.length && (o.kinds.has('objlink') || o.kinds.has('objautlink')))) {
    const target = ole1?.topicName ?? '';
    if (target) ctx.external(D.Object, target, location, link, 'linked object');
    else ctx.report(C.External, D.Object, location, data);
    return;
  }
  if (!bytes?.length) return;
  if (!ole1) data.reason = 'its data is not an OLE 1.0 object';
  const native = ole1?.data;
  if (native && className.toLowerCase() === 'package') {
    const pkg = parsePackage(native);
    if (pkg) Object.assign(data, packageData(pkg));
  } else if (native && native.length >= 8 && CFB_SIGNATURE.every((v, i) => native[i] === v)) {
    let cfb: CompoundFile | undefined;
    try {
      cfb = await CompoundFile.open(bufferSource(native), ctx.settings, ctx.settings.deadline);
    } catch (e) {
      if (!(e instanceof StopError) || e.category !== C.Container) throw e;
      data.reason = `its storage does not read: ${e.message}`;
    }
    if (cfb) {
      nestedNotes(ctx, cfb, location);
      const d = await describeObject(cfb, cfb.root, location);
      if (d.clsid) data.clsid = d.clsid;
      if (d.progId) data.progId = clip(d.progId);
      if (d.pkg) Object.assign(data, packageData(d.pkg));
      equation ||= d.equation;
      await compoundRules(ctx, cfb, location, false);
      if (d.linked) {
        linkedObject(ctx, d, location, link);
        if (equation) ctx.report(C.Embedded, D.Equation, location, data);
        return;
      }
    }
  }
  // OLE2Link is the class of a linked object, CVE-2017-0199's. Its storage names the target, and when the storage does
  // not read, the class still says the object loads from outside.
  if (className.toLowerCase() === 'ole2link') ctx.report(C.External, D.Object, location, data);
  else ctx.report(C.Embedded, D.OleObject, location, data);
  if (equation) ctx.report(C.Embedded, D.Equation, location, data);
}
