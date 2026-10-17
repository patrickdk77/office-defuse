import { type CfbEntry, type CompoundFile, showName } from './cfb';
import { parseCompObj } from './oleobj';

/**
 * The designer storage of a VBA user form (MS-OVBA 2.2.10), which holds the form and its controls as Forms 2.0
 * persists them (MS-OFORMS). A form site names its control's class by an index into Forms 2.0's own controls, or
 * into the form's class table, which may name any ActiveX control. A form whose class tables are empty therefore
 * holds Forms 2.0 controls only.
 */

/** The Forms 2.0 form, frame and multipage, the classes a designer and the storages under it are written as. */
const FORM_CLASSES = new Set(['C62A69F0-16DC-11CE-9E98-00AA00574A4F', '6E182020-F460-11CE-9BCD-00AA00608E01', '46E31370-3F7A-11CE-BED6-00AA00611080']);
const FORM_PROGIDS = new Set(['forms.form.1', 'forms.frame.1', 'forms.multipage.1']);
/** The UserForm designer, which \x03VBFrame names. */
const USER_FORM = 'C62A69F0-16DC-11CE-9E98-00AA00574A4F';
const NULL_CLSID = '00000000-0000-0000-0000-000000000000';
/** StdFont's GUID, as GuidAndFont writes it, which a StdFont structure follows; any other is TextProps. */
const STD_FONT = '0352e30b918fce119de300aa004bb851';
/** FORM_FLAG_DONTSAVECLASSTABLE, in BooleanProperties. */
const NO_CLASS_TABLE = 0x8000;

class FormError extends Error {}

/** The number of classes in the class table of a FormControl (MS-OFORMS 2.2.10.1), the `f` stream. */
export function classTableCount(f: Uint8Array): number {
  const b = Buffer.from(f.buffer, f.byteOffset, f.byteLength);
  const need = (at: number, n: number) => {
    if (at + n > b.length) throw new FormError('a form stream cut short');
  };
  need(0, 8);
  if (b[1] !== 0x04) throw new FormError(`a form of version ${b[1]}.${b[0]}`);
  const cbForm = b.readUInt16LE(2);
  const mask = b.readUInt32LE(4);
  // BooleanProperties follows BackColor, ForeColor and NextAvailableID, each 4 bytes when present.
  let booleans = 0x04;
  if (mask & (1 << 6)) {
    let at = 8;
    for (const bit of [1, 2, 3]) if (mask & (1 << bit)) at += 4;
    need(at, 4);
    booleans = b.readUInt32LE(at);
  }
  // cbForm counts the property mask and the data blocks; the stream data follows them.
  let p = 4 + cbForm;
  const picture = () => {
    need(p, 24);
    p += 16;
    if (b.readUInt32LE(p) !== 0x746c) throw new FormError('a form picture without its preamble');
    p += 8 + b.readUInt32LE(p + 4);
  };
  if (mask & (1 << 15)) picture();
  if (mask & (1 << 20)) {
    need(p, 16);
    const stdFont = b.subarray(p, p + 16).toString('hex') === STD_FONT;
    p += 16;
    if (stdFont) {
      need(p, 11);
      p += 11 + b[p + 10];
    } else {
      need(p, 4);
      p += 4 + b.readUInt16LE(p + 2);
    }
  }
  if (mask & (1 << 21)) picture();
  if (booleans & NO_CLASS_TABLE) return 0;
  need(p, 2);
  return b.readUInt16LE(p);
}

/**
 * Why a user form's designer storage may hold something other than Forms 2.0 controls, or undefined when it holds
 * only those: the UserForm designer, storages of Forms 2.0 classes, and forms whose class tables are empty.
 */
export async function designerProblem(cfb: CompoundFile, storage: CfbEntry, limit: number): Promise<string | undefined> {
  const frame = storage.children.find(c => c.kind === 'stream' && c.name === '\x03VBFrame');
  if (!frame) return 'it has no \\x03VBFrame stream';
  if (frame.size > limit) return 'its \\x03VBFrame stream is past the limit';
  const designer = /Begin\s+\{([0-9A-Fa-f-]{36})\}/.exec(Buffer.from(await cfb.bytes(frame)).toString('latin1'))?.[1]?.toUpperCase();
  if (designer !== USER_FORM) return `its designer is ${designer ?? 'not named'}, not the UserForm designer`;
  const walk = async (s: CfbEntry, top: boolean): Promise<string | undefined> => {
    for (const c of s.children) {
      if (c.kind === 'storage') {
        if (!/^i\d+$/.test(c.name)) return `it holds a storage named ${JSON.stringify(c.name)}`;
        const cls = c.clsid === NULL_CLSID ? '' : c.clsid;
        if (!cls) {
          const compObj = c.children.find(x => x.kind === 'stream' && x.name === '\x01CompObj');
          const progId = compObj && compObj.size <= 65536 ? parseCompObj(await cfb.bytes(compObj)).progId : undefined;
          if (!progId || !FORM_PROGIDS.has(progId.trim().toLowerCase())) return `it holds a control of class ${progId ?? 'not named'}`;
        } else if (!FORM_CLASSES.has(cls)) return `it holds a control of class ${cls}`;
        const inner = await walk(c, false);
        if (inner) return inner;
        continue;
      }
      if (c.name === 'f') {
        if (c.size > limit) return 'a form stream is past the limit';
        try {
          if (classTableCount(await cfb.bytes(c)) > 0) return 'a form names controls in its class table, which may be any ActiveX control';
        } catch (e) {
          if (!(e instanceof FormError)) throw e;
          return e.message;
        }
      } else if (!['o', 'x', '\x01CompObj'].includes(c.name) && !(top && c.name === '\x03VBFrame')) return `it holds a stream named ${JSON.stringify(c.name)}`;
    }
    return undefined;
  };
  return walk(storage, true);
}

/**
 * The properties a Forms 2.0 control persists, in the order of its property mask (MS-OFORMS 2.2). Each takes the
 * next bit. A number sits in the data block, aligned to its size; a string or a pair in the extra data block; a
 * picture in the stream data; a flag holds no data. A bit MS-OFORMS leaves undefined must be clear.
 */
type Prop = 'u8' | 'u16' | 'u32' | 'pair' | 'string' | 'picture' | 'flag' | 'undefined';

const BUTTON: Prop[] = ['u32', 'u32', 'u32', 'string', 'u32', 'pair', 'u8', 'picture', 'u16', 'flag', 'picture'];
const LABEL: Prop[] = ['u32', 'u32', 'u32', 'string', 'u32', 'pair', 'u8', 'u32', 'u16', 'u16', 'picture', 'u16', 'picture'];
// TextBox, ListBox, ComboBox, CheckBox, OptionButton and ToggleButton. The 16th property counts the column infos.
const MORPH: Prop[] = [
  ...['u32', 'u32', 'u32', 'u32', 'u8', 'u8', 'u8', 'u8', 'pair', 'u16', 'u32', 'u16', 'u16', 'u16', 'u16', 'u16'],
  ...['u8', 'u8', 'u8', 'undefined', 'u8', 'u8', 'string', 'string', 'u32', 'u32', 'u32', 'picture', 'picture', 'u16', 'undefined', 'flag', 'string'],
] as Prop[];
const MORPH_COLUMN_INFOS = 15;
const SPIN: Prop[] = ['u32', 'u32', 'u32', 'pair', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'picture', 'u8'];
const SCROLL: Prop[] = ['u32', 'u32', 'u32', 'pair', 'u8', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u32', 'u16', 'u32', 'picture'];
/** TextProps (MS-OFORMS 2.3.1), the font a control with text writes after its own properties. */
const TEXT_PROPS: Prop[] = ['string', 'u32', 'u32', 'u32', 'u8', 'u8', 'u8', 'u16'];
const COLUMN_INFO: Prop[] = ['u32'];

interface ControlLayout {
  props: Prop[];
  /** The property mask is 64 bits wide. */
  wide: boolean;
  text: boolean;
}

/** The Forms 2.0 controls office-defuse reads, by class id. Frames, multipages and tab strips hold other controls and are not here. */
export const FORMS_CONTROLS: ReadonlyMap<string, ControlLayout> = new Map([
  ['D7053240-CE69-11CD-A777-00DD01143C57', { props: BUTTON, wide: false, text: true }],
  ['978C9E23-D4B0-11CE-BF2D-00AA003F40D0', { props: LABEL, wide: false, text: true }],
  ...['8BD21D10', '8BD21D20', '8BD21D30', '8BD21D40', '8BD21D50', '8BD21D60'].map((p): [string, ControlLayout] => [`${p}-EC42-11CE-9E0D-00AA006002F3`, { props: MORPH, wide: true, text: true }]),
  ['79176FB0-B7F2-11CE-97EF-00AA006D2776', { props: SPIN, wide: false, text: false }],
  ['DFD181E0-5E2F-11CE-A449-00AA004A803D', { props: SCROLL, wide: false, text: false }],
]);

/** The ProgIDs of those controls, which a property bag or CompObj may name. */
export const FORMS_PROGIDS: ReadonlyMap<string, string> = new Map([
  ['forms.commandbutton.1', 'D7053240-CE69-11CD-A777-00DD01143C57'],
  ['forms.label.1', '978C9E23-D4B0-11CE-BF2D-00AA003F40D0'],
  ['forms.textbox.1', '8BD21D10-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.listbox.1', '8BD21D20-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.combobox.1', '8BD21D30-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.checkbox.1', '8BD21D40-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.optionbutton.1', '8BD21D50-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.togglebutton.1', '8BD21D60-EC42-11CE-9E0D-00AA006002F3'],
  ['forms.spinbutton.1', '79176FB0-B7F2-11CE-97EF-00AA006D2776'],
  ['forms.scrollbar.1', 'DFD181E0-5E2F-11CE-A449-00AA004A803D'],
]);

/** Strings a control persists are at most this many bytes, as Forms 2.0 and LibreOffice read them. */
const MAX_STRING = 0x20000;

/**
 * Reads one property block at `start`: version 2.0, the block size, the property mask and the properties it names.
 * Returns where the block ends and the numbers read, by property index. Throws FormError on a picture, a bit of an
 * undefined property, a bit past the last property, or a property that runs past the block.
 */
function readBlock(d: Buffer, start: number, props: Prop[], wide: boolean): { end: number; values: Map<number, number> } {
  const need = (at: number, n: number, what: string) => {
    if (at + n > d.length) throw new FormError(`${what} cut short`);
  };
  need(start, wide ? 12 : 8, 'a control property block');
  if (d[start] !== 0 || d[start + 1] !== 2) throw new FormError(`a control property block of version ${d[start + 1]}.${d[start]}`);
  const end = start + 4 + d.readUInt16LE(start + 2);
  if (end > d.length) throw new FormError('a control property block that runs past its stream');
  let low = d.readUInt32LE(start + 4);
  let high = wide ? d.readUInt32LE(start + 8) : 0;
  let pos = start + (wide ? 12 : 8);
  const align = (n: number) => {
    pos += (n - ((pos - start) % n)) % n;
  };
  const values = new Map<number, number>();
  const extra: Array<{ kind: 'pair' } | { kind: 'string'; size: number }> = [];
  props.forEach((p, i) => {
    const has = i < 32 ? (low >>> i) & 1 : (high >>> (i - 32)) & 1;
    if (i < 32) low &= ~(1 << i);
    else high &= ~(1 << (i - 32));
    if (!has || p === 'flag') return;
    if (p === 'undefined') throw new FormError('a control property MS-OFORMS leaves undefined');
    if (p === 'picture') throw new FormError('a control that holds a picture');
    if (p === 'pair') extra.push({ kind: 'pair' });
    else if (p === 'string') {
      align(4);
      need(pos, 4, 'a control string size');
      extra.push({ kind: 'string', size: d.readUInt32LE(pos) });
      pos += 4;
    } else {
      const n = p === 'u8' ? 1 : p === 'u16' ? 2 : 4;
      align(n);
      need(pos, n, 'a control property');
      values.set(i, n === 1 ? d[pos] : n === 2 ? d.readUInt16LE(pos) : d.readUInt32LE(pos));
      pos += n;
    }
  });
  if (low !== 0 || high !== 0) throw new FormError('a control property mask with bits past its last property');
  align(4);
  for (const x of extra) {
    const n = x.kind === 'pair' ? 8 : x.size & 0x7fffffff;
    if (n > MAX_STRING) throw new FormError('a control string past 128 KiB');
    pos += n;
    align(4);
  }
  if (pos > end) throw new FormError('control properties that run past their block');
  return { end, values };
}

/**
 * Why the persisted data of a Forms 2.0 control may not be what Forms 2.0 writes, or undefined when it is: the
 * control's property block, then its TextProps for a control with text, then the column infos a list control counts,
 * and nothing after. A picture or a mouse icon counts as a problem, since its bytes are a picture file Office decodes
 * and office-defuse cannot rewrite in place.
 */
export function controlProblem(classId: string, data: Uint8Array): string | undefined {
  const layout = FORMS_CONTROLS.get(classId.toUpperCase());
  if (!layout) return `its class ${classId} is not a Forms 2.0 control office-defuse reads`;
  const d = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  try {
    const own = readBlock(d, 0, layout.props, layout.wide);
    let pos = own.end;
    if (layout.text) pos = readBlock(d, pos, TEXT_PROPS, false).end;
    const columns = layout.props === MORPH ? (own.values.get(MORPH_COLUMN_INFOS) ?? 0) : 0;
    for (let k = 0; k < columns; k++) pos = readBlock(d, pos, COLUMN_INFO, false).end;
    if (pos !== d.length) return `${d.length - pos} bytes after the control's properties`;
    return undefined;
  } catch (e) {
    if (!(e instanceof FormError)) throw e;
    return e.message;
  }
}

/** Property bag names that load a picture, which a Forms 2.0 control takes from a part of its own. */
const BAG_PICTURES = new Set(['picture', 'mouseicon']);

/**
 * Why a Forms 2.0 control persisted as a property bag in its activeX part may hold more than plain values, or
 * undefined when it does not. `props` are the bag's name and value pairs, `nested` says ax:ocx held any other element,
 * inside an ax:ocxPr or beside one, such as a font or a picture.
 */
export function bagProblem(classId: string, props: ReadonlyArray<{ name: string; value: string }>, nested: boolean): string | undefined {
  if (!FORMS_CONTROLS.has(classId.toUpperCase())) return `its class ${classId} is not a Forms 2.0 control office-defuse reads`;
  if (nested) return 'it holds elements besides plain ax:ocxPr values, such as a font or a picture';
  const picture = props.find(p => BAG_PICTURES.has(p.name.toLowerCase()));
  if (picture) return `its ${picture.name} property loads a picture`;
  return undefined;
}

/** Streams a control's storage holds in an OOXML binary part: its class and its properties. */
export const OOXML_CONTROL_STREAMS: ReadonlySet<string> = new Set(['\x01CompObj', '\x03ObjInfo', 'contents']);
/** In a Word document, the storage also holds the control's name and its picture for printing. */
export const WORD_CONTROL_STREAMS: ReadonlySet<string> = new Set([...OOXML_CONTROL_STREAMS, '\x03OCXNAME', '\x03PRINT']);
/** A Forms 2.0 control's contents stream, without pictures, is far smaller. */
const MAX_CONTENTS = 65536;

/**
 * Why a control's storage may not hold a Forms 2.0 control as Office persists one, or undefined when it does. Its
 * class is the storage's class id, or its CompObj's ProgID when that is null, and must be `classId` when given. It
 * holds only the `allowed` streams and a contents stream that `controlProblem` passes.
 */
export async function controlStorageProblem(cfb: CompoundFile, storage: CfbEntry, allowed: ReadonlySet<string>, classId?: string): Promise<string | undefined> {
  const stream = (name: string) => storage.children.find(c => c.kind === 'stream' && c.name === name);
  let cls = storage.clsid === NULL_CLSID ? undefined : storage.clsid;
  if (!cls) {
    const compObj = stream('\x01CompObj');
    const progId = compObj && compObj.size <= MAX_CONTENTS ? parseCompObj(await cfb.bytes(compObj)).progId : undefined;
    cls = progId ? FORMS_PROGIDS.get(progId.trim().toLowerCase()) : undefined;
    if (!cls) return `its class ${progId ?? 'is not named'}${progId ? ' is not a Forms 2.0 control office-defuse reads' : ''}`;
  }
  if (classId && cls !== classId) return 'its storage names another class';
  if (!FORMS_CONTROLS.has(cls)) return `its class ${cls} is not a Forms 2.0 control office-defuse reads`;
  for (const c of storage.children) if (c.kind !== 'stream' || !allowed.has(c.name)) return `it holds a ${c.kind} named ${JSON.stringify(showName(c.name))}`;
  const contents = stream('contents');
  if (!contents) return 'it has no contents stream';
  if (contents.size > MAX_CONTENTS) return 'its contents stream is larger than a Forms 2.0 control without pictures';
  return controlProblem(cls, await cfb.bytes(contents));
}
