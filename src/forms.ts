import type { CfbEntry, CompoundFile } from './cfb';
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
