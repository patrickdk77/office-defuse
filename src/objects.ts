import { type CfbEntry, type CompoundFile, entryPath } from './cfb';
import { clip, type RunContext } from './context';
import { type PackagedFile, parseCompObj, parseOleStream, parsePackage } from './oleobj';
import { OfficeCategory as C, OfficeDetail as D } from './types';

/** The class ids of Equation Editor 3.0 and its kin, which CVE-2017-11882 and CVE-2018-0802 exploit, as oletools lists them. */
const EQUATION_CLSIDS = new Set(['0002CE02-0000-0000-C000-000000000046', '00021700-0000-0000-C000-000000000046', '0003000B-0000-0000-C000-000000000046', '0004A6B0-0000-0000-C000-000000000046']);
/** Equation Editor 3.0's ProgID, which rtfobj matches as a prefix. */
export const isEquationClass = (name: string) => name.trim().toLowerCase().startsWith('equation.3');

/** Streams that only an OLE object's storage holds. */
const OBJECT_STREAMS = new Set(['\x01OLE', '\x01COMPOBJ', '\x01OLE10NATIVE', '\x03OBJINFO']);
/** Streams that only an ActiveX control's storage holds. */
const CONTROL_STREAMS = new Set(['\x03OCXNAME', '\x03OCXDATA']);
/** A CompObj, OLE or ObjInfo stream is small. More is not read. */
const SMALL = 65536;
/** Packager's header is short, and its names are at most 1,024 bytes each. */
const PACKAGE_HEAD = 4096;

const upper = (e: CfbEntry) => e.name.toUpperCase();

/** Reports what the reader of a compound file inside this one noted, at the place it sits. */
export function nestedNotes(ctx: RunContext, cfb: CompoundFile, location: string): void {
  for (const n of cfb.notes) ctx.report(C.Container, n.detail, location, { bytes: n.bytes });
}

/** Where a finding about an entry points: `prefix`, then the entry's path. */
const at = (prefix: string, e: CfbEntry) => [prefix, entryPath(e)].filter(Boolean).join(' > ');

/**
 * The rules every compound file gets: VBA projects, embedded objects and controls, wherever they sit in the tree.
 * `prefix` names where the file sits when it is inside another, and `rootIsObject` says that its root is itself an
 * OLE object's storage, as an object in a presentation or an RTF file is.
 */
export async function compoundRules(ctx: RunContext, cfb: CompoundFile, prefix: string, rootIsObject: boolean): Promise<void> {
  const projects = new Set<CfbEntry>();
  for (const e of cfb.entries) {
    if (e.kind === 'stream') continue;
    const child = (name: string) => e.children.find(c => upper(c) === name);
    // A VBA project: a VBA storage with a dir stream, which names the modules (MS-OVBA 2.2).
    const vba = child('VBA');
    if (vba?.kind === 'storage' && vba.children.some(c => c.kind === 'stream' && upper(c) === 'DIR')) {
      projects.add(e);
      ctx.report(C.Macro, D.VbaProject, at(prefix, vba), { storage: at(prefix, vba) });
    }
    if (e.kind === 'storage' && (upper(e) === '_VBA_PROJECT_CUR' || upper(e) === 'MACROS') && e.children.length && !projects.has(e)) {
      // Excel's and Word's project storages, without the dir stream a whole project has.
      ctx.report(C.Macro, D.VbaProject, at(prefix, e), { storage: at(prefix, e) });
    }
    const names = new Set(e.children.map(upper));
    const parent = e.parent ? upper(e.parent) : '';
    const isObject =
      e.kind === 'root'
        ? rootIsObject
        : [...names].some(n => OBJECT_STREAMS.has(n) || CONTROL_STREAMS.has(n)) ||
          (parent === 'OBJECTPOOL' && e.parent?.parent?.kind === 'root') ||
          (e.parent?.kind === 'root' && /^MBD[0-9A-F]{8}$/i.test(e.name));
    if (isObject) await objectRules(ctx, cfb, e, at(prefix, e));
  }
}

/** What an OLE object's storage says about the object. */
export interface ObjectInfo {
  clsid: string;
  progId?: string;
  userType?: string;
  /** An ActiveX control's storage, with the control's name. */
  control: boolean;
  controlName?: string;
  linked: boolean;
  /** What a linked object loads, or ''. */
  target: string;
  /** The target is a URL moniker's, which loads from outside however it is spelled. */
  url: boolean;
  /** A Packager object's file, held or linked to. */
  pkg?: PackagedFile;
  equation: boolean;
}

/** Reads an OLE object's storage: its class, its CompObj and OLE streams, and Packager's native data. */
export async function describeObject(cfb: CompoundFile, e: CfbEntry, _location: string): Promise<ObjectInfo> {
  const child = (name: string) => e.children.find(c => c.kind === 'stream' && upper(c) === name);
  const small = async (name: string) => {
    const s = child(name);
    return s && s.size <= SMALL ? cfb.bytes(s) : undefined;
  };
  const clsid = e.clsid === '00000000-0000-0000-0000-000000000000' ? '' : e.clsid;
  const compObjBytes = await small('\x01COMPOBJ');
  const comp = compObjBytes ? parseCompObj(compObjBytes) : {};
  const info: ObjectInfo = { clsid, progId: comp.progId, userType: comp.userType, control: false, linked: false, target: '', url: false, equation: false };
  info.equation = EQUATION_CLSIDS.has(clsid) || isEquationClass(comp.progId ?? '');
  if ([...CONTROL_STREAMS].some(n => child(n))) {
    const name = await small('\x03OCXNAME');
    info.control = true;
    if (name) info.controlName = Buffer.from(name).toString('utf16le').replace(/\0+$/, '');
    return info;
  }
  const oleBytes = await small('\x01OLE');
  const ole = oleBytes ? parseOleStream(oleBytes) : { linked: false, target: '', url: false };
  Object.assign(info, ole);
  const native = child('\x01OLE10NATIVE');
  if (native) {
    // The stream starts with the native data's size, then Packager's data.
    info.pkg = parsePackage(await cfb.source(native).read(4, PACKAGE_HEAD));
  }
  return info;
}

/** One OLE object's storage: a control, a linked object, or an embedded one, with what it holds. */
async function objectRules(ctx: RunContext, cfb: CompoundFile, e: CfbEntry, location: string): Promise<void> {
  const d = await describeObject(cfb, e, location);
  const data: Record<string, string | number> = { storage: location };
  if (d.clsid) data.clsid = d.clsid;
  if (d.progId) data.progId = clip(d.progId);
  if (d.userType) data.className = clip(d.userType);
  if (d.control) {
    ctx.report(C.Control, D.ActiveX, location, d.controlName ? { ...data, name: clip(d.controlName) } : data);
    return;
  }
  if (d.pkg) {
    Object.assign(data, packageData(d.pkg));
    if (!d.pkg.embedded && d.pkg.sourcePath) ctx.external(D.Object, d.pkg.sourcePath, location, { storage: location }, 'package link');
  }
  if (d.linked) linkedObject(ctx, d, location, { storage: location, ...(d.clsid ? { clsid: d.clsid } : {}) });
  else ctx.report(C.Embedded, D.OleObject, location, data);
  if (d.equation) ctx.report(C.Embedded, D.Equation, location, data);
}

/**
 * Reports a linked object. A URL moniker's target loads from outside the file whatever its text looks like: a
 * CVE-2017-0199 sample puts a character with a stray high byte where the colon of https: goes.
 */
export function linkedObject(ctx: RunContext, d: { target: string; url: boolean }, location: string, data: Record<string, string>): void {
  if (d.target && !d.url) ctx.external(D.Object, d.target, location, data, 'linked object');
  else ctx.report(C.External, D.Object, location, d.target ? { ...data, target: clip(d.target, 2000) } : data);
}

/** What a finding shows of a Packager object's file: its name and paths, and its size when the object holds it. */
export function packageData(p: PackagedFile): Record<string, string | number> {
  return { fileName: clip(p.fileName), sourcePath: clip(p.sourcePath, 2000), tempPath: clip(p.tempPath, 2000), ...(p.embedded ? { size: p.size } : {}) };
}
