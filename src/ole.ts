import type { CfbEntry, CompoundFile } from './cfb';
import type { RunContext } from './context';
import { StopError } from './errors';
import type { Format } from './formats';
import { messageRules } from './msg';
import { compoundRules } from './objects';
import { openPackage, unlockWorkbook } from './offcrypto';
import { inspectPackage } from './package';
import { CompoundPlan } from './plan';
import { powerPointRules } from './ppt';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { visioRules } from './vsd';
import { isWordTemplate, wordRules } from './word';
import { excelRules, workbookKind } from './xls';
import { ZipArchive } from './zip';

const f = (format: string, mimeType: string): Format => ({ format, mimeType });

/** The families a compound file can hold, by the stream or storage that marks each. */
export type Family = 'word' | 'excel' | 'powerpoint' | 'visio' | 'outlook' | 'publisher' | 'project' | 'encrypted';

const upper = (e: CfbEntry) => e.name.toUpperCase();
const PUBLISHER_CLSIDS = new Set(['00021200-0000-0000-00C0-000000000046', '00021201-0000-0000-00C0-000000000046']);

/** What the root of a compound file holds, read by the names of its children, ignoring case as MS-CFB does. */
export function families(cfb: CompoundFile): Family[] {
  const names = new Map(cfb.root.children.map(c => [upper(c), c]));
  const stream = (n: string) => names.get(n)?.kind === 'stream';
  const out: Family[] = [];
  if (stream('ENCRYPTIONINFO') && stream('ENCRYPTEDPACKAGE')) out.push('encrypted');
  if (stream('WORDDOCUMENT')) out.push('word');
  if (stream('WORKBOOK') || stream('BOOK')) out.push('excel');
  if (stream('POWERPOINT DOCUMENT')) out.push('powerpoint');
  if (stream('VISIODOCUMENT')) out.push('visio');
  if (stream('__PROPERTIES_VERSION1.0')) out.push('outlook');
  // Publisher 98 and later keep their content in a Quill storage. Publisher 95 has none, and names itself by class id.
  if (names.get('QUILL')?.kind === 'storage' || (PUBLISHER_CLSIDS.has(cfb.root.clsid) && stream('CONTENTS'))) out.push('publisher');
  // Project keeps its data in a storage named by three spaces and a number, beside a Props stream.
  if ([...names.keys()].some(n => /^PROPS\d*$/.test(n)) && cfb.root.children.some(c => c.kind === 'storage' && /^ {3}\d+$/.test(c.name))) out.push('project');
  return out;
}

/**
 * The family whose format a compound file holds, by the streams and storages at its root. Throws for a root that
 * holds no Office document, or the documents of two families.
 */
export function mainFamily(cfb: CompoundFile): Family {
  const found = families(cfb);
  // Word 2007 writes a Word document beside an encrypted package, for readers that cannot decrypt it.
  const main = found.filter(x => !(x === 'word' && found.includes('encrypted')));
  if (main.length > 1) throw new StopError(C.Package, D.AmbiguousFormat, `The compound file holds ${main.join(' and ')} streams`, { reason: `${main.join(' and ')} streams` });
  if (main.length === 0) throw new StopError(C.Package, D.Unrecognized, 'The compound file holds no Office document stream', { reason: 'a compound file of another kind' });
  return main[0];
}

/** A stream at the root of a compound file, by name, ignoring case. */
const rootStream = (cfb: CompoundFile, name: string) => {
  const e = cfb.child(cfb.root, name);
  return e?.kind === 'stream' ? e : undefined;
};

/**
 * The format of a compound file of `family`. A Word document is a template when its FIB says so, and a workbook a
 * template or an add-in when its globals do; both are read without a password. An encrypted package names no format
 * until it is decrypted, so for one this is undefined.
 */
export async function compoundFormat(cfb: CompoundFile, family: Family, deadline?: number): Promise<Format | undefined> {
  switch (family) {
    case 'word': {
      const word = rootStream(cfb, 'WordDocument');
      return word && (await isWordTemplate(cfb, word)) ? f('dot', 'application/msword') : f('doc', 'application/msword');
    }
    case 'excel': {
      // Excel 97 to 2003 reads Workbook, and Excel 5.0 and 95 read Book. A file can hold both.
      let template = false;
      let addIn = false;
      for (const name of ['Workbook', 'Book']) {
        const s = rootStream(cfb, name);
        if (!s) continue;
        const kind = await workbookKind(cfb, s, deadline);
        template ||= kind.template;
        addIn ||= kind.addIn;
      }
      return addIn ? f('xla', 'application/vnd.ms-excel') : template ? f('xlt', 'application/vnd.ms-excel') : f('xls', 'application/vnd.ms-excel');
    }
    case 'powerpoint':
      return f('ppt', 'application/vnd.ms-powerpoint');
    case 'visio':
      return f('vsd', 'application/vnd.visio');
    case 'outlook':
      return cfb.root.clsid === '0006F046-0000-0000-C000-000000000046' ? f('oft', 'application/vnd.ms-outlook') : f('msg', 'application/vnd.ms-outlook');
    case 'publisher':
      return f('pub', 'application/x-mspublisher');
    case 'project':
      return f('mpp', 'application/vnd.ms-project');
    default:
      // An encrypted package names its format only once it is decrypted.
      return undefined;
  }
}

/**
 * Decrypts the package of an encrypted compound file with the run's passwords, then VelvetSweatshop, and reads it as
 * a ZIP archive. Throws ENCRYPTED/PASSWORD_REQUIRED when no password opens it.
 */
export async function openEncryptedPackage(ctx: RunContext, cfb: CompoundFile): Promise<ZipArchive> {
  const info = rootStream(cfb, 'EncryptionInfo');
  const pkg = rootStream(cfb, 'EncryptedPackage');
  if (!info || !pkg) throw new StopError(C.Package, D.Unrecognized, 'The compound file holds no encrypted package', { reason: 'no EncryptionInfo and EncryptedPackage streams' });
  const { source } = await openPackage(ctx, cfb, info, pkg);
  return ZipArchive.open(source, ctx.settings, ctx.settings.deadline);
}

/**
 * Inspects a compound file: the rules every compound file gets, over its whole tree, and then those of the format
 * its root names.
 */
export async function inspectCompound(ctx: RunContext, cfb: CompoundFile): Promise<void> {
  const family = mainFamily(cfb);
  ctx.format = await compoundFormat(cfb, family, ctx.settings.deadline);
  const plan = new CompoundPlan();
  if (family !== 'encrypted') ctx.run.output = { kind: 'cfb', cfb, plan };
  await compoundRules(ctx, cfb, '', false, plan);
  const stream = (name: string) => rootStream(cfb, name);
  // Beside an encrypted package, Word 2007 writes a WordDocument stream for older readers, of zeros in the files at
  // hand. Those readers read it if it holds a document, so it gets the Word rules when it does.
  const word = stream('WordDocument');
  if (word) await wordRules(ctx, cfb, word, family === 'encrypted');
  if (family === 'excel') {
    for (const name of ['Workbook', 'Book']) {
      const s = stream(name);
      if (s) await excelRules(ctx, cfb, s, unlockWorkbook);
    }
    // Excel keeps the data of a workbook's ActiveX controls in one Ctls stream.
    const controls = stream('Ctls');
    if (controls && controls.size > 0 && ctx.report(C.Control, D.ActiveX, controls.name, { stream: controls.name }).action === 'strip') plan.drop.add(controls);
  }
  if (family === 'outlook') await messageRules(ctx, cfb, cfb.root, 0, plan);
  const ppt = stream('PowerPoint Document');
  if (family === 'powerpoint' && ppt) await powerPointRules(ctx, cfb, ppt);
  const visio = stream('VisioDocument');
  if (family === 'visio' && visio) await visioRules(ctx, cfb, visio);
  if (family === 'encrypted') {
    // The package goes through the OOXML rules once a password opens it, and names the format.
    const zip = await openEncryptedPackage(ctx, cfb);
    const location = stream('EncryptedPackage')?.name ?? 'EncryptedPackage';
    for (const n of zip.notes) ctx.report(C.Container, n.detail, location, { bytes: n.bytes });
    // The rewrite writes the package decrypted, since office-defuse does not encrypt.
    const zctx = ctx.withZip(zip);
    await inspectPackage(zctx);
    if (ctx.format) ctx.run.output = { kind: 'zip', zip, plan: zctx.plan, format: ctx.format };
  }
}
