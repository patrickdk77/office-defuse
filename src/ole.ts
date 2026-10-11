import type { CfbEntry, CompoundFile } from './cfb';
import type { RunContext } from './context';
import { StopError } from './errors';
import type { Format } from './formats';
import { messageRules } from './msg';
import { compoundRules } from './objects';
import { openPackage, unlockWorkbook } from './offcrypto';
import { inspectPackage } from './package';
import { powerPointRules } from './ppt';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import { wordRules } from './word';
import { excelRules } from './xls';
import { ZipArchive } from './zip';

const f = (format: string, mimeType: string): Format => ({ format, mimeType });

/** The families a compound file can hold, by the stream or storage that marks each. */
type Family = 'word' | 'excel' | 'powerpoint' | 'visio' | 'outlook' | 'publisher' | 'project' | 'encrypted';

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
 * Inspects a compound file: the rules every compound file gets, over its whole tree, and then those of the format
 * its root names.
 */
export async function inspectCompound(ctx: RunContext, cfb: CompoundFile): Promise<void> {
  const found = families(cfb);
  // Word 2007 writes a Word document beside an encrypted package, for readers that cannot decrypt it.
  const main = found.filter(x => !(x === 'word' && found.includes('encrypted')));
  if (main.length > 1) throw new StopError(C.Package, D.AmbiguousFormat, `The compound file holds ${main.join(' and ')} streams`, { reason: `${main.join(' and ')} streams` });
  if (main.length === 0) throw new StopError(C.Package, D.Unrecognized, 'The compound file holds no Office document stream', { reason: 'a compound file of another kind' });
  const family = main[0];
  ctx.format = formatOf(family, cfb);
  await compoundRules(ctx, cfb, '', false);
  const stream = (name: string) => {
    const e = cfb.child(cfb.root, name);
    return e?.kind === 'stream' ? e : undefined;
  };
  // Beside an encrypted package, Word 2007 writes a WordDocument stream for older readers, of zeros in the files at
  // hand. Those readers read it if it holds a document, so it gets the Word rules when it does.
  const word = stream('WordDocument');
  if (word) {
    const r = await wordRules(ctx, cfb, word, family === 'encrypted');
    if (r?.template && family === 'word') ctx.format = f('dot', 'application/msword');
  }
  if (family === 'excel') {
    // Excel 97 to 2003 reads Workbook, and Excel 5.0 and 95 read Book. A file can hold both.
    let template = false;
    let addIn = false;
    for (const name of ['Workbook', 'Book']) {
      const s = stream(name);
      if (!s) continue;
      const r = await excelRules(ctx, cfb, s, unlockWorkbook);
      template ||= r.template;
      addIn ||= r.addIn;
    }
    if (addIn) ctx.format = f('xla', 'application/vnd.ms-excel');
    else if (template) ctx.format = f('xlt', 'application/vnd.ms-excel');
    // Excel keeps the data of a workbook's ActiveX controls in one Ctls stream.
    const controls = stream('Ctls');
    if (controls && controls.size > 0) ctx.report(C.Control, D.ActiveX, controls.name, { stream: controls.name });
  }
  if (family === 'outlook') await messageRules(ctx, cfb, cfb.root);
  const ppt = stream('PowerPoint Document');
  if (family === 'powerpoint' && ppt) await powerPointRules(ctx, cfb, ppt);
  const info = stream('EncryptionInfo');
  const pkg = stream('EncryptedPackage');
  if (family === 'encrypted' && info && pkg) {
    // The package goes through the OOXML rules once a password opens it, and names the format.
    const { source } = await openPackage(ctx, cfb, info, pkg);
    const zip = await ZipArchive.open(source, ctx.settings, ctx.settings.deadline);
    for (const n of zip.notes) ctx.report(C.Container, n.detail, pkg.name, { bytes: n.bytes });
    await inspectPackage(ctx.withZip(zip));
  }
}

function formatOf(family: Family, cfb: CompoundFile): Format | undefined {
  switch (family) {
    case 'word':
      return f('doc', 'application/msword');
    case 'excel':
      return f('xls', 'application/vnd.ms-excel');
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
