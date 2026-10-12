import type { RunContext } from './context';
import { type Container, endRun, guarded, type Limits, newRun, openContainer, runContext } from './engine';
import { checkTime, isIoError, StopError, stopsTree } from './errors';
import { nameFlat } from './flat';
import { type Format, RTF_FORMAT } from './formats';
import { bufferSource } from './io';
import { compoundFormat, mainFamily, openEncryptedPackage } from './ole';
import { namePackage } from './package';
import { type ByteSource, OfficeCategory as C, type DetectOptions, type FormatDetection } from './types';

/**
 * The format of a file whose container `openContainer` read, named by the steps the inspection names it by, without
 * the rules: the root of a compound file, and for Word and Excel whether it is a template or an add-in; the markers,
 * [Content_Types].xml and package relationships of a package; the root of a flat file; or RTF. An encrypted package
 * names its format when the passwords the run tries, which without a password is VelvetSweatshop alone, open it.
 */
async function nameContainer(ctx: RunContext, container: Container, src: ByteSource, size: number, limits: Limits): Promise<Format> {
  if (container.kind === 'rtf') return RTF_FORMAT;
  if (container.kind === 'xml') return nameFlat(ctx, src, size, limits);
  if (container.kind === 'zip') return namePackage(ctx.withZip(container.zip));
  const cfb = container.cfb;
  const family = mainFamily(cfb);
  const format = await compoundFormat(cfb, family, ctx.settings.deadline);
  if (format) return format;
  try {
    return await namePackage(ctx.withZip(await openEncryptedPackage(ctx, cfb)));
  } catch (e) {
    // An encrypted package that no password at hand opens is still office-defuse's. Only the password says which
    // format is inside, and the inspection decrypts it with the run's password or rejects it as encrypted.
    if (family === 'encrypted' && e instanceof StopError && e.category === C.Encrypted) return ENCRYPTED_PACKAGE;
    throw e;
  }
}

/** An encrypted Office Open XML package whose format no password at hand reveals: a compound file, with no extension of its own. */
const ENCRYPTED_PACKAGE: Format = { format: '', mimeType: 'application/x-ole-storage' };

/**
 * The format office-defuse reads `source` as, when it can name it from the content: the format and MIME type an
 * inspection reports, read the way the inspection reads them, without the rules, the plugins or the rewrite. A ZIP
 * archive, compound file or XML file whose format it cannot name, such as a plain archive, an installer or a package
 * of another kind, is not claimed. A file it names but cannot read past that point is claimed, and an inspection then
 * rejects it. So is an encrypted Office Open XML package no password at hand opens, as application/x-ole-storage with
 * no extension, since only the password says which format is inside. The offset is that of the ZIP data in an archive with data in front, and 0 otherwise. The name and the
 * declared type play no part. The source stays open.
 */
export async function detectOfficeSource(source: ByteSource, options: DetectOptions = {}): Promise<FormatDetection | undefined> {
  const run = newRun({ tempDir: options.tempDir, memoryThreshold: options.memoryThreshold }, 0, options.deadline);
  const src = guarded(source);
  try {
    const { limits, settings, ctx } = runContext(run);
    const size = await src.size();
    checkTime(run.deadline);
    const container = await openContainer(src, size, limits, settings);
    const format = await nameContainer(ctx, container, src, size, limits);
    if (run.temp.ioError !== undefined) throw run.temp.ioError;
    return { mimeType: format.mimeType, extension: format.format, offset: container.kind === 'zip' ? container.zip.start : 0 };
  } catch (e) {
    if (src.failure !== undefined) throw src.failure;
    if (run.temp.ioError !== undefined) throw run.temp.ioError;
    if (stopsTree(e) !== undefined || isIoError(e)) throw e;
    // Any other failure, as in an inspection, means office-defuse cannot name the file's format.
    return undefined;
  } finally {
    await endRun(run);
  }
}

/** `detectOfficeSource` over bytes in memory. */
export function detectOffice(bytes: Uint8Array, options: DetectOptions = {}): Promise<FormatDetection | undefined> {
  return detectOfficeSource(bufferSource(bytes), options);
}
