import type { ZipContext } from './context';
import { StopError } from './errors';
import type { Format } from './formats';
import { inspectOdf, nameOdf } from './odf';
import { inspectOoxml, nameOoxml } from './ooxml';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import type { ZipArchive, ZipEntry } from './zip';

/**
 * The package a ZIP archive's markers name: OOXML by [Content_Types].xml, OpenDocument by a mimetype entry or a
 * manifest. Throws for an archive with both or neither.
 */
export function packageKind(zip: ZipArchive): { kind: 'ooxml'; contentTypes: ZipEntry } | { kind: 'odf' } {
  const contentTypes = zip.get('[Content_Types].xml');
  const odf = zip.entries.some(e => e.name === 'mimetype') || zip.get('META-INF/manifest.xml') !== undefined;
  if (contentTypes && odf)
    throw new StopError(C.Package, D.AmbiguousFormat, 'The file has both [Content_Types].xml and an OpenDocument mimetype or manifest', { reason: 'OOXML and OpenDocument markers' });
  if (contentTypes) return { kind: 'ooxml', contentTypes };
  if (odf) return { kind: 'odf' };
  throw new StopError(C.Package, D.Unrecognized, 'The archive has neither [Content_Types].xml nor an OpenDocument mimetype or manifest', { reason: 'a plain ZIP archive' });
}

/** Inspects an OOXML or OpenDocument package, which its markers name. */
export async function inspectPackage(ctx: ZipContext): Promise<void> {
  const k = packageKind(ctx.zip);
  if (k.kind === 'ooxml') await inspectOoxml(ctx, k.contentTypes);
  else await inspectOdf(ctx);
}

/** Names an OOXML or OpenDocument package's format as its inspection does, without the rules. */
export async function namePackage(ctx: ZipContext): Promise<Format> {
  const k = packageKind(ctx.zip);
  return k.kind === 'ooxml' ? nameOoxml(ctx, k.contentTypes) : (await nameOdf(ctx)).format;
}
