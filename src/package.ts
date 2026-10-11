import type { ZipContext } from './context';
import { StopError } from './errors';
import { inspectOdf } from './odf';
import { inspectOoxml } from './ooxml';
import { OfficeCategory as C, OfficeDetail as D } from './types';

/** Inspects an OOXML or OpenDocument package, which its markers name. */
export async function inspectPackage(ctx: ZipContext): Promise<void> {
  const zip = ctx.zip;
  const contentTypes = zip.get('[Content_Types].xml');
  const odf = zip.entries.some(e => e.name === 'mimetype') || zip.get('META-INF/manifest.xml') !== undefined;
  if (contentTypes && odf)
    throw new StopError(C.Package, D.AmbiguousFormat, 'The file has both [Content_Types].xml and an OpenDocument mimetype or manifest', { reason: 'OOXML and OpenDocument markers' });
  if (contentTypes) await inspectOoxml(ctx, contentTypes);
  else if (odf) await inspectOdf(ctx);
  else throw new StopError(C.Package, D.Unrecognized, 'The archive has neither [Content_Types].xml nor an OpenDocument mimetype or manifest', { reason: 'a plain ZIP archive' });
}
