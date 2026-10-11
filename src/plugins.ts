import { RUN_OPTIONS, type RunPluginContext } from './contained';
import { disarmAtDepth } from './engine';
import { isFlatHead } from './flat';
import { weightOf } from './score';
import type { ByteSource, ContainedFilePlugin, OfficeOptions } from './types';
import { ZipArchive } from './zip';

/** What the ZIP reader needs to look for the markers of an Office package, with the default limits. */
const PROBE = { entries: 10_000, decompressedBytes: 512 * 1024 * 1024, totalDecompressedBytes: 2 * 1024 * 1024 * 1024 };

/**
 * Whether the content is a file office-defuse reads: a compound file, RTF, a ZIP archive with the marker of an Office
 * Open XML or OpenDocument package, or XML whose root is a flat OpenDocument document or a Flat OPC package. A ZIP
 * archive without a marker, or one the ZIP reader refuses, is not.
 */
async function isOffice(source: ByteSource, sniffed: string | undefined): Promise<boolean> {
  if (sniffed === 'application/x-ole-storage' || sniffed === 'application/rtf') return true;
  if (sniffed === 'application/xml') {
    const head = await source.read(0, Math.min(16384, await source.size()));
    return isFlatHead(Buffer.from(head).toString('utf8'));
  }
  if (sniffed !== 'application/zip') return false;
  try {
    const zip = await ZipArchive.open(source, PROBE);
    return zip.get('[Content_Types].xml') !== undefined || zip.entries.some(e => e.name === 'mimetype') || zip.get('META-INF/manifest.xml') !== undefined;
  } catch {
    return false;
  }
}

/**
 * Defuses Office and OpenDocument files inside other files, with `options` when given. Without them, it takes the
 * options of the office-defuse run that found the file, or none when another package found it. `limits.nestingDepth`
 * applies to the depth in the context. The file's findings go into the containing file's report, and its output's
 * findings into the output's report, each with its `weight`.
 */
export function officePlugin(options?: OfficeOptions): ContainedFilePlugin {
  return {
    kind: 'file',
    name: 'office',
    // It can only read what it can parse, so the content decides, whatever the name says.
    accepts: f => isOffice(f.source, f.sniffedType),
    async process(file, sink, context) {
      const run = options ?? (context as RunPluginContext)[RUN_OPTIONS] ?? {};
      const r = await disarmAtDepth(file.source, sink, run, context.depth, context.deadline);
      // A container in another package scores these by the weights this run used.
      for (const f of [...r.before.findings, ...(r.after?.findings ?? [])]) f.weight = weightOf(f, run);
      const findings = r.before.findings;
      // The sink holds the rewritten file, also when nothing was removed.
      if (r.status === 'clean') return { result: 'passed', findings, outputFindings: r.after?.findings };
      if (r.status === 'defused') return { result: 'scrubbed', findings, outputFindings: r.after?.findings };
      return { result: 'removed', findings };
    },
  };
}
