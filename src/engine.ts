import { CFB_SIGNATURE, CompoundFile } from './cfb';
import { RunContext, type RunSettings } from './context';
import { checkTime, StopError } from './errors';
import { FindingFactory } from './findings';
import type { Format } from './formats';
import { bufferSource } from './io';
import { inspectCompound } from './ole';
import { inspectPackage } from './package';
import { inspectRtf } from './rtf';
import { scoreFindings } from './score';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D, type DefuseFinding, type OfficeInspection, type OfficeLimits, type OfficeOptions, OfficeRisk } from './types';
import { ZipArchive } from './zip';

const DEFAULT_LIMITS = {
  entries: 10_000,
  decompressedBytes: 512 * 1024 * 1024,
  totalDecompressedBytes: 2 * 1024 * 1024 * 1024,
  xmlDepth: 256,
  xmlTokenSize: 16 * 1024 * 1024,
  xmlEntityDepth: 16,
  xmlEntityExpansions: 100_000,
  xmlEntitySize: 16 * 1024 * 1024,
  rtfDepth: 256,
};

type Limits = typeof DEFAULT_LIMITS & Pick<OfficeLimits, 'fileSize' | 'timeMs'>;

function resolveLimits(given: OfficeLimits = {}): Limits {
  return {
    entries: given.entries ?? DEFAULT_LIMITS.entries,
    decompressedBytes: given.decompressedBytes ?? DEFAULT_LIMITS.decompressedBytes,
    totalDecompressedBytes: given.totalDecompressedBytes ?? DEFAULT_LIMITS.totalDecompressedBytes,
    xmlDepth: given.xmlDepth ?? DEFAULT_LIMITS.xmlDepth,
    xmlTokenSize: given.xmlTokenSize ?? DEFAULT_LIMITS.xmlTokenSize,
    xmlEntityDepth: given.xmlEntityDepth ?? DEFAULT_LIMITS.xmlEntityDepth,
    xmlEntityExpansions: given.xmlEntityExpansions ?? DEFAULT_LIMITS.xmlEntityExpansions,
    xmlEntitySize: given.xmlEntitySize ?? DEFAULT_LIMITS.xmlEntitySize,
    rtfDepth: given.rtfDepth ?? DEFAULT_LIMITS.rtfDepth,
    fileSize: given.fileSize,
    timeMs: given.timeMs,
  };
}

/**
 * `incomplete` marks a run that stopped before it read the whole package. It is rejected without a score, whatever
 * the actions say, because an override can make the finding that stopped it look harmless.
 */
function finish(findings: DefuseFinding[], options: OfficeOptions, incomplete: boolean, format?: Format): OfficeInspection {
  const named = { format: format?.format ?? '', mimeType: format?.mimeType ?? '' };
  if (incomplete) return { status: 'rejected', score: null, risk: OfficeRisk.Unknown, findings, ...named };
  const rejected = findings.some(f => f.action === 'reject');
  const strippable = findings.some(f => f.action === 'strip');
  const { score, risk } = scoreFindings(findings, options);
  return { status: rejected ? 'rejected' : strippable ? 'strippable' : 'clean', score, risk, findings, ...named };
}

/**
 * Wraps the source for one run. A read that returns more or fewer bytes than the source's size allows fails with
 * code EIO, since the reader would take the shifted bytes for the file. Any error the source throws is kept, and the
 * engine throws it instead of a verdict on the file.
 */
function guarded(source: ByteSource): ByteSource & { failure?: unknown } {
  let size = 0;
  const g: ByteSource & { failure?: unknown } = {
    size: async () => {
      try {
        size = await source.size();
        return size;
      } catch (e) {
        g.failure ??= e;
        throw e;
      }
    },
    read: async (offset, length) => {
      try {
        const b = await source.read(offset, length);
        const want = Math.max(0, Math.min(length, size - offset));
        if (b.length !== want) throw Object.assign(new Error(`The source returned ${b.length} bytes at offset ${offset} where ${want} were expected`), { code: 'EIO' });
        return b;
      } catch (e) {
        g.failure ??= e;
        throw e;
      }
    },
  };
  return g;
}

type Container = { kind: 'zip'; zip: ZipArchive } | { kind: 'cfb'; cfb: CompoundFile } | { kind: 'rtf' };

const ambiguous = (reason: string) => new StopError(C.Package, D.AmbiguousFormat, `The file is ${reason}`, { reason });

/**
 * Reads the file's container. A file that starts with the compound file signature is read as one, and one that
 * starts with {\rt is RTF, which is how Word tells them. The ZIP reader runs too, from the end: a file both read is
 * ambiguous, since Office goes by the first bytes and a ZIP reader by the last. A compound file that does not read is
 * a ZIP archive with data in front, if the ZIP reader read one. Word reads any file that starts with {\rt as RTF, so
 * that does not fail.
 */
async function openContainer(src: ByteSource, size: number, limits: Limits, settings: RunSettings): Promise<Container> {
  const head = await src.read(0, Math.min(8, size));
  const cfbSignature = head.length === 8 && head.every((b, i) => b === CFB_SIGNATURE[i]);
  const rtfSignature = Buffer.from(head.subarray(0, 4)).toString('latin1') === '{\\rt';
  const probe =
    cfbSignature || rtfSignature
      ? await ZipArchive.open(src, limits, settings.deadline).catch((e: unknown) => {
          if (e instanceof StopError) return undefined;
          throw e;
        })
      : undefined;
  if (rtfSignature) {
    if (probe) throw ambiguous('both RTF and a ZIP archive');
    return { kind: 'rtf' };
  }
  if (cfbSignature) {
    try {
      const cfb = await CompoundFile.open(src, settings, settings.deadline);
      if (probe) throw ambiguous('both a compound file and a ZIP archive');
      return { kind: 'cfb', cfb };
    } catch (e) {
      if (!probe || !(e instanceof StopError) || e.category !== C.Container) throw e;
    }
  }
  return { kind: 'zip', zip: probe ?? (await ZipArchive.open(src, limits, settings.deadline)) };
}

/** Inspects an Office or OpenDocument file without changing it. The source is closed when the run ends. */
export async function inspectOfficeSource(source: ByteSource, options: OfficeOptions = {}): Promise<OfficeInspection> {
  const limits = resolveLimits(options.limits);
  const deadline = limits.timeMs !== undefined ? Date.now() + limits.timeMs : undefined;
  const factory = new FindingFactory(options.actionOverrides);
  const findings: DefuseFinding[] = [];
  const src = guarded(source);
  const settings: RunSettings = {
    entries: limits.entries,
    decompressedBytes: limits.decompressedBytes,
    totalDecompressedBytes: limits.totalDecompressedBytes,
    rtfDepth: limits.rtfDepth,
    deadline,
    password: options.password,
  };
  const ctx = new RunContext(
    factory,
    {
      depth: limits.xmlDepth,
      tokenSize: limits.xmlTokenSize,
      entityDepth: limits.xmlEntityDepth,
      entityExpansions: limits.xmlEntityExpansions,
      entitySize: limits.xmlEntitySize,
    },
    settings,
  );
  try {
    const size = await src.size();
    if (limits.fileSize !== undefined && size > limits.fileSize) throw new StopError(C.Limit, D.FileSize, `${size} bytes`, { size, limit: limits.fileSize });
    checkTime(deadline);
    const container = await openContainer(src, size, limits, settings);
    if (container.kind === 'rtf') await inspectRtf(ctx, src, size);
    else if (container.kind === 'cfb') {
      for (const n of container.cfb.notes) findings.push(factory.make(C.Container, n.detail, undefined, { bytes: n.bytes }));
      await inspectCompound(ctx, container.cfb);
    } else {
      const zip = container.zip;
      for (const n of zip.notes) findings.push(factory.make(C.Container, n.detail, undefined, { bytes: n.bytes }));
      await inspectPackage(ctx.withZip(zip));
    }
    findings.push(...ctx.findings);
    return finish(findings, options, false, ctx.format);
  } catch (e) {
    if (src.failure !== undefined) throw src.failure;
    findings.push(...ctx.findings);
    if (e instanceof StopError) {
      findings.push(factory.make(e.category, e.detail, e.location ?? (typeof e.data?.name === 'string' ? e.data.name : undefined), { ...e.data, reason: e.data?.reason ?? e.message }));
    } else if ((e as { code?: unknown })?.code === 'DEFUSE_LIMIT') {
      // The family's time limit, from this package or any other.
      findings.push(factory.make(C.Limit, D.Time));
    } else {
      // Anything unexpected rejects the file rather than escaping as an exception.
      findings.push(factory.make(C.Package, D.Unreadable, undefined, { reason: String((e as Error)?.message ?? e).slice(0, 200) }));
    }
    return finish(findings, options, true, ctx.format);
  } finally {
    try {
      await source.close?.();
    } catch {
      // The verdict stands whether or not the source closes.
    }
  }
}

/** Inspects Office or OpenDocument bytes. */
export async function inspectOffice(bytes: Uint8Array, options: OfficeOptions = {}): Promise<OfficeInspection> {
  return inspectOfficeSource(bufferSource(bytes), options);
}
