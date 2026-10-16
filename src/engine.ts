import { CFB_SIGNATURE, CompoundFile } from './cfb';
import { type Output, RunContext, type RunSettings, type RunState } from './context';
import { checkTime, isIoError, NestedIoError, NestingLimitError, StopError, stopsTree } from './errors';
import { FindingFactory } from './findings';
import { inspectFlat, startsAsXml, writeFlatOdf, writeFlatOpc } from './flat';
import type { Format } from './formats';
import { bufferSink, bufferSource, copyToSink, SpillSink, TempDir } from './io';
import { inspectCompound } from './ole';
import { inspectPackage } from './package';
import { writeCompound } from './rewrite-cfb';
import { writeRtf } from './rewrite-rtf';
import { writePackage } from './rewrite-zip';
import { inspectRtf } from './rtf';
import { scoreFindings } from './score';
import {
  type ByteSink,
  type ByteSource,
  OfficeCategory as C,
  OfficeDetail as D,
  type DefuseFinding,
  type OfficeDisarmResult,
  type OfficeInspection,
  type OfficeLimits,
  type OfficeOptions,
  OfficeRisk,
} from './types';
import type { XmlLimits } from './xml';
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

export type Limits = typeof DEFAULT_LIMITS & Pick<OfficeLimits, 'fileSize' | 'timeMs' | 'nestingDepth'>;

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
    nestingDepth: given.nestingDepth,
  };
}

const xmlLimitsOf = (limits: Limits): XmlLimits => ({
  depth: limits.xmlDepth,
  tokenSize: limits.xmlTokenSize,
  entityDepth: limits.xmlEntityDepth,
  entityExpansions: limits.xmlEntityExpansions,
  entitySize: limits.xmlEntitySize,
});

const PLUGIN_KEPT = new Set<string>([D.PluginPassed, D.PluginScrubbed]);

/**
 * `incomplete` marks a run that stopped before it read the whole package. It is rejected without a score, whatever
 * the actions say, because an override can make the finding that stopped it look harmless. In the check of an output,
 * `isOutput`, the files plugins kept count in the score; in an upload they do not, and a file a plugin cleaned makes
 * the upload strippable, since the output must be written.
 */
function finish(findings: DefuseFinding[], options: OfficeOptions, incomplete: boolean, format?: Format, isOutput = false, superseded?: WeakSet<DefuseFinding>): OfficeInspection {
  const named = { format: format?.format ?? '', mimeType: format?.mimeType ?? '' };
  if (incomplete) return { status: 'rejected', score: null, risk: OfficeRisk.Unknown, findings, ...named };
  const rejected = findings.some(f => f.action === 'reject');
  const strippable = findings.some(f => f.action === 'strip' || (!isOutput && f.detail === D.PluginScrubbed));
  const scored = findings.filter(f => !superseded?.has(f) && (isOutput || !PLUGIN_KEPT.has(f.detail)));
  const { score, risk } = scoreFindings(scored, options);
  return { status: rejected ? 'rejected' : strippable ? 'strippable' : 'clean', score, risk, findings, ...named };
}

/**
 * Wraps the source for one run. A read that returns more or fewer bytes than the source's size allows fails with
 * code EIO, since the reader would take the shifted bytes for the file. Any error the source throws is kept, and the
 * engine throws it instead of a verdict on the file.
 */
export function guarded(source: ByteSource): ByteSource & { failure?: unknown } {
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

export type Container = { kind: 'zip'; zip: ZipArchive } | { kind: 'cfb'; cfb: CompoundFile } | { kind: 'rtf' } | { kind: 'xml' };

const ambiguous = (reason: string) => new StopError(C.Package, D.AmbiguousFormat, `The file is ${reason}`, { reason });

/**
 * Reads the file's container. A file that starts with the compound file signature is read as one, and one that
 * starts with {\rt is RTF, which is how Word tells them. The ZIP reader runs too, from the end: a file both read is
 * ambiguous, since Office goes by the first bytes and a ZIP reader by the last. A compound file that does not read is
 * a ZIP archive with data in front, if the ZIP reader read one. Word reads any file that starts with {\rt as RTF, so
 * that does not fail. A file that starts as XML is a flat OpenDocument file or a Flat OPC package, and one that is a
 * ZIP archive as well is ambiguous.
 */
export async function openContainer(src: ByteSource, size: number, limits: Limits, settings: RunSettings): Promise<Container> {
  const head = await src.read(0, Math.min(1024, size));
  const cfbSignature = head.length >= 8 && CFB_SIGNATURE.every((b, i) => head[i] === b);
  const rtfSignature = Buffer.from(head.subarray(0, 4)).toString('latin1') === '{\\rt';
  const xml = !cfbSignature && !rtfSignature && startsAsXml(head);
  const probe =
    cfbSignature || rtfSignature || xml
      ? await ZipArchive.open(src, limits, settings.deadline).catch((e: unknown) => {
          if (e instanceof StopError) return undefined;
          throw e;
        })
      : undefined;
  if (rtfSignature) {
    if (probe) throw ambiguous('both RTF and a ZIP archive');
    return { kind: 'rtf' };
  }
  if (xml) {
    if (probe) throw ambiguous('both XML and a ZIP archive');
    return { kind: 'xml' };
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

/** A run of the engine over one file at `depth`, with the deadline the whole upload shares. */
export function newRun(options: OfficeOptions, depth: number, deadline: number | undefined, verify?: RunState['verify']): RunState {
  return {
    options,
    depth,
    deadline,
    temp: new TempDir(options.tempDir),
    threshold: options.memoryThreshold ?? 8 * 1024 * 1024,
    retained: 0,
    verify,
    kept: new Map(),
    held: [],
    superseded: new WeakSet(),
    sources: [],
  };
}

/** Releases what a run holds: the copies plugins returned and the temporary directory. */
export async function endRun(run: RunState): Promise<void> {
  try {
    for (const s of run.sources.splice(0)) await s.close?.();
    for (const h of run.held.splice(0)) await h.dispose();
  } finally {
    await run.temp.cleanup();
  }
}

/** The limits, settings and context a run reads its file with, from the run's options. */
export function runContext(run: RunState): { limits: Limits; settings: RunSettings; ctx: RunContext } {
  const options = run.options;
  const limits = resolveLimits(options.limits);
  const password = options.password;
  const settings: RunSettings = {
    entries: limits.entries,
    decompressedBytes: limits.decompressedBytes,
    totalDecompressedBytes: limits.totalDecompressedBytes,
    rtfDepth: limits.rtfDepth,
    deadline: run.deadline,
    passwords: typeof password === 'string' ? [password] : (password ?? []),
  };
  return { limits, settings, ctx: new RunContext(new FindingFactory(options.actionOverrides), xmlLimitsOf(limits), settings, run) };
}

/** The finding a limit error from this package or any other gives at depth 0. */
function limitFinding(factory: FindingFactory, e: unknown) {
  return factory.make(C.Limit, (e as { limit?: unknown }).limit === 'nesting' ? D.NestingDepth : D.Time);
}

/**
 * Inspects `source` and leaves what the rewrite needs in `run.output`. It throws an I/O error instead of reporting a
 * verdict on the file, and below depth 0 it throws a limit error from the family, which stops the whole upload.
 */
async function analyze(run: RunState, source: ByteSource): Promise<OfficeInspection> {
  const options = run.options;
  const deadline = run.deadline;
  const { limits, settings, ctx } = runContext(run);
  const factory = ctx.factory;
  const findings: DefuseFinding[] = [];
  const src = guarded(source);
  const isOutput = run.verify !== undefined;
  try {
    if (limits.nestingDepth !== undefined && run.depth > limits.nestingDepth) throw new NestingLimitError('Nesting limit exceeded');
    const size = await src.size();
    if (limits.fileSize !== undefined && size > limits.fileSize) throw new StopError(C.Limit, D.FileSize, `${size} bytes`, { size, limit: limits.fileSize });
    checkTime(deadline);
    const container = await openContainer(src, size, limits, settings);
    if (container.kind === 'rtf') await inspectRtf(ctx, src, size);
    else if (container.kind === 'xml') {
      const output = await inspectFlat(ctx, src, size, limits);
      if (output && ctx.format) run.output = output;
    } else if (container.kind === 'cfb') {
      for (const n of container.cfb.notes) findings.push(factory.make(C.Container, n.detail, undefined, { bytes: n.bytes }));
      await inspectCompound(ctx, container.cfb);
    } else {
      const zip = container.zip;
      for (const n of zip.notes) findings.push(factory.make(C.Container, n.detail, undefined, n.offset === undefined ? { bytes: n.bytes } : { bytes: n.bytes, offset: n.offset }));
      const zctx = ctx.withZip(zip);
      await inspectPackage(zctx);
      if (ctx.format) run.output = { kind: 'zip', zip, plan: zctx.plan, format: ctx.format };
    }
    findings.push(...ctx.findings);
    if (run.temp.ioError !== undefined) throw run.temp.ioError;
    return finish(findings, options, false, ctx.format, isOutput, run.superseded);
  } catch (e) {
    if (src.failure !== undefined) throw src.failure;
    if (run.temp.ioError !== undefined) throw run.temp.ioError;
    run.output = undefined;
    findings.push(...ctx.findings);
    if (e instanceof StopError) {
      findings.push(factory.make(e.category, e.detail, e.location ?? (typeof e.data?.name === 'string' ? e.data.name : undefined), { ...e.data, reason: e.data?.reason ?? e.message }));
    } else if (stopsTree(e) === 'DEFUSE_LIMIT') {
      // The family's time and nesting limits, from this package or any other, stop the whole upload.
      if (run.depth > 0) throw e;
      findings.push(limitFinding(factory, e));
    } else if (isIoError(e)) {
      throw e;
    } else {
      // Anything unexpected rejects the file rather than escaping as an exception.
      findings.push(factory.make(C.Package, D.Unreadable, undefined, { reason: String((e as Error)?.message ?? e).slice(0, 200) }));
    }
    return finish(findings, options, true, ctx.format, isOutput);
  }
}

async function closeSource(source: ByteSource): Promise<void> {
  try {
    await source.close?.();
  } catch {
    // The verdict stands whether or not the source closes.
  }
}

/** Inspects an Office or OpenDocument file without changing it. The source is closed when the run ends. */
export async function inspectOfficeSource(source: ByteSource, options: OfficeOptions = {}): Promise<OfficeInspection> {
  const deadline = options.limits?.timeMs !== undefined ? Date.now() + options.limits.timeMs : undefined;
  const run = newRun(options, 0, deadline);
  try {
    return await analyze(run, source);
  } finally {
    await closeSource(source);
    await endRun(run);
  }
}

/** Inspects Office or OpenDocument bytes. */
export async function inspectOffice(bytes: Uint8Array, options: OfficeOptions = {}): Promise<OfficeInspection> {
  return inspectOfficeSource(bufferSource(bytes), options);
}

/** Copies to the caller's sink. On failure it aborts the sink instead of closing it, so a partial copy does not look whole. */
async function deliver(from: ByteSource, sink: ByteSink): Promise<void> {
  try {
    await copyToSink(from, sink);
  } catch (e) {
    try {
      await sink.abort?.(e);
    } catch {
      /* the copy's error is the one to report */
    }
    throw e;
  }
  await sink.close();
}

/** Writes the rewrite of what the run read. */
async function write(run: RunState, output: Output, sink: ByteSink, source: ByteSource): Promise<void> {
  const resolved = resolveLimits(run.options.limits);
  const limits = xmlLimitsOf(resolved);
  if (output.kind === 'zip') await writePackage(run, output.zip, output.plan, output.format, limits, sink);
  else if (output.kind === 'cfb') await writeCompound(run, output.cfb, output.plan, sink);
  else if (output.kind === 'flatodf') await writeFlatOdf(run, source, output.plan, limits, sink);
  else if (output.kind === 'flatopc') await writeFlatOpc(run, output, limits, resolved, sink);
  else await writeRtf(run, source, output.size, output.cuts, sink);
}

/** Rejects the upload, keeping its findings and, unless a limit stopped the run, its score. */
function refuse(before: OfficeInspection, finding: DefuseFinding, limited: boolean): OfficeDisarmResult {
  before.findings.push(finding);
  before.status = 'rejected';
  if (limited) {
    before.score = null;
    before.risk = OfficeRisk.Unknown;
  }
  return { status: 'rejected', before, removed: [] };
}

/**
 * Writes a defused copy of `source` to `sink`. Nothing reaches the sink unless the output passes inspection. The
 * source is closed when the run ends.
 */
export async function disarmOfficeSource(source: ByteSource, sink: ByteSink, options: OfficeOptions = {}): Promise<OfficeDisarmResult> {
  return disarmAtDepth(source, sink, options, 0);
}

/** Internal entry that carries the depth of a file inside another, and the deadline it shares with the upload. */
export async function disarmAtDepth(
  source: ByteSource,
  sink: ByteSink,
  options: OfficeOptions,
  depth: number,
  deadline = options.limits?.timeMs !== undefined ? Date.now() + options.limits.timeMs : undefined,
): Promise<OfficeDisarmResult> {
  const run = newRun(options, depth, deadline);
  const factory = new FindingFactory(options.actionOverrides);
  let input = source;
  try {
    let before: OfficeInspection;
    try {
      // A caller's source can change between reads, so the inspection and the rewrite read one snapshot. A file
      // inside another already has a source of the engine's, and one over the size limit is rejected unread.
      const fileSize = options.limits?.fileSize;
      if (depth === 0 && (fileSize === undefined || (await source.size()) <= fileSize)) {
        const snapshot = new SpillSink(run.temp, run.threshold);
        run.held.push(snapshot);
        try {
          await copyToSink(source, snapshot, () => checkTime(deadline));
        } finally {
          await snapshot.close();
        }
        input = snapshot.source();
      }
      before = await analyze(run, input);
    } catch (e) {
      if (stopsTree(e) !== 'DEFUSE_LIMIT' || depth > 0) throw e;
      before = finish([limitFinding(factory, e)], options, true);
    }
    const output = run.output;
    if (before.status === 'rejected' || !output) return { status: 'rejected', before, removed: [] };
    const out = new SpillSink(run.temp, run.threshold);
    run.held.push(out);
    try {
      await write(run, output, out, input);
      await out.close();
    } catch (e) {
      await out.close().catch(() => {
        // The write's failure, handled below, is the one that counts.
      });
      if (run.temp.ioError !== undefined) throw run.temp.ioError;
      if (isIoError(e)) throw e;
      if (stopsTree(e) === 'DEFUSE_LIMIT') {
        if (depth > 0) throw e;
        return refuse(before, limitFinding(factory, e), true);
      }
      return refuse(before, factory.make(C.Processing, D.VerificationFailed, undefined, { reasons: `write failed: ${String((e as Error)?.message ?? e).slice(0, 100)}` }), false);
    }
    // The output is checked with the caller's overrides and limits, and the plugins' verdicts in place of the plugins.
    const verifyOptions: OfficeOptions = {
      // No plugin is asked again. An empty list still reads VBA project parts the way the upload's run did.
      scriptPlugins: options.scriptPlugins && [],
      actionOverrides: options.actionOverrides,
      scoreWeights: options.scoreWeights,
      scoreBands: options.scoreBands,
      tempDir: options.tempDir,
      memoryThreshold: options.memoryThreshold,
      limits: { ...options.limits, fileSize: undefined, timeMs: undefined, nestingDepth: undefined },
    };
    const outSource = out.source();
    try {
      const check = newRun(verifyOptions, depth, deadline, run.kept);
      let after: OfficeInspection;
      try {
        after = await analyze(check, outSource);
      } finally {
        await endRun(check);
      }
      const reasons = after.findings.filter(f => f.action !== 'info').map(f => `${f.category}/${f.detail}`);
      // A check that stopped early verified nothing, even when an override made the finding that stopped it info.
      if (after.status === 'rejected' && !reasons.length) reasons.push('the output check stopped early');
      if (after.format !== before.format) reasons.push(`format ${before.format} -> ${after.format}`);
      if (reasons.length) return refuse(before, factory.make(C.Processing, D.VerificationFailed, undefined, { reasons: [...new Set(reasons)].slice(0, 10).join(', ') }), false);
      await deliver(outSource, sink);
      return { status: before.status === 'clean' ? 'clean' : 'defused', before, after, removed: before.findings.filter(f => f.action === 'strip') };
    } finally {
      await closeSource(outSource);
    }
  } catch (e) {
    // An I/O error inside a nested file fails the upload's run too, instead of counting as a failed plugin.
    if (depth > 0 && (isIoError(e) || (e !== undefined && e === run.temp.ioError))) throw new NestedIoError(e);
    throw e;
  } finally {
    if (input !== source) await closeSource(input);
    if (depth === 0) await closeSource(source);
    await endRun(run);
  }
}

/** Defuses Office or OpenDocument bytes. `bytes` holds the output unless the file was rejected. */
export async function disarmOffice(bytes: Uint8Array, options: OfficeOptions = {}): Promise<OfficeDisarmResult & { bytes?: Uint8Array }> {
  const sink = bufferSink();
  const result = await disarmOfficeSource(bufferSource(bytes), sink, options);
  return result.status === 'rejected' ? result : { ...result, bytes: sink.result() };
}
