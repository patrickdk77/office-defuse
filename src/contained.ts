import type { Expected, RunContext, RunState } from './context';
import { stopsTree } from './errors';
import { SpillSink } from './io';
import { sniffType } from './sniff';
import { OfficeCategory as C, type ContainedFile, type ContainedFileResult, OfficeDetail as D, type DefuseFinding, type OfficeOptions, type PluginContext } from './types';

/** The key under which office-defuse hands its run's options to a plugin. No other package can name it. */
export const RUN_OPTIONS = Symbol('office-defuse run options');

/** A plugin context as office-defuse makes it. */
export type RunPluginContext = PluginContext & { [RUN_OPTIONS]?: OfficeOptions };

/** What the plugins decided about a file inside the file. Undefined from `decideContained` means no plugin took it. */
export interface ContainedDecision {
  keep: boolean;
  /** The bytes a plugin kept in place of the file's own. */
  replacement?: SpillSink;
}

/** A file found inside the file, as the format rules hand it over. */
export interface InnerFile {
  name?: string;
  /** The MIME type the container gives the file, such as a part's content type. */
  declaredType?: string;
  location: string;
  /** The file's bytes, in order. */
  chunks(): AsyncIterable<Uint8Array>;
}

/**
 * Throws on an error a plugin threw that stops the whole tree. An I/O error fails this run as one of its own would,
 * so its cause goes up.
 */
function passOn(run: RunState, e: unknown): void {
  const code = stopsTree(e);
  if (code === 'DEFUSE_IO') {
    const cause = (e as { cause?: unknown }).cause ?? e;
    run.temp.ioError ??= cause;
    throw cause;
  }
  if (code === 'DEFUSE_LIMIT') throw e;
}

/**
 * Adds the findings a plugin reported from inside a file, labeled with the file's name and place. A reject inside the
 * file means the file goes, not the one that holds it, so it is recorded here as strip.
 */
export function pushContained(ctx: RunContext, found: DefuseFinding[] | undefined, location: string, name: string | undefined): void {
  const label = name || 'unnamed file';
  ctx.push(
    (found ?? []).map(f => ({
      ...f,
      action: f.action === 'reject' ? 'strip' : f.action,
      location: f.location ? `${location} > ${f.location}` : location,
      data: f.data && { ...f.data },
      attachment: f.attachment ? `${label} > ${f.attachment}` : label,
    })),
  );
}

/**
 * Shows a file found inside the file to the plugins, as pdf-defuse shows an attached file. The first plugin that
 * accepts it decides: 'passed' and 'scrubbed' keep it, with the bytes the plugin wrote when it wrote any, and
 * 'removed', a failure or a scrub that wrote nothing remove it. Its findings come into this run's report. When no
 * plugin accepts the file, this returns undefined and the rules the file gets without plugins decide.
 *
 * In the run that checks an output, the plugins are not asked again. A file whose MD5 matches what a plugin kept in
 * the upload's run is taken as kept, with the findings the plugin reported for its output.
 */
export async function decideContained(ctx: RunContext, inner: InnerFile): Promise<ContainedDecision | undefined> {
  const run = ctx.run;
  const plugins = run.options.filePlugins ?? [];
  if (!plugins.length && !run.verify?.size) return undefined;
  const decoded = new SpillSink(run.temp, run.threshold);
  try {
    for await (const c of inner.chunks()) await decoded.write(c);
    await decoded.close();
    const hash = decoded.digest();
    const name = inner.name;
    const expected = run.verify?.get(hash);
    if (expected) {
      ctx.report(C.Embedded, expected.result === 'passed' ? D.PluginPassed : D.PluginScrubbed, inner.location, { plugin: expected.plugin, name: name ?? '' });
      pushContained(ctx, expected.findings, inner.location, name);
      return { keep: true };
    }
    if (!plugins.length) return undefined;
    const source = decoded.source();
    try {
      const head = await source.read(0, Math.min(1024, decoded.length));
      const file: ContainedFile = { name, declaredType: inner.declaredType, sniffedType: sniffType(head), size: decoded.length, location: inner.location, depth: run.depth + 1, source };
      const context: RunPluginContext = { depth: file.depth, deadline: run.deadline, [RUN_OPTIONS]: run.options };
      for (const plugin of plugins) {
        try {
          if (!(await plugin.accepts(file))) continue;
        } catch (e) {
          passOn(run, e);
          continue;
        }
        // Copies a plugin keeps wait in memory until the rewrite, so all of them together stay under one threshold.
        const out = new SpillSink(run.temp, Math.max(0, run.threshold - run.retained));
        const data = { plugin: plugin.name, name: name ?? '' };
        try {
          const answer = await plugin.process(file, out, context);
          const r: Partial<ContainedFileResult> = typeof answer === 'string' ? { result: answer } : (answer ?? {});
          await out.close();
          let keep = false;
          if (r.result === 'passed') {
            ctx.report(C.Embedded, D.PluginPassed, inner.location, data);
            // Bytes the plugin wrote are the ones kept, such as a document rewritten with nothing removed. A container
            // the rewrite cannot change keeps the file as it was, which the plugin passed too.
            const passed: Expected = { result: 'passed', plugin: plugin.name, findings: r.outputFindings };
            run.kept.set(hash, passed);
            if (out.length > 0) run.kept.set(out.digest(), passed);
            keep = true;
          } else if (r.result === 'scrubbed' && out.length > 0) {
            ctx.report(C.Embedded, D.PluginScrubbed, inner.location, data);
            run.kept.set(out.digest(), { result: 'scrubbed', plugin: plugin.name, findings: r.outputFindings });
            keep = true;
          } else ctx.removal(C.Embedded, r.result === 'removed' ? D.PluginRemoved : D.PluginFailed, inner.location, data);
          pushContained(ctx, r.findings, inner.location, name);
          const replacement = keep && out.length > 0 ? out : undefined;
          if (replacement) {
            if (!replacement.spilled) run.retained += replacement.length;
            run.held.push(replacement);
          } else await out.dispose();
          return { keep, replacement };
        } catch (e) {
          await out.dispose();
          passOn(run, e);
          ctx.removal(C.Embedded, D.PluginFailed, inner.location, data);
          return { keep: false };
        }
      }
      return undefined;
    } finally {
      await source.close?.();
    }
  } finally {
    await decoded.dispose();
  }
}

/**
 * The rule a file gets for being there, such as EMBEDDED/OLE_OBJECT, applies when no plugin took the file. A plugin
 * that kept the file turns those findings to info, since the content stays.
 */
export function presence(decision: ContainedDecision | undefined): 'info' | undefined {
  return decision?.keep ? 'info' : undefined;
}

/** Bytes in a spill sink the run holds until it ends, as the rewrite takes a plugin's output. */
export async function spillOf(run: RunState, bytes: Uint8Array): Promise<SpillSink> {
  const sink = new SpillSink(run.temp, run.threshold);
  run.held.push(sink);
  await sink.write(bytes);
  await sink.close();
  return sink;
}
