import { RUN_OPTIONS, type RunPluginContext } from './contained';
import { detectOfficeSource } from './detect';
import { disarmAtDepth } from './engine';
import { weightOf } from './score';
import type { ContainedFilePlugin, OfficeOptions } from './types';

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
    // It takes what detectOffice names: the content decides, whatever the name says.
    accepts: async f => (await detectOfficeSource(f.source, { name: f.name, declaredType: f.declaredType })) !== undefined,
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
