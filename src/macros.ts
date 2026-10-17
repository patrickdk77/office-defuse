import { type CfbEntry, type CompoundFile, showName } from './cfb';
import { passOn, spillOf } from './contained';
import { clip, type Expected, type RunContext } from './context';
import { designerProblem } from './forms';
import type { CompoundPlan } from './plan';
import { OfficeCategory as C, type ContainedScript, OfficeDetail as D, type DefuseFinding } from './types';
import { isPurged, purge, readProject, VbaError, type VbaProject } from './vba';

/** What the script plugins decided about a VBA project. */
export interface VbaDecision {
  keep: boolean;
  /** Why the project could not go to the plugins. */
  reason?: string;
  /** The designer storages of the project's user forms, which stay with it and get no object rules of their own. */
  designers?: CfbEntry[];
  /**
   * A plugin passed or removed every module, and the project stays or goes by those answers alone. They are final:
   * the report holds them, and nothing of office-defuse's own about the project.
   */
  answered?: boolean;
}

/** Keys of the run's kept map for a VBA project, which never match the MD5 of a file. */
const keyOf = (project: VbaProject) => `vba:${project.digest}`;

/**
 * Shows each module of the VBA project in `vba` to the script plugins. The first plugin that accepts a module decides
 * it. The project stays when every module passed; `plan` then writes it without its compiled code, so Office compiles
 * the source the plugins read, and the run's kept map remembers the verdict for the output check. A module no plugin
 * takes, or one a plugin does not pass, leaves the project to the VBA_PROJECT finding, as without plugins.
 *
 * In the run that checks an output, the plugins are not asked again: a project whose sources match one the plugins
 * passed keeps their findings, and a compiled code finding if the output still holds p-code.
 */
export async function decideVba(ctx: RunContext, cfb: CompoundFile, vba: CfbEntry, location: string, plan: CompoundPlan): Promise<VbaDecision> {
  const run = ctx.run;
  const plugins = run.options.scriptPlugins ?? [];
  if (!plugins.length && !run.verify?.size) return { keep: false };
  let project: VbaProject;
  try {
    project = await readProject(cfb, vba, ctx.settings.decompressedBytes);
  } catch (e) {
    if (!(e instanceof VbaError)) throw e;
    return { keep: false, reason: clip(e.message) };
  }
  const expected = run.verify?.get(keyOf(project));
  if (expected) {
    ctx.push((expected.findings ?? []).map(f => ({ ...f, data: f.data && { ...f.data } })));
    const designers = await formDesigners(ctx, cfb, project, location);
    if (!designers) return { keep: false };
    if (!(await isPurged(cfb, project))) ctx.removal(C.Macro, D.CompiledCode, location, { storage: location });
    run.macrosKept = true;
    run.macrosAnswered = true;
    return { keep: true, designers, answered: true };
  }
  if (!plugins.length) return { keep: false };
  const start = ctx.findings.length;
  let keep = project.modules.length > 0;
  let answered = keep;
  const used = new Set<string>();
  for (const [i, m] of project.modules.entries()) {
    const at = `${location}/${showName(m.stream.name)}`;
    const script: ContainedScript = {
      text: m.source,
      trigger: `vba-${m.kind}`,
      location: at,
      fieldName: m.name,
      language: 'vba',
      form: 'office',
      ...(i === 0 ? { references: project.references.map(r => ({ ...r })) } : {}),
    };
    let decided = false;
    for (const plugin of plugins) {
      try {
        if (!(await plugin.accepts(script))) continue;
      } catch (e) {
        passOn(run, e);
        continue;
      }
      decided = true;
      used.add(plugin.name);
      const data = { plugin: plugin.name, module: clip(m.name) };
      try {
        const r = await plugin.process(script);
        if (r.result === 'passed') ctx.report(C.Macro, D.PluginPassed, at, data);
        else {
          keep = false;
          if (r.result !== 'removed') answered = false;
          ctx.removal(C.Macro, D.PluginRemoved, at, data);
        }
        pushScript(ctx, r.findings);
      } catch (e) {
        passOn(run, e);
        keep = false;
        answered = false;
        ctx.removal(C.Macro, D.PluginFailed, at, data);
      }
      break;
    }
    if (!decided) {
      keep = false;
      answered = false;
    }
  }
  if (!keep) {
    if (answered) run.macrosAnswered = true;
    return { keep: false, answered };
  }
  const passed: Expected = { result: 'passed', plugin: [...used].join(', '), findings: ctx.findings.slice(start).map(f => ({ ...f, data: f.data && { ...f.data } })) };
  const designers = await formDesigners(ctx, cfb, project, location);
  if (!designers) return { keep: false };
  run.kept.set(keyOf(project), passed);
  if (!(await isPurged(cfb, project))) {
    let streams: ReturnType<typeof purge>;
    try {
      streams = purge(project);
    } catch (e) {
      if (!(e instanceof VbaError)) throw e;
      run.kept.delete(keyOf(project));
      return { keep: false, reason: clip(e.message) };
    }
    for (const [entry, bytes] of streams.replace) plan.replace.set(entry, await spillOf(run, bytes));
    for (const entry of streams.drop) plan.drop.add(entry);
    ctx.removal(C.Macro, D.CompiledCode, location, { storage: location });
  }
  run.macrosKept = true;
  run.macrosAnswered = true;
  return { keep: true, designers, answered: true };
}

/**
 * The designer storages of the project's user forms, beside its VBA storage, named as their modules. Undefined when
 * one may hold a control that is not a Forms 2.0 control, which CONTROL/ACTIVEX reports at the designer.
 */
async function formDesigners(ctx: RunContext, cfb: CompoundFile, project: VbaProject, location: string): Promise<CfbEntry[] | undefined> {
  const holder = project.storage.parent;
  if (!holder) return [];
  const base = location.slice(0, location.length - showName(project.storage.name).length);
  const designers: CfbEntry[] = [];
  for (const m of project.modules) {
    if (m.kind !== 'form') continue;
    const designer = holder.children.find(c => c.kind === 'storage' && c.name.toUpperCase() === m.name.toUpperCase());
    if (!designer) continue;
    const problem = await designerProblem(cfb, designer, ctx.settings.decompressedBytes);
    const at = `${base}${showName(designer.name)}`;
    if (problem && ctx.report(C.Control, D.ActiveX, at, { storage: at, reason: clip(problem) }).action === 'strip') return undefined;
    designers.push(designer);
  }
  return designers;
}

/** Adds a script plugin's findings, which already name the module. A reject removes the module, so it counts as strip. */
function pushScript(ctx: RunContext, found: DefuseFinding[] | undefined): void {
  ctx.push((found ?? []).map(f => ({ ...f, action: f.action === 'reject' ? 'strip' : f.action, data: f.data && { ...f.data } })));
}
