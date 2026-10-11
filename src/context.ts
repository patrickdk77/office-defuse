import type { CompoundFile } from './cfb';
import type { FindingFactory } from './findings';
import type { Format } from './formats';
import type { SpillSink, TempDir } from './io';
import { checkLink, isLocalTarget, mismatchedLabel } from './links';
import { type CompoundPlan, type FlatOpcLayout, PackagePlan, type RtfCut } from './plan';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D, type DefuseFinding, type OfficeFinding, type OfficeFindingAction, type OfficeOptions } from './types';
import { type EntityBudget, showReferences, type XmlHandler, type XmlLimits, type XmlNotes, XmlTokenizer } from './xml';
import type { ZipArchive, ZipEntry } from './zip';

const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Cuts text from the file to a length that fits a finding's data. A reference to an external entity shows as `&name;`. */
export const clip = (raw: string, n = 200) => {
  const s = showReferences(raw);
  return s.length > n ? `${s.slice(0, n - 3)}...` : s;
};

/** The limits that bound a container and what is decoded from it, and the caller's password. */
export interface RunSettings {
  entries: number;
  decompressedBytes: number;
  totalDecompressedBytes: number;
  /** Group nesting in an RTF file, and destinations open at once. */
  rtfDepth: number;
  deadline?: number;
  /** Tried in order on an encrypted file, before the password the format opens with by default. */
  passwords: readonly string[];
}

/** What a plugin returned for a file, which the output check expects to find again by the MD5 of its bytes. */
export interface Expected {
  result: 'passed' | 'scrubbed';
  plugin: string;
  findings?: DefuseFinding[];
}

/** What the rewrite starts from: the container the run read, and what its findings remove. */
export type Output =
  | { kind: 'zip'; zip: ZipArchive; plan: PackagePlan; format: Format }
  | { kind: 'cfb'; cfb: CompoundFile; plan: CompoundPlan }
  | { kind: 'rtf'; size: number; cuts: RtfCut[] }
  | { kind: 'flatodf'; plan: PackagePlan }
  /** A Flat OPC package, unpacked into `zip`, whose OOXML format is `format`. */
  | { kind: 'flatopc'; zip: ZipArchive; plan: PackagePlan; format: Format; layout: FlatOpcLayout };

/** What one run of the engine holds besides its findings: the caller's options, temporary files and plugin results. */
export interface RunState {
  options: OfficeOptions;
  /** The depth of the file in the upload: 0 for the upload itself. */
  depth: number;
  deadline?: number;
  temp: TempDir;
  /** memoryThreshold, in bytes. */
  threshold: number;
  /** Bytes of plugin output held in memory until the rewrite. */
  retained: number;
  /** Set when the run checks an output: what the plugins returned for the upload's files. */
  verify?: Map<string, Expected>;
  /** What the plugins returned in this run, by the MD5 of the bytes kept. */
  kept: Map<string, Expected>;
  /** Plugin output that waits for the rewrite. The run disposes of it when it ends. */
  held: SpillSink[];
  /**
   * Findings whose content a plugin kept, such as EMBEDDED/PACKAGE for a package a plugin passed. They stay in the
   * report as info and leave the score to what the plugin found, as pdf-defuse leaves out a kept attached file.
   */
  superseded: WeakSet<DefuseFinding>;
  /** Sources the rewrite reads from, such as a compound file a plugin returned, which the run closes when it ends. */
  sources: ByteSource[];
  /** The container, once the format rules have read it. */
  output?: Output;
}

/** A finding the rules made, with its action, which the rewrite reads. Past 200 of one kind, `finding` is absent. */
export interface Reported {
  action: OfficeFindingAction;
  finding?: OfficeFinding;
}

/** What every context of one inspection shares, so a package inside the file reports into the same list. */
interface Shared {
  findings: DefuseFinding[];
  perKind: Map<string, number>;
  overflow: Map<string, Record<string, string | number>>;
  /** Entity expansion in every part, which the entity limits bound. */
  budget: EntityBudget;
  format?: Format;
  run: RunState;
  /** Strip findings made so far, counted past the 200 kept of each kind too. */
  strips: number;
}

/** What one inspection holds while the format rules run. */
export class RunContext {
  protected readonly shared: Shared;
  constructor(
    readonly factory: FindingFactory,
    readonly xmlLimits: XmlLimits,
    readonly settings: RunSettings,
    shared: Shared | RunState,
  ) {
    this.shared = 'run' in shared ? shared : { findings: [], perKind: new Map(), overflow: new Map(), budget: { expansions: 0, size: 0 }, run: shared, strips: 0 };
  }

  get findings(): DefuseFinding[] {
    return this.shared.findings;
  }

  get run(): RunState {
    return this.shared.run;
  }

  /** How many strip findings the run has made, so a rule can tell whether anything it read made one. */
  get strips(): number {
    return this.shared.strips;
  }

  /** Set as soon as the rules know it, so a run that stops later still names the format. */
  get format(): Format | undefined {
    return this.shared.format;
  }

  set format(f: Format | undefined) {
    this.shared.format = f;
  }

  /** A context for a ZIP archive the rules read, reporting into the same findings, with a plan of its own. */
  withZip(zip: ZipArchive): ZipContext {
    return new ZipContext(zip, new PackagePlan(), this.factory, this.xmlLimits, this.settings, this.shared);
  }

  /** A context for rules that fill in a plan of their own without an archive, as for a flat XML file. */
  withPlan(): PlanContext {
    return new PlanContext(new PackagePlan(), this.factory, this.xmlLimits, this.settings, this.shared);
  }

  /**
   * Records a finding. Past 200 of one kind, as in pdf-defuse, one more finding counts the rest. `action` replaces the
   * action the overrides give, for content whose fate a plugin decided.
   */
  report(category: C, detail: D, location?: string, data?: Record<string, string | number>, action?: OfficeFindingAction): Reported {
    const act = action ?? this.factory.actionFor(category, detail);
    if (act === 'strip') this.shared.strips++;
    const k = `${category}/${detail}/${act}`;
    const n = (this.shared.perKind.get(k) ?? 0) + 1;
    this.shared.perKind.set(k, n);
    if (n <= 200) {
      const finding = this.factory.make(category, detail, location, data);
      finding.action = act;
      if (action === 'info') this.run.superseded.add(finding);
      this.shared.findings.push(finding);
      return { action: act, finding };
    }
    let counted = this.shared.overflow.get(k);
    if (!counted) {
      // The first overflow fills the description's placeholders.
      counted = { ...data, count: 0 };
      this.shared.overflow.set(k, counted);
      const finding = this.factory.make(category, detail, 'additional occurrences', counted);
      finding.action = act;
      if (action === 'info') this.run.superseded.add(finding);
      this.shared.findings.push(finding);
    }
    counted.count = Number(counted.count) + 1;
    return { action: act };
  }

  /** Reports content that goes whatever the overrides say: its action is strip, or reject when the caller asked for that. */
  removal(category: C, detail: D, location?: string, data?: Record<string, string | number>): Reported {
    const act = this.factory.actionFor(category, detail);
    return this.report(category, detail, location, data, act === 'info' ? 'strip' : act);
  }

  /** Adds findings from another package as they are, such as those a plugin made inside a file. */
  push(findings: DefuseFinding[]): void {
    this.shared.findings.push(...findings);
    this.shared.strips += findings.filter(f => f.action === 'strip').length;
  }

  /** Reports a link, classified by pdf-defuse's checks of its target. */
  link(target: string, location: string, data: Record<string, string> = {}): Reported {
    const v = checkLink(target);
    return this.report(v.category, v.detail, location, { ...data, target: clip(target, 2000), ...(v.host ? { host: clip(v.host) } : {}) });
  }

  /**
   * Reports a safe link whose text or tooltip names another site than the one it opens, as LINK/TEXT_MISMATCH. Returns
   * undefined when no label names a site, or every site a label names is the target's.
   */
  labelMismatch(target: string, labels: ReadonlyArray<string | undefined>, location: string, data: Record<string, string> = {}): Reported | undefined {
    const label = mismatchedLabel(target, labels);
    if (label === undefined) return undefined;
    return this.report(C.Link, D.TextMismatch, location, { ...data, target: clip(target, 2000), label: clip(label) });
  }

  /**
   * Reports a reference that loads from outside the file on its own. It is EXTERNAL/LOCAL_FILE, with `kind` saying
   * what refers, when the target is on this machine, and `detail` otherwise.
   */
  external(detail: D, target: string, location: string, data: Record<string, string>, kind: string): Reported {
    const d = { ...data, target: clip(target, 2000) };
    if (isLocalTarget(target)) return this.report(C.External, D.LocalFile, location, { ...d, kind });
    return this.report(C.External, detail, location, d);
  }

  /**
   * Tokenizes XML text, or Office's VML dialect, handing the events to `handler`. A document type declaration is
   * XML/DOCTYPE without an internal subset and XML/INTERNAL_SUBSET with one. A reference to an external entity in
   * content is XML/EXTERNAL_ENTITY.
   */
  async parseXml(part: string, chunks: AsyncIterable<Uint8Array> | Iterable<Uint8Array>, handler: XmlHandler, vml = false, plan?: PackagePlan, after?: () => Promise<void>): Promise<void> {
    const notes: XmlNotes = {
      doctype: d => {
        const ids = { ...(d.publicId === undefined ? {} : { publicId: clip(d.publicId) }), ...(d.systemId === undefined ? {} : { systemId: clip(d.systemId) }) };
        if (!d.internalSubset) {
          if (this.report(C.Xml, D.Doctype, part, { part, name: clip(d.name), ...ids }).action === 'strip' && plan) plan.part(part).doctype = true;
        } else {
          const counts = { entities: d.entities, externalEntities: d.externalEntities, attributeDefaults: d.attributeDefaults, skipped: d.skipped };
          if (this.report(C.Xml, D.InternalSubset, part, { part, name: clip(d.name), ...ids, ...counts }).action === 'strip' && plan) plan.part(part).subset = true;
        }
      },
      externalEntity: r => {
        const reported = this.report(C.Xml, D.ExternalEntity, part, {
          part,
          entity: clip(r.name),
          target: clip(r.systemId, 2000),
          ...(r.publicId === undefined ? {} : { publicId: clip(r.publicId) }),
        });
        if (!plan) return;
        const p = plan.part(part);
        p.hasExternal = true;
        if (reported.action === 'strip') p.external = true;
      },
    };
    const t = new XmlTokenizer(handler, this.xmlLimits, part, { vml, notes, budget: this.shared.budget });
    // `after` runs between chunks, so a handler can hand what it gathered to something that waits.
    for await (const chunk of chunks) {
      t.write(chunk);
      await after?.();
    }
    t.end();
    await after?.();
  }
}

/** The context of rules that fill in a rewrite plan: a ZIP package's, or a flat XML file's. */
export class PlanContext extends RunContext {
  constructor(
    /** What the rewrite removes, which the rules fill in as they report. */
    readonly plan: PackagePlan,
    factory: FindingFactory,
    xmlLimits: XmlLimits,
    settings: RunSettings,
    shared: Shared,
  ) {
    super(factory, xmlLimits, settings, shared);
  }
}

/** The context of the rules that read a ZIP archive. */
export class ZipContext extends PlanContext {
  constructor(
    readonly zip: ZipArchive,
    plan: PackagePlan,
    factory: FindingFactory,
    xmlLimits: XmlLimits,
    settings: RunSettings,
    shared: Shared,
  ) {
    super(plan, factory, xmlLimits, settings, shared);
  }

  /** Inflates an entry and tokenizes it as XML, or as Office's VML dialect, handing the events to `handler`. */
  parse(entry: ZipEntry, handler: XmlHandler, vml = false): Promise<void> {
    return this.parseXml(entry.name, this.zip.read(entry), handler, vml, this.plan);
  }

  /** Inflates an entry, which checks it, hands each chunk to `visit`, and tells whether its bytes are an OLE file. */
  async drain(entry: ZipEntry, visit?: (chunk: Uint8Array) => void): Promise<boolean> {
    const head: number[] = [];
    for await (const chunk of this.zip.read(entry)) {
      visit?.(chunk);
      for (let i = 0; i < chunk.length && head.length < OLE_SIGNATURE.length; i++) head.push(chunk[i]);
    }
    return head.length === OLE_SIGNATURE.length && head.every((b, i) => b === OLE_SIGNATURE[i]);
  }
}
