import type { FindingFactory } from './findings';
import type { Format } from './formats';
import { checkLink, isLocalTarget } from './links';
import { OfficeCategory as C, OfficeDetail as D, type DefuseFinding } from './types';
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
  /** Tried first on an encrypted file, before the password the format opens with by default. */
  password?: string;
}

/** What every context of one inspection shares, so a package inside the file reports into the same list. */
interface Shared {
  findings: DefuseFinding[];
  perKind: Map<string, number>;
  overflow: Map<string, Record<string, string | number>>;
  /** Entity expansion in every part, which the entity limits bound. */
  budget: EntityBudget;
  format?: Format;
}

/** What one inspection holds while the format rules run. */
export class RunContext {
  protected readonly shared: Shared;
  constructor(
    readonly factory: FindingFactory,
    readonly xmlLimits: XmlLimits,
    readonly settings: RunSettings,
    shared?: Shared,
  ) {
    this.shared = shared ?? { findings: [], perKind: new Map(), overflow: new Map(), budget: { expansions: 0, size: 0 } };
  }

  get findings(): DefuseFinding[] {
    return this.shared.findings;
  }

  /** Set as soon as the rules know it, so a run that stops later still names the format. */
  get format(): Format | undefined {
    return this.shared.format;
  }

  set format(f: Format | undefined) {
    this.shared.format = f;
  }

  /** A context for a ZIP archive the rules read, reporting into the same findings. */
  withZip(zip: ZipArchive): ZipContext {
    return new ZipContext(zip, this.factory, this.xmlLimits, this.settings, this.shared);
  }

  /** Records a finding. Past 200 of one kind, as in pdf-defuse, one more finding counts the rest. */
  report(category: C, detail: D, location?: string, data?: Record<string, string | number>): void {
    const k = `${category}/${detail}`;
    const n = (this.shared.perKind.get(k) ?? 0) + 1;
    this.shared.perKind.set(k, n);
    if (n <= 200) {
      this.shared.findings.push(this.factory.make(category, detail, location, data));
      return;
    }
    let counted = this.shared.overflow.get(k);
    if (!counted) {
      // The first overflow fills the description's placeholders.
      counted = { ...data, count: 0 };
      this.shared.overflow.set(k, counted);
      this.shared.findings.push(this.factory.make(category, detail, 'additional occurrences', counted));
    }
    counted.count = Number(counted.count) + 1;
  }

  /** Reports a link, classified by pdf-defuse's checks of its target. */
  link(target: string, location: string, data: Record<string, string> = {}): void {
    const v = checkLink(target);
    this.report(v.category, v.detail, location, { ...data, target: clip(target, 2000), ...(v.host ? { host: clip(v.host) } : {}) });
  }

  /**
   * Reports a reference that loads from outside the file on its own. It is EXTERNAL/LOCAL_FILE, with `kind` saying
   * what refers, when the target is on this machine, and `detail` otherwise.
   */
  external(detail: D, target: string, location: string, data: Record<string, string>, kind: string): void {
    const d = { ...data, target: clip(target, 2000) };
    if (isLocalTarget(target)) this.report(C.External, D.LocalFile, location, { ...d, kind });
    else this.report(C.External, detail, location, d);
  }

  /**
   * Tokenizes XML text, or Office's VML dialect, handing the events to `handler`. A document type declaration is
   * XML/DOCTYPE without an internal subset and XML/INTERNAL_SUBSET with one. A reference to an external entity in
   * content is XML/EXTERNAL_ENTITY.
   */
  async parseXml(part: string, chunks: AsyncIterable<Uint8Array> | Iterable<Uint8Array>, handler: XmlHandler, vml = false): Promise<void> {
    const notes: XmlNotes = {
      doctype: d => {
        const ids = { ...(d.publicId === undefined ? {} : { publicId: clip(d.publicId) }), ...(d.systemId === undefined ? {} : { systemId: clip(d.systemId) }) };
        if (!d.internalSubset) this.report(C.Xml, D.Doctype, part, { part, name: clip(d.name), ...ids });
        else {
          const counts = { entities: d.entities, externalEntities: d.externalEntities, attributeDefaults: d.attributeDefaults, skipped: d.skipped };
          this.report(C.Xml, D.InternalSubset, part, { part, name: clip(d.name), ...ids, ...counts });
        }
      },
      externalEntity: r =>
        this.report(C.Xml, D.ExternalEntity, part, { part, entity: clip(r.name), target: clip(r.systemId, 2000), ...(r.publicId === undefined ? {} : { publicId: clip(r.publicId) }) }),
    };
    const t = new XmlTokenizer(handler, this.xmlLimits, part, { vml, notes, budget: this.shared.budget });
    for await (const chunk of chunks) t.write(chunk);
    t.end();
  }
}

/** The context of the rules that read a ZIP archive. */
export class ZipContext extends RunContext {
  constructor(
    readonly zip: ZipArchive,
    factory: FindingFactory,
    xmlLimits: XmlLimits,
    settings: RunSettings,
    shared: Shared,
  ) {
    super(factory, xmlLimits, settings, shared);
  }

  /** Inflates an entry and tokenizes it as XML, or as Office's VML dialect, handing the events to `handler`. */
  parse(entry: ZipEntry, handler: XmlHandler, vml = false): Promise<void> {
    return this.parseXml(entry.name, this.zip.read(entry), handler, vml);
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
