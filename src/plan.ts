import type { CfbEntry, CompoundFile } from './cfb';
import type { SpillSink } from './io';
import type { OoxmlGraph } from './ooxml';
import type { OfficeFinding } from './types';
import type { ZipEntry } from './zip';

/** What the rewrite does to one element of an XML part. */
export interface ElementEdit {
  /** The element goes, with everything inside it. */
  remove?: boolean;
  /** Its tags go and its content stays, as with a hyperlink around text. */
  unwrap?: boolean;
  /** Attributes left out, each as its namespace URI, a space and its local name. */
  strip?: Set<string>;
  /** Attribute values replaced, by qualified name. */
  set?: Map<string, string>;
  /** The element's content replaced with this text, as with base64 data a plugin rewrote. */
  content?: string;
}

/** The key `ElementEdit.strip` takes for an attribute. */
export const attributeKey = (ns: string, local: string) => `${ns} ${local}`;

/** A Word field the rewrite removes, by the element numbers of its marks. Its result stays. */
export interface FieldMarks {
  begin: number;
  separate?: number;
  end?: number;
  /** A w:fldSimple, whose tags go and whose result runs stay. */
  simple: boolean;
}

/** What the rewrite does inside one XML part. */
export interface PartPlan {
  elements: Map<number, ElementEdit>;
  fields: FieldMarks[];
  /** The part has an internal subset that goes. Its entity references are expanded and its defaults written out. */
  subset: boolean;
  /** References to external entities in the part go. */
  external: boolean;
  /** A document type declaration without an internal subset goes. */
  doctype: boolean;
  /** The part refers to an external entity. */
  hasExternal: boolean;
}

/** An OpenDocument object element whose fate waits for the plugins' verdict on the entry it names. */
export interface DeferredObject {
  part: string;
  index: number;
  /** The key of the entry it refers to. */
  target: string;
  finding?: OfficeFinding;
  strips: boolean;
}

/** Everything a strip finding names in a ZIP package, for the rewrite. */
export class PackagePlan {
  readonly parts = new Map<string, PartPlan>();
  /** Entries that go with every reference to them, by nameKey. */
  readonly drop = new Set<string>();
  /** OpenDocument folders that go whole, as nameKey prefixes ending in "/". */
  readonly folders = new Set<string>();
  /** Relationship Ids that go, by the name of their relationships part. */
  readonly rels = new Map<string, Set<string>>();
  /** Excel 4.0 macro sheets that become empty worksheets, by nameKey. */
  readonly sheets = new Set<string>();
  /** Entries whose bytes a plugin replaced, by nameKey. */
  readonly replace = new Map<string, SpillSink>();
  /** OpenDocument object references, decided once the plugins have seen the entries they name. */
  readonly objects: DeferredObject[] = [];
  /** What the OOXML rules read of the package's parts and relationships. */
  graph?: OoxmlGraph;
  /** What the OpenDocument rules read of the package's layout. */
  odf?: OdfLayout;

  part(name: string): PartPlan {
    let p = this.parts.get(name);
    if (!p) {
      p = { elements: new Map(), fields: [], subset: false, external: false, doctype: false, hasExternal: false };
      this.parts.set(name, p);
    }
    return p;
  }

  /** Adds an edit to an element. Removing it covers every other edit. */
  element(part: string, index: number, edit: ElementEdit): void {
    const elements = this.part(part).elements;
    const had = elements.get(index);
    if (!had) {
      elements.set(index, edit);
      return;
    }
    if (had.remove || edit.remove) {
      elements.set(index, { remove: true });
      return;
    }
    had.unwrap ||= edit.unwrap;
    if (edit.strip) had.strip = new Set([...(had.strip ?? []), ...edit.strip]);
    if (edit.set) had.set = new Map([...(had.set ?? []), ...edit.set]);
    if (edit.content !== undefined) had.content = edit.content;
  }

  dropRelationship(relsPart: string, id: string): void {
    let ids = this.rels.get(relsPart);
    if (!ids) {
      ids = new Set();
      this.rels.set(relsPart, ids);
    }
    ids.add(id);
  }
}

/** The entries of a ZIP package the rewrite needs besides the plan: what the format rules found. */
export interface OdfLayout {
  mimetype?: ZipEntry;
  manifest?: ZipEntry;
  /** The media type the package declares, which the output's mimetype entry holds. */
  mediaType: string;
}

/** Everything a strip finding names in a compound file, for the rewrite. */
export class CompoundPlan {
  /** Streams and storages that go, each with everything under it. */
  readonly drop = new Set<CfbEntry>();
  /** Streams whose content a plugin replaced. */
  readonly replace = new Map<CfbEntry, SpillSink>();
  /** Storages whose content a plugin replaced with a compound file of its own, read from what it wrote. */
  readonly grafts = new Map<CfbEntry, CompoundFile>();
}

/** A range of an RTF file the rewrite leaves out, or replaces with `text` and then the bytes of the file at `copy`. */
export interface RtfCut {
  start: number;
  end: number;
  text?: Uint8Array;
  copy?: [number, number];
}

/** What the output of a Flat OPC package keeps of the file besides its parts. */
export interface FlatOpcLayout {
  /** The prolog's mso-application processing instructions, which name the application that opens the file. */
  instructions: string[];
  /** Each part's attributes besides its name and content type, such as pkg:compression, by key. */
  extra: Map<string, Array<[string, string]>>;
}
