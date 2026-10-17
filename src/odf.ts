import { type ContainedDecision, decideContained, ownFinding } from './contained';
import { clip, type PlanContext, type Reported, type ZipContext } from './context';
import { StopError } from './errors';
import { type Format, ODF_FORMATS } from './formats';
import { callsDde, formulaCalls } from './formulas';
import { isFragment, isInsidePackage } from './links';
import { attributeKey, type ElementEdit } from './plan';
import { OfficeCategory as C, OfficeDetail as D } from './types';
import type { XmlElement, XmlHandler } from './xml';
import { nameKey, type ZipEntry } from './zip';

export const NS = {
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  script: 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  form: 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  database: 'urn:oasis:names:tc:opendocument:xmlns:database:1.0',
  chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
};

/**
 * Folders of code and settings. Each is reported by name, as a whole, so the files in them are inflated and checked
 * but not parsed.
 */
const FOLDERS: Array<[string, D, C]> = [
  ['basic', D.OdfBasic, C.Macro],
  ['scripts', D.OdfScripts, C.Macro],
  ['dialogs', D.OdfDialogs, C.Macro],
  ['configurations2', D.OdfConfiguration, C.Ui],
];

/** Event names that fire as the document opens. Focus and activation count only for the document's own events. */
const ON_OPEN = new Set(['load', 'new', 'create', 'loadfinished', 'viewcreated']);
const ON_OPEN_DOCUMENT = new Set(['focus', 'activate']);

const LINK_ELEMENTS = new Set([`${NS.text} a`, `${NS.draw} a`, `${NS.draw} area-rectangle`, `${NS.draw} area-circle`, `${NS.draw} area-polygon`]);
const EXTERNAL_KIND = new Map<string, D>([
  [`${NS.meta} template`, D.Template],
  [`${NS.meta} auto-reload`, D.Reload],
  [`${NS.draw} image`, D.Image],
  [`${NS.draw} fill-image`, D.Image],
  [`${NS.style} background-image`, D.Image],
  [`${NS.text} section-source`, D.Include],
  [`${NS.table} table-source`, D.Include],
  [`${NS.draw} object`, D.Object],
  [`${NS.draw} object-ole`, D.Object],
  [`${NS.draw} floating-frame`, D.Frame],
]);

/** Elements whose reference always leaves the package: it resolves against the document's own location. */
const ALWAYS_OUTSIDE = new Set([`${NS.meta} template`, `${NS.meta} auto-reload`, `${NS.text} section-source`, `${NS.table} table-source`]);

/**
 * LibreOffice reads urn:oasis:names:tc:<any>:xmlns:<name>:1.<any> as the OpenDocument namespace <name> 1.0, so this
 * does too. The comparison ignores case, which can only find more.
 */
function odfNs(ns: string): string {
  const m = /^urn:oasis:names:tc:[^:]*:xmlns:([^:]*):1\.[^:]+$/i.exec(ns);
  return m ? `urn:oasis:names:tc:opendocument:xmlns:${m[1].toLowerCase()}:1.0` : ns;
}

/** An attribute by normalized namespace and local name. */
export function oattr(el: XmlElement, ns: string, local: string): string | undefined {
  for (const a of el.attributes) if (a.local === local && odfNs(a.ns) === ns) return a.value;
  return undefined;
}

/** The ODF_NS-normalized namespace of `ns`, exported for the rewrite. */
export const odfNamespace = (ns: string) => odfNs(ns);

/** The edit that takes an element's xlink:href away, under whatever namespace URI the attribute has. */
function stripHref(el: XmlElement): ElementEdit {
  return { strip: new Set(el.attributes.filter(a => a.local === 'href' && odfNs(a.ns) === NS.xlink).map(a => attributeKey(a.ns, a.local))) };
}

/** The xlink:href of an element, for the rewrite. */
export const hrefOf = (el: XmlElement) => oattr(el, NS.xlink, 'href');

/** A path inside the package as the manifest and the zip name it: without "./" and leading slashes. */
export const packagePath = (href: string) => nameKey(href.trim().replace(/^(\.\/)+/, ''));

/**
 * The rules that read inside one XML entry of an ODF package, or inside a flat OpenDocument file. A flat file has no
 * entries, so in one every relative reference names a file beside the document, which loads from outside it.
 */
export class EntryRules implements XmlHandler {
  /** Depth of the office:scripts element being read, 0 outside one. */
  private scriptsDepth = 0;
  /** An office:script element being read, with whether it holds anything. */
  private script?: { depth: number; language: string; content: boolean; index: number };
  /** A link element being read, whose text and title go into the label check at its end. */
  private label?: { index: number; href: string; text: string; title?: string; unwrap: boolean };

  constructor(
    private readonly ctx: PlanContext,
    private readonly entry: string,
    /** Whether an object refers to each entry, by key. Only entries are kept, so a part full of references costs no memory. */
    private readonly oleRefs: Map<string, boolean>,
    private readonly flat = false,
  ) {}

  /**
   * Whether a reference names an entry of the package. LibreOffice's vnd.sun.star.Package: URLs name one too. In a
   * flat file no entry exists, so only such a URL stays inside, naming nothing.
   */
  private inside(href: string): boolean {
    const own = /^\s*vnd\.sun\.star\.package:(.*)$/is.exec(href);
    if (own) return this.flat || isInsidePackage(own[1]);
    return !this.flat && isInsidePackage(href);
  }

  private report(category: C, detail: D, data?: Record<string, string | number>): Reported {
    return this.ctx.report(category, detail, this.entry, data);
  }

  /** Records the edit for the element when the finding strips. */
  private edit(r: Reported, el: XmlElement, edit: ElementEdit = { remove: true }): void {
    if (r.action === 'strip') this.ctx.plan.element(this.entry, el.index, edit);
  }

  start(raw: XmlElement): void {
    if (this.script) this.script.content = true;
    const el = { ...raw, ns: odfNs(raw.ns) };
    const key = `${el.ns} ${el.local}`;
    const href = oattr(el, NS.xlink, 'href');
    if (el.ns === NS.office && el.local === 'scripts' && !this.scriptsDepth) this.scriptsDepth = el.depth;
    else if (el.ns === NS.office && el.local === 'script') this.script = { depth: el.depth, language: oattr(el, NS.script, 'language') ?? '', content: false, index: el.index };
    else if (el.ns === NS.script && el.local === 'event-listener') this.scriptEvent(el, href);
    else if (el.ns === NS.presentation && el.local === 'event-listener') this.presentationEvent(el, href);
    else if ((el.ns === NS.text && el.local === 'dde-connection-decl') || (el.ns === NS.office && el.local === 'dde-source')) {
      this.edit(this.report(C.Dde, D.Link, { application: clip(oattr(el, NS.office, 'dde-application') ?? ''), topic: clip(oattr(el, NS.office, 'dde-topic') ?? '') }), raw);
    } else if (el.ns === NS.draw && el.local === 'applet') {
      this.edit(this.report(C.Embedded, D.Applet, { code: clip(oattr(el, NS.draw, 'code') ?? ''), ...(href ? { target: clip(href) } : {}) }), raw);
      return;
    } else if (el.ns === NS.draw && el.local === 'plugin') {
      this.edit(this.report(C.Embedded, D.Plugin, { mimeType: clip(oattr(el, NS.draw, 'mime-type') ?? ''), ...(href ? { target: clip(href) } : {}) }), raw);
      return;
    } else if (el.ns === NS.draw && el.local === 'object-ole' && (href === undefined || this.inside(href))) {
      // The object's entry may still be ahead in the archive, so its plugin verdict decides this element later. The
      // preview image beside the element stays either way.
      const r = this.report(C.Embedded, D.OleObject, href ? { target: clip(href) } : undefined);
      this.ctx.plan.objects.push({ part: this.entry, index: el.index, target: href ? packagePath(href) : '', finding: r.finding, strips: r.action === 'strip' });
      if (href && this.oleRefs.has(packagePath(href))) this.oleRefs.set(packagePath(href), true);
    }
    for (const a of el.attributes) {
      // The cell keeps its value without the formula.
      if (odfNs(a.ns) !== NS.table || (a.local !== 'formula' && a.local !== 'expression')) continue;
      const strip = { strip: new Set([attributeKey(a.ns, a.local)]) };
      if (callsDde(a.value)) this.edit(this.report(C.Dde, D.Formula, { formula: clip(a.value) }), raw, strip);
      for (const [fn, kind] of formulaCalls(a.value)) this.edit(this.report(kind === 'load' ? C.External : C.Macro, D.Formula, { function: clip(fn), formula: clip(a.value) }), raw, strip);
    }
    if (href === undefined || isFragment(href) || (el.ns === NS.script && el.local === 'event-listener') || (el.ns === NS.presentation && el.local === 'event-listener')) return;
    // A database document names its own embedded database this way.
    if (el.ns === NS.database && el.local === 'connection-resource' && /^sdbc:embedded:/i.test(href.trim())) return;
    // A chart names the document that holds it, its data source, as "..".
    if (el.ns === NS.chart && el.local === 'chart' && /^\.\.\/?$/.test(href.trim())) return;
    if (LINK_ELEMENTS.has(key)) {
      // The text of a link stays without it. An image map area is nothing but its link.
      const r = this.ctx.link(href, this.entry, { element: clip(el.qname) });
      this.edit(r, raw, el.local === 'a' ? { unwrap: true } : { remove: true });
      if (r.action !== 'strip' && !this.label) this.label = { index: el.index, href, text: '', title: oattr(el, NS.office, 'title') ?? oattr(el, NS.xlink, 'title'), unwrap: el.local === 'a' };
    } else if (el.ns === NS.form && el.local === 'form') {
      // A form's own href is where it sends its fields, which goes as pdf-defuse's form submit actions go.
      this.edit(this.report(C.Action, D.SubmitForm, { element: clip(el.qname), target: clip(href, 2000) }), raw, stripHref(raw));
    } else if (el.ns === NS.form) this.edit(this.ctx.link(href, this.entry, { element: clip(el.qname) }), raw, stripHref(raw));
    else if (ALWAYS_OUTSIDE.has(key) || !this.inside(href)) {
      this.edit(this.ctx.external(EXTERNAL_KIND.get(key) ?? D.Other, href, this.entry, { element: clip(el.qname) }, clip(el.local)), raw);
    }
  }

  text(t: string): void {
    if (this.script && t.trim()) this.script.content = true;
    if (this.label && this.label.text.length < 4096) this.label.text += t;
  }

  end(el: XmlElement): void {
    const label = this.label;
    if (label && label.index === el.index) {
      this.label = undefined;
      const r = this.ctx.labelMismatch(label.href, [label.text, label.title], this.entry, { element: clip(el.qname) });
      if (r) this.edit(r, el, label.unwrap ? { unwrap: true } : { remove: true });
    }
    if (this.scriptsDepth === el.depth && odfNs(el.ns) === NS.office && el.local === 'scripts') this.scriptsDepth = 0;
    if (this.script && this.script.depth === el.depth) {
      if (this.script.content) this.edit(this.report(C.Macro, /basic/i.test(this.script.language) ? D.OdfBasic : D.OdfScripts, { language: clip(this.script.language) }), el, { remove: true });
      this.script = undefined;
    }
  }

  private scriptEvent(el: XmlElement, href: string | undefined): void {
    const event = oattr(el, NS.script, 'event-name') ?? '';
    // LibreOffice runs nothing for a listener whose script URL and macro name are both empty. An empty href does not
    // hide a macro name.
    const target = [href, oattr(el, NS.script, 'macro-name')].find(v => v !== undefined && v !== '');
    if (target === undefined) return;
    const name = event
      .slice(event.lastIndexOf(':') + 1)
      .toLowerCase()
      .replace(/^on/, '')
      .replace(/-/g, '');
    const onOpen = ON_OPEN.has(name) || (this.scriptsDepth > 0 && ON_OPEN_DOCUMENT.has(name));
    const data = { event: clip(event), target: clip(target), language: clip(oattr(el, NS.script, 'language') ?? '') };
    this.edit(this.report(C.Macro, onOpen ? D.EventOnOpen : D.Event, data), el);
  }

  private presentationEvent(el: XmlElement, href: string | undefined): void {
    const action = oattr(el, NS.presentation, 'action') ?? '';
    const data = { event: clip(oattr(el, NS.script, 'event-name') ?? ''), action: clip(action), ...(href ? { target: clip(href) } : {}) };
    if (action === 'execute') this.edit(this.report(C.Action, D.RunProgram, data), el);
    else if (action === 'verb') this.edit(this.report(C.Action, D.OleVerb, data), el);
    else if (href !== undefined && !isFragment(href)) this.edit(this.ctx.link(href, this.entry, data), el);
  }
}

export interface Manifest {
  paths: Set<string>;
  root?: string;
  encrypted: boolean;
}

/** The manifest. Only paths that name an entry of the archive are kept, so a manifest of millions costs no memory. */
async function readManifest(ctx: ZipContext, entry: ZipEntry): Promise<Manifest> {
  const keys = new Set(ctx.zip.entries.map(e => nameKey(e.name)));
  const m: Manifest = { paths: new Set(), encrypted: false };
  await ctx.parse(entry, {
    start(el) {
      if (odfNs(el.ns) !== NS.manifest) return;
      if (el.local === 'encryption-data') m.encrypted = true;
      if (el.local !== 'file-entry') return;
      const path = oattr(el, NS.manifest, 'full-path');
      if (path === undefined) return;
      if (path === '/') {
        if (m.root !== undefined) ctx.report(C.Package, D.DuplicateEntry, entry.name, { kind: 'manifest path', name: '/' });
        m.root = oattr(el, NS.manifest, 'media-type') ?? '';
        return;
      }
      const key = packagePath(path);
      if (!keys.has(key)) return;
      if (m.paths.has(key)) ctx.report(C.Package, D.DuplicateEntry, entry.name, { kind: 'manifest path', name: clip(path) });
      m.paths.add(key);
    },
  });
  return m;
}

/** The code folder an entry sits in, as the folder's path and its finding. */
export function folderOf(name: string): { prefix: string; detail: D; category: C } | undefined {
  const segs = name.split('/');
  for (let i = 0; i < segs.length - 1; i++) {
    const hit = FOLDERS.find(([folder]) => folder === segs[i].toLowerCase());
    if (hit) return { prefix: `${segs.slice(0, i + 1).join('/')}/`, detail: hit[1], category: hit[2] };
  }
  return undefined;
}

/** The most of the mimetype entry that is kept. The longest media type in the table is 56 characters. */
const MAX_MIMETYPE = 256;

/** What names an OpenDocument package: the media type its format comes from, and the entries that give it. */
export interface OdfNaming {
  format: Format;
  /** The media type the format was named by. */
  mediaType: string;
  mimetype?: ZipEntry;
  manifestEntry?: ZipEntry;
  manifest?: Manifest;
}

/**
 * Names an OpenDocument package's format by the media type in its mimetype entry or, without one, its manifest's root
 * entry. Throws when the two name different formats, or the media type is not one this package reads. `layout` hears,
 * with the entry it concerns, of each way the mimetype entry and the manifest leave the layout OpenDocument asks for.
 */
export async function nameOdf(ctx: ZipContext, layout: (location: string, reason: string) => void = () => undefined): Promise<OdfNaming> {
  const zip = ctx.zip;
  const mimetype = zip.entries.find(e => e.name === 'mimetype');
  const manifestEntry = zip.get('META-INF/manifest.xml');
  const note = (reason: string) => layout(mimetype?.name ?? 'META-INF/manifest.xml', reason);
  let declared: string | undefined;
  if (mimetype) {
    let head = Buffer.alloc(0);
    await ctx.drain(mimetype, c => {
      if (head.length < MAX_MIMETYPE) head = Buffer.concat([head, c.subarray(0, MAX_MIMETYPE - head.length)]);
    });
    // A longer entry holds no media type this package reads, whatever its start says.
    declared = head.toString('latin1') + (mimetype.size > MAX_MIMETYPE ? '...' : '');
    const first = zip.entries.reduce((a, b) => (b.localOffset < a.localOffset ? b : a));
    if (first !== mimetype) note('mimetype is not the first entry');
    if (mimetype.method !== 0) note('mimetype is compressed');
    if (mimetype.localExtra) note('mimetype has an extra field');
    if (declared !== declared.trim()) note('mimetype has white space around the media type');
    declared = declared.trim();
  } else note('there is no mimetype entry');
  const manifest = manifestEntry ? await readManifest(ctx, manifestEntry) : undefined;
  if (!manifest) note('there is no manifest');
  const root = manifest?.root;
  if (declared !== undefined && root !== undefined && declared !== root && ODF_FORMATS.has(declared) && ODF_FORMATS.has(root)) {
    throw new StopError(C.Package, D.AmbiguousFormat, 'The mimetype entry and the manifest name different formats', { mimetype: clip(declared), manifest: clip(root) });
  }
  const mediaType = declared ?? root ?? '';
  const format = ODF_FORMATS.get(mediaType);
  if (!format) throw new StopError(C.Package, D.Unrecognized, 'The OpenDocument media type is not one this package reads', { mediaType: clip(mediaType) });
  return { format, mediaType, mimetype, manifestEntry, manifest };
}

/** Inspects an OpenDocument package. */
export async function inspectOdf(ctx: ZipContext): Promise<void> {
  const zip = ctx.zip;
  const { format, mediaType: type, mimetype, manifestEntry, manifest } = await nameOdf(ctx, (location, reason) => ctx.report(C.Package, D.OdfLayout, location, { reason }));
  ctx.format = format;
  if (manifest?.encrypted) throw new StopError(C.Encrypted, D.PasswordRequired, 'The manifest lists encrypted entries');

  const folders = new Map<string, { detail: D; category: C; entries: number }>();
  const oleRefs = new Map(zip.entries.map(e => [nameKey(e.name), false]));
  const oleEntries: ZipEntry[] = [];
  /** Entries that are whole files in formats of their own, such as pictures, which go to the plugins. */
  const wholeFiles: ZipEntry[] = [];
  for (const entry of zip.entries) {
    if (entry === mimetype || entry === manifestEntry) continue;
    const folder = folderOf(entry.name);
    if (folder) {
      // LibreOffice writes an empty Configurations2/accelerator/current.xml into the documents it saves. An empty
      // file holds no code or setting.
      if (entry.size > 0) {
        const f = folders.get(folder.prefix) ?? { detail: folder.detail, category: folder.category, entries: 0 };
        f.entries++;
        folders.set(folder.prefix, f);
      }
      await ctx.drain(entry);
    } else if (!entry.directory && /\.xml$/i.test(entry.name)) await ctx.parse(entry, new EntryRules(ctx, entry.name, oleRefs));
    else if (await ctx.drain(entry)) oleEntries.push(entry);
    else if (!entry.directory && !/^meta-inf\//i.test(entry.name)) wholeFiles.push(entry);
  }
  const plan = ctx.plan;
  // A folder goes whole, its empty files with it.
  for (const [prefix, f] of folders) if (ctx.report(f.category, f.detail, prefix, { entries: f.entries }).action === 'strip') plan.folders.add(`${nameKey(prefix)}/`);
  // OLE entries go to the plugins. When none takes one, the elements that refer to it decide, or EMBEDDED/OLE_FILE.
  const decided = new Map<string, ContainedDecision>();
  for (const e of oleEntries) {
    const key = nameKey(e.name);
    const decision = await decideContained(ctx, { name: e.name, location: e.name, chunks: () => zip.read(e) });
    if (decision) {
      decided.set(key, decision);
      if (!decision.keep) plan.drop.add(key);
      else if (decision.replacement) plan.replace.set(key, decision.replacement);
    }
    if (!oleRefs.get(key)) {
      const r = ownFinding(ctx, decision, C.Embedded, D.OleFile, e.name, { part: e.name });
      if (!decision && r.action === 'strip') plan.drop.add(key);
    }
  }
  // A whole file no plugin takes stays, since no rule names it.
  for (const e of wholeFiles) {
    const decision = await decideContained(ctx, { name: e.name, location: e.name, chunks: () => zip.read(e) });
    if (decision && !decision.keep) plan.drop.add(nameKey(e.name));
    else if (decision?.replacement) plan.replace.set(nameKey(e.name), decision.replacement);
  }
  const files = new Set(zip.entries.filter(e => !e.directory).map(e => nameKey(e.name)));
  for (const o of plan.objects) {
    const decision = decided.get(o.target);
    // A plugin's answer on the object's file is final, so the object's own finding is taken back.
    if (decision?.answered) ctx.retract(o.finding);
    const kept = decision?.keep;
    // A plugin kept the object's file, so the object stays.
    if (kept === true) continue;
    if (kept === false || o.strips) {
      plan.element(o.part, o.index, { remove: true });
      if (files.has(o.target)) plan.drop.add(o.target);
    }
  }
  if (manifest) {
    for (const e of zip.entries) {
      if (e.directory || e === mimetype || e === manifestEntry || /^meta-inf\//i.test(e.name)) continue;
      if (!manifest.paths.has(nameKey(e.name)) && ctx.report(C.Package, D.UnreferencedPart, e.name, { part: e.name }).action === 'strip') plan.drop.add(nameKey(e.name));
    }
  }
  plan.odf = { mimetype, manifest: manifestEntry, mediaType: type };
}
