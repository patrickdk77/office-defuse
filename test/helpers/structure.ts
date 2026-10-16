import { bufferSource } from '../../src/io';
import { type XmlElement, XmlTokenizer } from '../../src/xml';
import { nameKey, ZipArchive } from '../../src/zip';

const LIMITS = { depth: 256, tokenSize: 1 << 24, entityDepth: 16, entityExpansions: 100_000, entitySize: 1 << 24 };
const ZIP_LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const R_NS = new Set(['http://schemas.openxmlformats.org/officedocument/2006/relationships', 'http://purl.oclc.org/ooxml/officedocument/relationships']);
const MANIFEST_NS = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
const VML_TYPE = 'application/vnd.openxmlformats-officedocument.vmldrawing';

async function read(zip: ZipArchive, name: string): Promise<Buffer> {
  const entry = zip.get(name);
  if (!entry) throw new Error(`${name} is missing`);
  const chunks: Buffer[] = [];
  for await (const c of zip.read(entry)) chunks.push(Buffer.from(c));
  return Buffer.concat(chunks);
}

function elements(bytes: Uint8Array, part: string, vml = false): XmlElement[] {
  const out: XmlElement[] = [];
  const t = new XmlTokenizer({ start: el => out.push(el) }, LIMITS, part, { vml });
  t.write(bytes);
  t.end();
  return out;
}

const value = (el: XmlElement, local: string) => el.attributes.find(a => a.ns === '' && a.local === local)?.value;

function resolve(relsName: string, target: string): string {
  const folder = relsName.slice(0, relsName.toLowerCase().indexOf('_rels/'));
  const out: string[] = [];
  for (const seg of (target.startsWith('/') ? target.slice(1) : folder + target).split('#')[0].split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return nameKey(out.join('/'));
}

/**
 * The problems a reader of a rewritten package could trip on, checked with code rather than with Office: the archive
 * reads with no stray bytes, every XML part is well-formed, every part has a content type, every relationship that
 * points inside the package names a part that is there, every relationship Id a part uses is in its relationships
 * part, and an OpenDocument package lists exactly its entries in its manifest, behind a stored mimetype entry.
 */
export async function structureProblems(bytes: Uint8Array): Promise<string[]> {
  const problems: string[] = [];
  const zip = await ZipArchive.open(bufferSource(bytes), ZIP_LIMITS);
  for (const n of zip.notes) problems.push(`container: ${n.detail}`);
  const keys = new Set(zip.entries.map(e => nameKey(e.name)));
  if (zip.get('[Content_Types].xml')) {
    const types = elements(await read(zip, '[Content_Types].xml'), '[Content_Types].xml');
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const el of types) {
      if (el.ns !== TYPES_NS) continue;
      if (el.local === 'Default') defaults.set((value(el, 'Extension') ?? '').toLowerCase(), (value(el, 'ContentType') ?? '').toLowerCase());
      if (el.local === 'Override') {
        const k = nameKey((value(el, 'PartName') ?? '').replace(/^\//, ''));
        if (!keys.has(k)) problems.push(`override for a missing part ${k}`);
        overrides.set(k, (value(el, 'ContentType') ?? '').toLowerCase());
      }
    }
    const typeOf = (name: string) => overrides.get(nameKey(name)) ?? defaults.get(/\.([^./]*)$/.exec(name)?.[1]?.toLowerCase() ?? '');
    const ids = new Map<string, Set<string>>();
    for (const e of zip.entries) {
      if (e.directory || e.name === '[Content_Types].xml') continue;
      const ct = typeOf(e.name);
      if (ct === undefined) problems.push(`no content type for ${e.name}`);
      const m = /^(.*?)_rels\/([^/]*)\.rels$/i.exec(e.name);
      if (!m) continue;
      const owner = nameKey(m[1] + m[2]);
      if (owner !== '' && !keys.has(owner)) problems.push(`relationships of a missing part ${e.name}`);
      const set = new Set<string>();
      for (const el of elements(await read(zip, e.name), e.name)) {
        if (el.ns !== RELS_NS || el.local !== 'Relationship') continue;
        set.add(value(el, 'Id') ?? '');
        const target = value(el, 'Target') ?? '';
        if ((value(el, 'TargetMode') ?? 'Internal') === 'Internal' && !/^[a-z][a-z0-9+.-]*:/i.test(target) && !keys.has(resolve(e.name, target)))
          problems.push(`relationship to a missing part ${target} in ${e.name}`);
      }
      ids.set(owner, set);
    }
    for (const e of zip.entries) {
      const ct = typeOf(e.name);
      if (e.directory || ct === undefined || !(ct.endsWith('+xml') || ct === 'application/xml' || ct === VML_TYPE)) continue;
      // PowerPoint writes an add-in's presentation part empty.
      if (e.size === 0 && ct === 'application/vnd.ms-powerpoint.addin.macroenabled.main+xml') continue;
      const known = ids.get(nameKey(e.name)) ?? new Set<string>();
      for (const el of elements(await read(zip, e.name), e.name, ct === VML_TYPE)) {
        for (const a of el.attributes) {
          const relationship = R_NS.has(a.ns.toLowerCase()) || (a.ns === 'urn:schemas-microsoft-com:office:office' && a.local === 'relid');
          if (relationship && a.value !== '' && !known.has(a.value)) problems.push(`${e.name} names relationship ${a.value}, which is not there`);
        }
      }
    }
  } else {
    const first = zip.entries.reduce((a, b) => (b.localOffset < a.localOffset ? b : a));
    if (first.name !== 'mimetype' || first.method !== 0 || first.localExtra) problems.push('mimetype is not the first entry, stored, without an extra field');
    const listed = new Set<string>();
    for (const el of elements(await read(zip, 'META-INF/manifest.xml'), 'META-INF/manifest.xml')) {
      if (el.ns !== MANIFEST_NS || el.local !== 'file-entry') continue;
      const path = el.attributes.find(a => a.local === 'full-path')?.value ?? '';
      if (path === '/' || path.endsWith('/')) continue;
      listed.add(nameKey(path));
      if (!keys.has(nameKey(path))) problems.push(`manifest lists a missing entry ${path}`);
    }
    for (const e of zip.entries) {
      if (e.directory || e.name === 'mimetype' || /^meta-inf\//i.test(e.name)) continue;
      if (!listed.has(nameKey(e.name))) problems.push(`manifest leaves out ${e.name}`);
      // The rules read the XML entries outside the code and settings folders, which hold any bytes.
      if (/\.xml$/i.test(e.name) && !/(^|\/)(basic|scripts|dialogs|configurations2)\//i.test(e.name)) elements(await read(zip, e.name), e.name);
    }
  }
  return problems;
}

/** The content of one entry of a ZIP archive, as text, or undefined when the archive has no such entry. */
export async function entryText(bytes: Uint8Array, name: string): Promise<string | undefined> {
  const zip = await ZipArchive.open(bufferSource(bytes), ZIP_LIMITS);
  return zip.get(name) ? (await read(zip, name)).toString('utf8') : undefined;
}

/** The names of a ZIP archive's entries, in the order of its central directory. */
export async function entryNames(bytes: Uint8Array): Promise<string[]> {
  return (await ZipArchive.open(bufferSource(bytes), ZIP_LIMITS)).entries.map(e => e.name);
}
