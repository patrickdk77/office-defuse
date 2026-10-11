import { spawnSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as v8 from 'node:v8';
import * as vm from 'node:vm';
import * as zlib from 'node:zlib';
import { crc32 } from '../../src/crc32';
import type { ByteSink, DefuseFinding } from '../../src/types';
import { ZipWriter } from '../../src/zip-writer';

/** The package root. This file runs from build-test/test/helpers. */
export const root = path.join(__dirname, '..', '..', '..');
export const fixtures = path.join(root, 'test', 'fixtures');
export const fixture = (...parts: string[]) => fs.readFileSync(path.join(fixtures, ...parts));

const cli = path.join(root, 'dist', 'cli.js');
/** Runs the built command-line tool. */
export const runCli = (args: string[], opts: { input?: Buffer; env?: NodeJS.ProcessEnv; cwd?: string } = {}) =>
  spawnSync(process.execPath, [cli, ...args], { input: opts.input, env: { ...process.env, ...opts.env }, cwd: opts.cwd, maxBuffer: 64 * 1024 * 1024 });

export const kind = (f: DefuseFinding) => `${f.category}/${f.detail}`;
export const kinds = (i: { findings: DefuseFinding[] }) => i.findings.map(kind);
/** The findings of one kind, such as 'LINK/SAFE'. */
export const of = (i: { findings: DefuseFinding[] }, k: string) => i.findings.filter(f => kind(f) === k);
export const has = (i: { findings: DefuseFinding[] }, k: string) => of(i, k).length > 0;

/** Returns `value`, or fails the test when it is null or undefined. `what` names the value in the failure. */
export function must<T>(value: T | null | undefined, what: string): T {
  if (value === null || value === undefined) throw new Error(`${what} is missing`);
  return value;
}

/** The error a promise rejects with, or undefined when it resolves. */
export async function failure(p: Promise<unknown>): Promise<unknown> {
  try {
    await p;
  } catch (e) {
    return e;
  }
  return undefined;
}

v8.setFlagsFromString('--expose_gc');
const collect = vm.runInNewContext('gc') as () => void;

/**
 * The most memory, in the heap and outside it, that `run` holds on to at once. Garbage is collected before each
 * sample, so the number counts what is live, not what the collector has yet to free.
 */
export async function livePeak(run: () => Promise<unknown>): Promise<number> {
  const live = () => {
    collect();
    const m = process.memoryUsage();
    return m.heapUsed + m.external;
  };
  const base = live();
  let peak = 0;
  const timer = setInterval(() => {
    peak = Math.max(peak, live() - base);
  }, 25);
  try {
    await run();
  } finally {
    clearInterval(timer);
  }
  return peak;
}

/** A sink that collects what the ZIP writer writes. */
export function bufferSink(): ByteSink & { result(): Buffer; closed: boolean } {
  const chunks: Buffer[] = [];
  const sink = {
    closed: false,
    async write(chunk: Uint8Array) {
      chunks.push(Buffer.from(chunk));
    },
    async close() {
      sink.closed = true;
    },
    result: () => Buffer.concat(chunks),
  };
  return sink;
}

/** A ZIP of the entries in order, written by the package's own writer. */
export async function zipOf(entries: Array<[string, string | Uint8Array]>, options: { stored?: string[] } = {}): Promise<Buffer> {
  const sink = bufferSink();
  const w = new ZipWriter(sink);
  for (const [name, data] of entries) await w.add(name, typeof data === 'string' ? Buffer.from(data, 'utf8') : data, { compress: !options.stored?.includes(name) });
  await w.finish();
  return sink.result();
}

/** Header fields a raw entry can set apart from what its data says. */
interface Fields {
  flags: number;
  method: number;
  crc: number;
  csize: number;
  usize: number;
  name: Uint8Array;
}

/** One entry of an archive built byte by byte, to break one rule of the format at a time. */
export interface RawEntry {
  name: string | Uint8Array;
  data: string | Uint8Array;
  /** 0 stored, 8 deflate, the default. */
  method?: number;
  flags?: number;
  /** The bytes stored, in place of the data compressed by `method`. */
  payload?: Uint8Array;
  /** Fields of the local header that differ from the entry's own. */
  local?: Partial<Fields>;
  /** Fields of the central directory header that differ from the entry's own. */
  central?: Partial<Fields & { offset: number; disk: number }>;
  localExtra?: Uint8Array;
  centralExtra?: Uint8Array;
  /** Bytes ahead of the local header. */
  before?: Uint8Array;
  /** Bytes after the data, such as a data descriptor. */
  after?: Uint8Array;
}

export interface RawEnd {
  disk: number;
  cdDisk: number;
  countDisk: number;
  count: number;
  cdSize: number;
  cdOffset: number;
  commentLength: number;
}

const u16 = (v: number) => Buffer.from([v & 0xff, (v >>> 8) & 0xff]);
const u32 = (v: number) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(v >>> 0);
  return b;
};
const u64 = (v: number | bigint) => {
  const b = Buffer.alloc(8);
  b.writeBigUInt64LE(BigInt(v));
  return b;
};
export const le = { u16, u32, u64 };

/** An extra field record. */
export const extraRecord = (id: number, body: Uint8Array) => Buffer.concat([u16(id), u16(body.length), body]);

/**
 * Lays out local headers, data, the central directory and the end record. `prefix` goes first, and with `relative`
 * the offsets do not count it. `end` overrides fields of the end record, `cdOffsetDelta` moves its directory offset,
 * and `centralOrder` lists the entries in the directory in another order than in the file.
 */
export function rawZip(
  entries: RawEntry[],
  o: { prefix?: Uint8Array; relative?: boolean; trailing?: Uint8Array; comment?: Uint8Array; end?: Partial<RawEnd>; cdOffsetDelta?: number; centralOrder?: number[] } = {},
): Buffer {
  const prefix = Buffer.from(o.prefix ?? []);
  const out: Buffer[] = [prefix];
  let length = prefix.length;
  const put = (b: Uint8Array) => {
    out.push(Buffer.from(b));
    length += b.length;
  };
  const central: Buffer[] = [];
  for (const e of entries) {
    if (e.before) put(e.before);
    const data = typeof e.data === 'string' ? Buffer.from(e.data, 'utf8') : Buffer.from(e.data);
    const method = e.method ?? 8;
    const payload = e.payload ?? (method === 8 ? zlib.deflateRawSync(data) : data);
    const own: Fields = { flags: e.flags ?? 0, method, crc: crc32(data), csize: payload.length, usize: data.length, name: typeof e.name === 'string' ? Buffer.from(e.name, 'utf8') : e.name };
    const offset = length - (o.relative ? prefix.length : 0);
    const l = { ...own, ...e.local };
    const localExtra = e.localExtra ?? new Uint8Array(0);
    put(
      Buffer.concat([u32(0x04034b50), u16(20), u16(l.flags), u16(l.method), u16(0), u16(0x21), u32(l.crc), u32(l.csize), u32(l.usize), u16(l.name.length), u16(localExtra.length), l.name, localExtra]),
    );
    put(payload);
    if (e.after) put(e.after);
    const c = { ...own, offset, disk: 0, ...e.central };
    const centralExtra = e.centralExtra ?? new Uint8Array(0);
    central.push(
      Buffer.concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(c.flags),
        u16(c.method),
        u16(0),
        u16(0x21),
        u32(c.crc),
        u32(c.csize),
        u32(c.usize),
        u16(c.name.length),
        u16(centralExtra.length),
        u16(0),
        u16(c.disk),
        u16(0),
        u32(0),
        u32(c.offset),
        c.name,
        centralExtra,
      ]),
    );
  }
  const cd = Buffer.concat(o.centralOrder ? o.centralOrder.map(i => central[i]) : central);
  const cdOffset = length - (o.relative ? prefix.length : 0) + (o.cdOffsetDelta ?? 0);
  put(cd);
  const comment = Buffer.from(o.comment ?? []);
  const end: RawEnd = { disk: 0, cdDisk: 0, countDisk: entries.length, count: entries.length, cdSize: cd.length, cdOffset, commentLength: comment.length, ...o.end };
  put(Buffer.concat([u32(0x06054b50), u16(end.disk), u16(end.cdDisk), u16(end.countDisk), u16(end.count), u32(end.cdSize), u32(end.cdOffset), u16(end.commentLength), comment]));
  if (o.trailing) put(o.trailing);
  return Buffer.concat(out);
}

/** A signed data descriptor with 32-bit sizes, or 64-bit ones with `wide`. */
export function descriptorOf(data: string | Uint8Array, payload: Uint8Array, wide = false): Buffer {
  const d = typeof data === 'string' ? Buffer.from(data, 'utf8') : data;
  return Buffer.concat([u32(0x08074b50), u32(crc32(d)), ...(wide ? [u64(payload.length), u64(d.length)] : [u32(payload.length), u32(d.length)])]);
}

export const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export const MS = 'http://schemas.microsoft.com/office/2006/relationships/';
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const WORD_MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

export type Rel = [id: string, type: string, target: string, external?: boolean];

export const relsXml = (...rels: Rel[]) =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels
    .map(([id, type, target, external]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`)
    .join('')}</Relationships>`;

export const contentTypesXml = (overrides: Record<string, string>, defaults: Record<string, string> = {}) =>
  `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${Object.entries(
    defaults,
  )
    .map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`)
    .join('')}${Object.entries(overrides)
    .map(([p, t]) => `<Override PartName="/${p}" ContentType="${t}"/>`)
    .join('')}</Types>`;

export const wordXml = (body: string) =>
  `${XML}<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>${body}</w:body></w:document>`;

/** A worksheet part whose sheetData holds `cells`, a row of `<c>` elements. */
export const sheetXml = (cells: string, after = '') => `${XML}<worksheet xmlns="${S}"><sheetData><row r="1">${cells}</row></sheetData>${after}</worksheet>`;

export interface WordOptions {
  /** Further parts, by name. */
  parts?: Record<string, string | Uint8Array>;
  /** Content type overrides beyond the main part. */
  overrides?: Record<string, string>;
  defaults?: Record<string, string>;
  /** Relationships of word/document.xml. */
  rels?: Rel[];
  /** Further package relationships. */
  packageRels?: Rel[];
  /** The main part's content type. */
  main?: string;
  /** A document type declaration for word/document.xml, after its XML declaration. */
  doctype?: string;
}

/** `xml`, which starts with the XML declaration, with a document type declaration whose internal subset is `subset`. */
export const withSubset = (xml: string, subset: string) => xml.replace(XML, `${XML}<!DOCTYPE r [${subset}]>`);

/** A Word package whose document body is `body`. */
export function wordPackage(body: string, o: WordOptions = {}): Promise<Buffer> {
  const entries: Array<[string, string | Uint8Array]> = [
    ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': o.main ?? WORD_MAIN, ...o.overrides }, o.defaults)],
    ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'], ...(o.packageRels ?? []))],
    ['word/document.xml', wordXml(body).replace(XML, XML + (o.doctype ?? ''))],
  ];
  if (o.rels) entries.push(['word/_rels/document.xml.rels', relsXml(...o.rels)]);
  for (const [name, data] of Object.entries(o.parts ?? {})) entries.push([name, data]);
  return zipOf(entries);
}

/** A Word package with one more XML part, word/part.xml, related from the document and typed `type`. */
export const wordWithPart = (part: string, type = 'application/xml', o: WordOptions = {}) =>
  wordPackage('<w:p/>', {
    ...o,
    parts: { 'word/part.xml': part, ...o.parts },
    overrides: { 'word/part.xml': type, ...o.overrides },
    rels: [['rIdPart', `${T}customXml`, 'part.xml'], ...(o.rels ?? [])],
  });

export const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"';

export interface OdfOptions {
  /** The mimetype entry's content, or null for none. */
  mimetype?: string | null;
  /** The manifest root's media type, or null for no manifest at all. */
  manifestType?: string | null;
  /** Further entries, listed in the manifest unless `unlisted` names them. */
  files?: Record<string, string | Uint8Array>;
  unlisted?: string[];
  /** Raw text added inside the manifest. */
  manifestExtra?: string;
  /** The mimetype entry deflated instead of stored. */
  deflateMimetype?: boolean;
}

/** An OpenDocument package whose content.xml body is `body`, an office:text by default. */
export function odfPackage(body: string, o: OdfOptions = {}): Promise<Buffer> {
  const type = o.mimetype === undefined ? 'application/vnd.oasis.opendocument.text' : o.mimetype;
  const manifestType = o.manifestType === undefined ? (type ?? 'application/vnd.oasis.opendocument.text') : o.manifestType;
  const files = { 'content.xml': `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${ODF_NS} office:version="1.3">${body}</office:document-content>`, ...o.files };
  const listed = Object.keys(files).filter(n => !o.unlisted?.includes(n));
  const entries: Array<[string, string | Uint8Array]> = [];
  if (type !== null) entries.push(['mimetype', type]);
  for (const [n, d] of Object.entries(files)) entries.push([n, d]);
  if (manifestType !== null) {
    entries.push([
      'META-INF/manifest.xml',
      `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${manifestType}"/>${listed
        .map(n => `<manifest:file-entry manifest:full-path="${n}" manifest:media-type="text/xml"/>`)
        .join('')}${o.manifestExtra ?? ''}</manifest:manifest>`,
    ]);
  }
  return zipOf(entries, { stored: o.deflateMimetype ? [] : ['mimetype'] });
}

export const textBody = (inner: string) => `<office:body><office:text>${inner}</office:text></office:body>`;

/** The OLE signature and an inert marker, all office-defuse looks at. */
export const OLE = Buffer.concat([Buffer.from('d0cf11e0a1b11ae1', 'hex'), Buffer.from('inert marker')]);
