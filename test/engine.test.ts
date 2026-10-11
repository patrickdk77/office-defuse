import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { expect } from 'chai';
import { type ByteSource, fileSource, inspectOffice, inspectOfficeSource } from '../src';
import { TimeLimitError } from '../src/errors';
import { contentTypesXml, failure, fixture, fixtures, kinds, livePeak, must, relsXml, root, T, WORD_MAIN, withSubset, wordPackage, wordXml, XML, zipOf } from './helpers/util';

const sourceOf = (bytes: Buffer, read?: (offset: number, length: number) => Promise<Uint8Array>): ByteSource => ({
  size: async () => bytes.length,
  read: read ?? (async (offset, length) => bytes.subarray(offset, offset + length)),
});

describe('engine', () => {
  it('gives the same result from a file source as from a buffer, and closes the source once', async () => {
    const file = path.join(fixtures, 'docs', 'excel.xlsm');
    let closed = 0;
    const src = fileSource(file);
    const close = must(src.close, 'close');
    const counting: ByteSource = {
      size: () => src.size(),
      read: (o, n) => src.read(o, n),
      close: async () => {
        closed++;
        await close();
      },
    };
    expect(await inspectOfficeSource(counting)).to.deep.equal(await inspectOffice(fs.readFileSync(file)));
    expect(closed).to.equal(1);
  });

  it('reads a file that shrinks while it is read only as far as it goes', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-shrink-'));
    try {
      const file = path.join(dir, 'a.bin');
      fs.writeFileSync(file, Buffer.alloc(100, 1));
      const src = fileSource(file);
      expect(await src.size()).to.equal(100);
      fs.truncateSync(file, 40);
      expect((await src.read(0, 100)).length).to.equal(40);
      await must(src.close, 'close')();
      // A source closed twice, or one that never opened, closes quietly.
      await must(src.close, 'close')();
      await must(fileSource(path.join(dir, 'none')).close, 'close')();
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('throws errors from the source instead of judging the file', async () => {
    const missing = await failure(inspectOfficeSource(fileSource(path.join(os.tmpdir(), 'office-defuse-no-such-file.docx'))));
    expect((missing as NodeJS.ErrnoException).code).to.equal('ENOENT');
    const bytes = fixture('lo', 'clean.docx');
    const custom = new Error('bucket went away');
    expect(await failure(inspectOfficeSource(sourceOf(bytes, async () => Promise.reject(custom))))).to.equal(custom);
    expect(await failure(inspectOfficeSource({ size: async () => Promise.reject(custom), read: async () => new Uint8Array(0) }))).to.equal(custom);
  });

  it('fails with EIO when a read returns more or fewer bytes than asked', async () => {
    const bytes = fixture('lo', 'clean.docx');
    const short = await failure(inspectOfficeSource(sourceOf(bytes, async (o, n) => bytes.subarray(o, o + Math.max(0, n - 1)))));
    expect((short as NodeJS.ErrnoException).code).to.equal('EIO');
    const long = await failure(inspectOfficeSource(sourceOf(bytes, async (o, n) => Buffer.concat([bytes.subarray(o, o + n), Buffer.from('x')]))));
    expect((long as NodeJS.ErrnoException).code).to.equal('EIO');
  });

  it('keeps its verdict when the source fails to close', async () => {
    const bytes = fixture('lo', 'clean.docx');
    const r = await inspectOfficeSource({ ...sourceOf(bytes), close: async () => Promise.reject(new Error('close failed')) });
    expect(r.status).to.equal('clean');
  });

  it('rejects a file over the size limit without reading it', async () => {
    const bytes = fixture('lo', 'clean.docx');
    let reads = 0;
    const r = await inspectOfficeSource(
      sourceOf(bytes, async (o, n) => {
        reads++;
        return bytes.subarray(o, o + n);
      }),
      { limits: { fileSize: 100 } },
    );
    expect([r.status, r.score, kinds(r), reads]).to.deep.equal(['rejected', null, ['LIMIT/FILE_SIZE'], 0]);
    expect(r.findings[0].data).to.include({ size: bytes.length, limit: 100 });
  });

  it('applies the archive and XML limits from the options, and turns one off with Infinity', async () => {
    const bytes = fixture('lo', 'clean.docx');
    const last = async (limits: object) => kinds(await inspectOffice(bytes, { limits })).slice(-1)[0];
    expect(await last({ entries: 3 })).to.equal('LIMIT/ENTRY_COUNT');
    expect(await last({ decompressedBytes: 500 })).to.equal('LIMIT/DECOMPRESSED_SIZE');
    expect(await last({ totalDecompressedBytes: 5000 })).to.equal('LIMIT/TOTAL_DECOMPRESSED_SIZE');
    expect(await last({ xmlDepth: 3 })).to.equal('LIMIT/XML_DEPTH');
    expect(await last({ xmlTokenSize: 100 })).to.equal('LIMIT/XML_TOKEN_SIZE');
    expect((await inspectOffice(bytes, { limits: { entries: Number.POSITIVE_INFINITY, xmlTokenSize: Number.POSITIVE_INFINITY } })).status).to.equal('clean');
  });

  it('applies the entity limits to the whole package, from the options or the defaults', async () => {
    // Two parts with 30 references each to an entity whose six characters refer to another of one: 60 expansions
    // and 210 characters in each part, so only the two together pass 119 and 419.
    const subset = '<!ENTITY e "&f;abc"><!ENTITY f "x">';
    const refs = '&e;'.repeat(30);
    const bytes = await zipOf([
      ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': WORD_MAIN, 'word/part.xml': 'application/xml' })],
      ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
      ['word/document.xml', withSubset(wordXml(`<w:p><w:r><w:t>${refs}</w:t></w:r></w:p>`), subset)],
      ['word/_rels/document.xml.rels', relsXml(['rId2', `${T}customXml`, 'part.xml'])],
      ['word/part.xml', withSubset(`${XML}<root>${refs}</root>`, subset)],
    ]);
    const last = async (limits: object) => kinds(await inspectOffice(bytes, { limits })).slice(-1)[0];
    expect(await last({ xmlEntityExpansions: 119 })).to.equal('LIMIT/XML_ENTITY_EXPANSIONS');
    expect(await last({ xmlEntitySize: 419 })).to.equal('LIMIT/XML_ENTITY_SIZE');
    expect(await last({ xmlEntityDepth: 1 })).to.equal('LIMIT/XML_ENTITY_DEPTH');
    const fits = await inspectOffice(bytes, { limits: { xmlEntityExpansions: 120, xmlEntitySize: 420, xmlEntityDepth: 2 } });
    expect([fits.status, kinds(fits)]).to.deep.equal(['strippable', ['XML/INTERNAL_SUBSET', 'XML/INTERNAL_SUBSET']]);
    // The billion laughs stops at the default limit, without a score.
    const laughs = `<!ENTITY l0 "lol">${Array.from({ length: 9 }, (_, i) => `<!ENTITY l${i + 1} "${`&l${i};`.repeat(10)}">`).join('')}`;
    const bomb = await inspectOffice(await wordPackage('<w:p><w:r><w:t>&l9;</w:t></w:r></w:p>', { doctype: `<!DOCTYPE w:document [${laughs}]>` }));
    expect([bomb.status, bomb.score, bomb.risk, kinds(bomb)]).to.deep.equal(['rejected', null, 'UNKNOWN', ['XML/INTERNAL_SUBSET', 'LIMIT/XML_ENTITY_EXPANSIONS']]);
    expect(bomb.findings[1].location).to.equal('word/document.xml');
  });

  it('stops at the time limit with LIMIT/TIME, and its error carries the family code', async () => {
    const bytes = fixture('docs', 'excel.xlsm');
    const slow = sourceOf(bytes, async (o, n) => {
      await new Promise(resolve => setTimeout(resolve, 5));
      return bytes.subarray(o, o + n);
    });
    const r = await inspectOfficeSource(slow, { limits: { timeMs: 1 } });
    expect([r.status, r.score, kinds(r).slice(-1)[0]]).to.deep.equal(['rejected', null, 'LIMIT/TIME']);
    const e = new TimeLimitError('x');
    expect([e.code, e.limit]).to.deep.equal(['DEFUSE_LIMIT', 'time']);
  });

  it('rejects the file as unreadable when something unexpected fails', async () => {
    const bytes = fixture('lo', 'clean.docx');
    // A source that hands back a plain array rather than bytes.
    const odd = sourceOf(bytes, async (o, n) => Array.from(bytes.subarray(o, o + n)) as unknown as Uint8Array);
    const r = await inspectOfficeSource(odd);
    expect([r.status, r.score, kinds(r)]).to.deep.equal(['rejected', null, ['PACKAGE/UNREADABLE']]);
    expect(r.findings[0].data?.reason).to.be.a('string');
  });

  it('names the format of a package it rejects after recognizing it', async () => {
    const r = await inspectOffice(fixture('docs', 'malformed.docx'));
    expect([r.status, r.format, r.mimeType]).to.deep.equal(['rejected', 'docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']);
  });

  it('inspects in memory far smaller than what a part inflates to', async () => {
    // One XML part of about 40 MB, mostly text, deflates to a fraction of that.
    const zip = await wordPackage(`<w:p><w:r><w:t>${'x'.repeat(40 * 1024 * 1024)}</w:t></w:r></w:p>`);
    const peak = await livePeak(async () => expect((await inspectOffice(zip)).status).to.equal('clean'));
    expect(zip.length).to.be.lessThan(2 * 1024 * 1024);
    expect(peak).to.be.lessThan(16 * 1024 * 1024);
  });

  it('exports the same API to ES modules and CommonJS', async () => {
    const cjs = require(path.join(root, 'dist', 'index.js')) as Record<string, unknown>;
    const esm = (await (new Function('s', 'return import(s)') as (s: string) => Promise<Record<string, unknown>>)(path.join(root, 'dist', 'index.mjs'))) as Record<string, unknown>;
    const names = Object.keys(cjs)
      .filter(n => n !== '__esModule')
      .sort();
    expect(names).to.include.members([
      'inspectOffice',
      'inspectOfficeSource',
      'bufferSource',
      'fileSource',
      'allFindingSpecs',
      'scoreFindings',
      'COMBINATIONS',
      'DEFAULT_BANDS',
      'OfficeCategory',
      'OfficeDetail',
      'OfficeRisk',
    ]);
    for (const n of names) expect(esm[n], n).to.equal(cjs[n]);
  });
});
