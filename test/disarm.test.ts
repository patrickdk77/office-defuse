import { spawnSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { expect } from 'chai';
import { type ByteSource, bufferSource, disarmOffice, disarmOfficeSource, inspectOffice, type OfficeDisarmResult } from '../src';
import { entryNames, entryText, structureProblems } from './helpers/structure';
import { contentTypesXml, fixture, kinds, MS, must, odfPackage, of, relsXml, S, sheetXml, T, textBody, W, WORD_MAIN, wordPackage, XML, zipOf } from './helpers/util';

/** Disarms `bytes` and checks that the result has an output whose structure holds together. */
async function defused(bytes: Uint8Array, options = {}): Promise<OfficeDisarmResult & { bytes: Uint8Array }> {
  const r = await disarmOffice(bytes, options);
  const out = must(r.bytes, `output (${r.status}: ${r.before.findings.map(f => `${f.category}/${f.detail} ${JSON.stringify(f.data ?? {})}`).join('; ')})`);
  expect(await structureProblems(out)).to.deep.equal([]);
  return { ...r, bytes: out };
}

const strips = (r: { findings: Array<{ action: string }> }) => r.findings.filter(f => f.action !== 'info');

describe('disarm: ZIP packages', () => {
  it('writes every fixture with strip findings so that its output has none, and every clean one unchanged in content', async () => {
    for (const [folder, file] of [
      ['docs', 'actions.pptx'],
      ['docs', 'customui.docm'],
      ['docs', 'doctype.docx'],
      ['docs', 'events.vsdm'],
      ['docs', 'excel.xlsm'],
      ['docs', 'external.docx'],
      ['docs', 'external.odt'],
      ['docs', 'fields.docx'],
      ['docs', 'layout.odt'],
      ['docs', 'local-template.docx'],
      ['docs', 'macros.odt'],
      ['docs', 'master.odm'],
      ['docs', 'objects.docx'],
      ['docs', 'objects.odp'],
      ['docs', 'vba.docm'],
      ['lo', 'dde.ods'],
      ['lo', 'dde.odt'],
      ['lo', 'links.docx'],
      ['lo', 'links.odt'],
      ['zip', 'hidden-data.docx'],
      ['zip', 'leading-bytes.docx'],
      ['zip', 'trailing-bytes.docx'],
    ]) {
      const r = await defused(fixture(folder, file));
      expect([file, r.status, r.before.status], file).to.deep.equal([file, 'defused', 'strippable']);
      expect(strips(must(r.after, 'after')), file).to.deep.equal([]);
      expect(must(r.after, 'after').format, file).to.equal(r.before.format);
      expect(r.removed.length, file).to.equal(strips(r.before).length);
    }
  });

  it('keeps the compressed bytes of every entry a clean package does not change', async () => {
    const input = fixture('lo', 'clean.docx');
    const r = await defused(input);
    expect([r.status, r.after?.status]).to.deep.equal(['clean', 'clean']);
    expect(await entryNames(r.bytes)).to.deep.equal(await entryNames(input));
    for (const name of await entryNames(input)) expect(await entryText(r.bytes, name), name).to.equal(await entryText(input, name));
    // Python's zipfile and Info-ZIP read the output.
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-out-'));
    try {
      const file = path.join(dir, 'out.docx');
      fs.writeFileSync(file, r.bytes);
      const py = spawnSync('python3', ['-I', '-c', 'import sys, zipfile; z = zipfile.ZipFile(sys.argv[1]); assert z.testzip() is None', file], { encoding: 'utf8' });
      expect(py.status, py.stderr).to.equal(0);
      const unzip = spawnSync('unzip', ['-t', file], { encoding: 'utf8' });
      expect(unzip.status, unzip.stdout).to.equal(0);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('removes a Word field with its code and keeps its result, including fields another field builds', async () => {
    const r = await defused(fixture('docs', 'fields.docx'));
    const doc = must(await entryText(r.bytes, 'word/document.xml'), 'document');
    expect(doc).to.not.match(/DDE|QUOTE|INCLUDETEXT|MACROBUTTON|fileserver/);
    expect(doc.match(/<w:instrText/g)).to.have.length(1);
    // The page number field has no finding and stays whole.
    expect(doc).to.include('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
    expect(doc.match(/>(result|simple|computed)</g)).to.deep.equal(['>result<', '>simple<', '>computed<', '>result<', '>result<', '>result<', '>result<']);
  });

  it('drops a relationship a strip finding names, and every element that names it, keeping the text of a hyperlink', async () => {
    const body = '<w:p><w:hyperlink r:id="rIdBad"><w:r><w:t>bad link</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdGood"><w:r><w:t>good link</w:t></w:r></w:hyperlink></w:p>';
    const input = await wordPackage(body, {
      rels: [
        ['rIdBad', `${T}hyperlink`, 'http://203.0.113.7/', true],
        ['rIdGood', `${T}hyperlink`, 'https://example.com/', true],
      ],
    });
    const r = await defused(input);
    expect(kinds(r.before)).to.deep.equal(['LINK/IP_HOST', 'LINK/SAFE']);
    const doc = must(await entryText(r.bytes, 'word/document.xml'), 'document');
    expect(doc).to.include('<w:p><w:r><w:t>bad link</w:t></w:r><w:hyperlink r:id="rIdGood"><w:r><w:t>good link</w:t></w:r></w:hyperlink></w:p>');
    expect(await entryText(r.bytes, 'word/_rels/document.xml.rels')).to.not.include('203.0.113.7');
  });

  it('drops a part with its relationships, its content type and the parts only it reaches', async () => {
    const r = await defused(fixture('docs', 'vba.docm'));
    const names = await entryNames(r.bytes);
    expect(names.filter(n => /vba/i.test(n))).to.deep.equal([]);
    expect(await entryText(r.bytes, '[Content_Types].xml')).to.not.match(/vbaProject|vbaData/i);
    // The main part keeps its content type, so the file stays a .docm.
    expect(must(r.after, 'after').format).to.equal('docm');
  });

  it('turns a macro sheet into an empty worksheet of the same name and drops its Auto_Open name', async () => {
    const r = await defused(fixture('docs', 'excel.xlsm'));
    expect(await entryText(r.bytes, 'xl/macrosheets/sheet1.xml')).to.equal(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    );
    const workbook = must(await entryText(r.bytes, 'xl/workbook.xml'), 'workbook');
    expect(workbook).to.include('<sheet name="Macro1"');
    expect(workbook).to.not.include('Auto_Open');
    expect(await entryText(r.bytes, 'xl/_rels/workbook.xml.rels')).to.include(`Type="${T}worksheet" Target="macrosheets/sheet1.xml"`);
    // A macro sheet typed by a default rather than an override gets an override.
    const byDefault = await zipOf([
      ['[Content_Types].xml', contentTypesXml({ 'xl/workbook.xml': 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' }, { mac: 'application/vnd.ms-excel.macrosheet+xml' })],
      ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'xl/workbook.xml'])],
      ['xl/workbook.xml', `${XML}<workbook xmlns="${S}" xmlns:r="${T.slice(0, -1)}"><sheets><sheet name="M" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ['xl/_rels/workbook.xml.rels', relsXml(['rId1', `${MS}xlMacrosheet`, 'm.mac'])],
      ['xl/m.mac', `${XML}<worksheet xmlns="${S}"><sheetData/></worksheet>`],
    ]);
    const d = await defused(byDefault);
    expect(await entryText(d.bytes, '[Content_Types].xml')).to.include(
      '<Override PartName="/xl/m.mac" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    );
  });

  it('removes callbacks from a custom ribbon, macro attributes from drawings and actions from slides', async () => {
    const ribbon = await defused(fixture('docs', 'customui.docm'));
    const ui = must(await entryText(ribbon.bytes, 'customUI/customUI14.xml'), 'customUI');
    expect(ui).to.not.match(/onLoad|onAction/);
    expect(ui).to.include('<button id="b1" label="Go"/>');
    const sheet = await defused(fixture('docs', 'excel.xlsm'));
    expect(await entryText(sheet.bytes, 'xl/drawings/drawing1.xml')).to.include('<xdr:sp textlink="">');
    expect(await entryText(sheet.bytes, 'xl/drawings/vmlDrawing1.vml')).to.not.include('FmlaMacro');
    const slides = await defused(fixture('docs', 'actions.pptx'));
    // A web link on the same slide stays.
    expect((await entryText(slides.bytes, 'ppt/slides/slide1.xml'))?.match(/<a:hlink[^>]*>/g)).to.deep.equal(['<a:hlinkClick r:id="rId1"/>']);
  });

  it('drops an emptied external link part with the reference to it, and the list that held only that reference', async () => {
    const workbook = `${XML}<workbook xmlns="${S}" xmlns:r="${T.slice(0, -1)}"><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>`;
    const input = await zipOf([
      [
        '[Content_Types].xml',
        contentTypesXml({
          'xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          'xl/sheet1.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
          'xl/link.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
        }),
      ],
      ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'xl/workbook.xml'])],
      ['xl/workbook.xml', workbook],
      ['xl/_rels/workbook.xml.rels', relsXml(['rId1', `${T}worksheet`, 'sheet1.xml'], ['rId2', `${T}externalLink`, 'link.xml'])],
      ['xl/sheet1.xml', sheetXml('<c r="A1"><v>1</v></c>')],
      ['xl/link.xml', `${XML}<externalLink xmlns="${S}"><ddeLink ddeService="cmd" ddeTopic="/c calc"/></externalLink>`],
    ]);
    const r = await defused(input);
    expect(await entryNames(r.bytes)).to.not.include('xl/link.xml');
    expect(await entryText(r.bytes, 'xl/workbook.xml')).to.include('</sheets></workbook>');
  });

  it('expands entities and writes declared attributes out when it drops an internal subset, and drops references to external entities', async () => {
    const subset = '<!ENTITY dde "DDE&#65;UTO"><!ENTITY run "<w:r><w:t>&dde;</w:t></w:r>"><!ENTITY ext SYSTEM "https://example.invalid/x"><!ATTLIST w:p w:rsidR CDATA "00AB">';
    const body = '<w:p>&run;<w:r><w:t>&ext;</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">a&amp;b &lt;&dde;&gt;</w:t></w:r></w:p>';
    const input = await wordPackage(body, { doctype: `<!DOCTYPE w:document [${subset}]>` });
    const r = await defused(input);
    expect(kinds(r.before)).to.deep.equal(['XML/INTERNAL_SUBSET', 'XML/EXTERNAL_ENTITY']);
    const doc = must(await entryText(r.bytes, 'word/document.xml'), 'document');
    expect(doc).to.not.include('DOCTYPE');
    expect(doc).to.include('<w:p w:rsidR="00AB"><w:r><w:t>DDEAUTO</w:t></w:r><w:r><w:t></w:t></w:r></w:p>');
    expect(doc).to.include('<w:p w:rsidR="00AB"><w:r><w:t xml:space="preserve">a&amp;b &lt;DDEAUTO&gt;</w:t></w:r></w:p>');
    // Inspected again, the part reads as it read before, entity for entity.
    expect(kinds(must(r.after, 'after'))).to.deep.equal([]);
  });

  it('expands an entity reference in an attribute value, and keeps a document type with no internal subset', async () => {
    const part = `${XML}<!DOCTYPE r [<!ENTITY h "https://example.com/?a=1&#38;#38;b=2">]><r xmlns="urn:x" a="&h;" b="&amp;"/>`;
    const r = await defused(await wordPackage('<w:p/>', { parts: { 'word/custom.xml': part }, overrides: { 'word/custom.xml': 'application/xml' }, rels: [['rIdX', `${T}customXml`, 'custom.xml']] }));
    expect(await entryText(r.bytes, 'word/custom.xml')).to.equal(`${XML}<r xmlns="urn:x" a="https://example.com/?a=1&amp;b=2" b="&amp;"/>`);
    const plain = `${XML}<!DOCTYPE r SYSTEM "r.dtd"><r xmlns="urn:x"/>`;
    const kept = await defused(
      await wordPackage('<w:p/>', { parts: { 'word/custom.xml': plain }, overrides: { 'word/custom.xml': 'application/xml' }, rels: [['rIdX', `${T}customXml`, 'custom.xml']] }),
    );
    expect([kept.status, await entryText(kept.bytes, 'word/custom.xml')]).to.deep.equal(['clean', plain]);
  });

  it('leaves out an element that came from entity text by writing that reference out without it', async () => {
    const input = await wordPackage('<w:p>&f;</w:p>', { doctype: '<!DOCTYPE w:document [<!ENTITY f "<w:fldSimple w:instr=\'DDEAUTO a b\'><w:r><w:t>kept</w:t></w:r></w:fldSimple>">]>' });
    const r = await disarmOffice(input, { actionOverrides: [{ category: 'XML' as never, detail: 'INTERNAL_SUBSET' as never, action: 'info' }] });
    const doc = must(await entryText(must(r.bytes, 'output'), 'word/document.xml'), 'document');
    // The subset stays, under the override, and the reference with the field in it is written out without the field.
    expect(doc).to.include('<!DOCTYPE w:document');
    expect(doc).to.include('<w:p><w:r><w:t>kept</w:t></w:r></w:p>');
  });

  it('refuses an output that would need a subset the overrides keep and the rule strips', async () => {
    const input = await wordPackage('<w:p><w:r><w:t>&ext;</w:t></w:r></w:p>', { doctype: '<!DOCTYPE w:document [<!ENTITY ext SYSTEM "https://example.invalid/x">]>' });
    const r = await disarmOffice(input, { actionOverrides: [{ category: 'XML' as never, detail: 'EXTERNAL_ENTITY' as never, action: 'info' }] });
    expect([r.status, r.bytes, of(r.before, 'PROCESSING/VERIFICATION_FAILED')[0]?.data?.reasons]).to.deep.equal(['rejected', undefined, 'XML/INTERNAL_SUBSET']);
  });

  it('writes OpenDocument packages with the mimetype first and the manifest in step with the entries', async () => {
    const macros = await defused(fixture('docs', 'macros.odt'));
    const names = await entryNames(macros.bytes);
    expect(names[0]).to.equal('mimetype');
    expect(names.filter(n => /^(Basic|Scripts|Dialogs|Configurations2)\//.test(n))).to.deep.equal([]);
    expect(await entryText(macros.bytes, 'content.xml')).to.not.include('script:event-listener');
    const objects = await defused(fixture('docs', 'objects.odp'));
    const content = must(await entryText(objects.bytes, 'content.xml'), 'content');
    // The object goes and its preview image stays.
    expect(content).to.not.match(/object-ole|draw:applet|draw:plugin|floating-frame|presentation:event-listener/);
    expect(content).to.include('xlink:href="./ObjectReplacements/Object 1"');
    expect(await entryNames(objects.bytes)).to.not.include('Object 1');
    const layout = await defused(fixture('docs', 'layout.odt'));
    expect((await entryNames(layout.bytes))[0]).to.equal('mimetype');
  });

  it('writes a mimetype entry and a manifest for an OpenDocument package that lacks them', async () => {
    const r = await defused(await odfPackage(textBody('<text:p>x</text:p>'), { mimetype: null }));
    expect(await entryText(r.bytes, 'mimetype')).to.equal('application/vnd.oasis.opendocument.text');
    const bare = await defused(await odfPackage(textBody('<text:p>x</text:p>'), { manifestType: null }));
    expect(await entryText(bare.bytes, 'META-INF/manifest.xml')).to.include('manifest:full-path="content.xml"');
  });

  it('removes OpenDocument links whole or by their attribute, as the element needs', async () => {
    const body = textBody(
      '<text:p><text:a xlink:href="http://203.0.113.5/x">kept text</text:a></text:p><form:form><form:button form:name="b" xlink:href="javascript:x()"/></form:form><table:table><table:table-row><table:table-cell table:formula="of:=DDE(&quot;cmd&quot;;&quot;x&quot;;&quot;y&quot;)" office:value="3"/></table:table-row></table:table>',
    );
    const r = await defused(await odfPackage(body));
    const content = must(await entryText(r.bytes, 'content.xml'), 'content');
    expect(content).to.include('<text:p>kept text</text:p>');
    expect(content).to.include('<form:button form:name="b"/>');
    expect(content).to.include('<table:table-cell office:value="3"/>');
  });

  it('decrypts an encrypted package into its plain form, with each password of a list tried in turn', async () => {
    const velvet = await defused(fixture('cases', 'ole-velvetsweatshop-package.xlsx'));
    expect([velvet.status, velvet.before.format, must(velvet.after, 'after').format, kinds(must(velvet.after, 'after'))]).to.deep.equal(['defused', 'xlsx', 'xlsx', []]);
    expect(Buffer.from(velvet.bytes.subarray(0, 2)).toString('latin1')).to.equal('PK');
    const withList = await defused(fixture('ole', 'password.docx'), { password: ['wrong', 'also wrong', 'inert-password'] });
    expect(kinds(withList.before)).to.deep.equal(['ENCRYPTED/USER_PASSWORD', 'ENCRYPTED/AES_256']);
  });

  it('rejects a file it cannot rewrite clean and writes nothing to the sink', async () => {
    let written = 0;
    let closed = false;
    const sink = {
      write: async (c: Uint8Array) => {
        written += c.length;
      },
      close: async () => {
        closed = true;
      },
    };
    const r = await disarmOfficeSource(bufferSource(fixture('docs', 'malformed.docx')), sink);
    expect([r.status, r.after, r.removed, written, closed]).to.deep.equal(['rejected', undefined, [], 0, false]);
  });

  it('reads the upload once, so a source that changes after the read cannot change the output', async () => {
    const bytes = Buffer.from(await wordPackage('<w:p><w:r><w:t>first</w:t></w:r></w:p>'));
    let reads = 0;
    const source: ByteSource = {
      size: async () => bytes.length,
      read: async (offset, length) => {
        reads++;
        const b = Buffer.from(bytes.subarray(offset, offset + length));
        if (reads > 1000) b.fill(0);
        return b;
      },
    };
    const chunks: Buffer[] = [];
    const r = await disarmOfficeSource(source, { write: async c => void chunks.push(Buffer.from(c)), close: async () => undefined });
    expect(r.status).to.equal('clean');
    expect(reads).to.equal(1);
    expect(await entryText(Buffer.concat(chunks), 'word/document.xml')).to.include('first');
  });

  it('aborts the sink when copying the output to it fails', async () => {
    let aborted: unknown;
    const failure = new Error('disk full');
    const sink = {
      write: async () => Promise.reject(failure),
      close: async () => undefined,
      abort: async (e: unknown) => {
        aborted = e;
      },
    };
    let thrown: unknown;
    try {
      await disarmOfficeSource(bufferSource(fixture('lo', 'clean.docx')), sink);
    } catch (e) {
      thrown = e;
    }
    expect([thrown, aborted]).to.deep.equal([failure, failure]);
  });

  it('gives the same verdict as inspect for the upload', async () => {
    for (const file of ['excel.xlsm', 'objects.docx', 'macros.odt']) {
      const bytes = fixture('docs', file);
      expect((await disarmOffice(bytes)).before, file).to.deep.equal(await inspectOffice(bytes));
    }
  });

  it('keeps a part whose finding an override turns to info', async () => {
    const r = await defused(fixture('docs', 'vba.docm'), { actionOverrides: [{ category: 'MACRO' as never, detail: 'VBA_PROJECT' as never, action: 'info' }] });
    expect([r.status, (await entryNames(r.bytes)).filter(n => /vbaProject\.bin$/i.test(n))]).to.deep.equal(['clean', ['word/vbaProject.bin']]);
  });

  it('rewrites a part held in UTF-16 in UTF-16', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-16"?><w:document xmlns:w="${W}" xmlns:r="${T.slice(0, -1)}"><w:body><w:p><w:fldSimple w:instr="DDEAUTO a b"><w:r><w:t>été \u{1f600}</w:t></w:r></w:fldSimple></w:p></w:body></w:document>`;
    const bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, 'utf16le')]);
    const input = await zipOf([
      ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': WORD_MAIN })],
      ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
      ['word/document.xml', bytes],
    ]);
    const r = await defused(input);
    const raw = await rawEntry(r.bytes, 'word/document.xml');
    expect([...raw.subarray(0, 2)]).to.deep.equal([0xff, 0xfe]);
    expect(raw.subarray(2).toString('utf16le')).to.equal(xml.replace('<w:fldSimple w:instr="DDEAUTO a b">', '').replace('</w:fldSimple>', ''));
  });
});

async function rawEntry(bytes: Uint8Array, name: string): Promise<Buffer> {
  const { ZipArchive } = await import('../src/zip');
  const zip = await ZipArchive.open(bufferSource(bytes), { entries: 1000, decompressedBytes: 1 << 30, totalDecompressedBytes: 1 << 30 });
  const chunks: Buffer[] = [];
  for await (const c of zip.read(must(zip.get(name), name))) chunks.push(Buffer.from(c));
  return Buffer.concat(chunks);
}

describe('disarm: entities and overrides', () => {
  const info = (category: string, detail?: string) => ({ category: category as never, ...(detail ? { detail: detail as never } : {}), action: 'info' as const });

  it('writes a reference to an external entity back, or leaves it out, inside entity text it writes out', async () => {
    const subset = '<!ENTITY ext SYSTEM "https://example.invalid/x"><!ENTITY f "<w:fldSimple w:instr=\'DDEAUTO a b\'><w:r><w:t>kept &ext;</w:t></w:r></w:fldSimple>">';
    const input = await wordPackage('<w:p>&f;</w:p>', { doctype: `<!DOCTYPE w:document [${subset}]>` });
    const kept = await disarmOffice(input, { actionOverrides: [info('XML')] });
    expect(await entryText(must(kept.bytes, 'output'), 'word/document.xml')).to.include('<w:p><w:r><w:t>kept &ext;</w:t></w:r></w:p>');
    const dropped = await disarmOffice(input, { actionOverrides: [info('XML', 'INTERNAL_SUBSET')] });
    expect(await entryText(must(dropped.bytes, 'output'), 'word/document.xml')).to.include('<w:p><w:r><w:t>kept </w:t></w:r></w:p>');
  });

  it('removes a document type declaration without a subset when an override strips it', async () => {
    const part = `${XML}<!DOCTYPE r SYSTEM "r.dtd"><r xmlns="urn:x"/>`;
    const input = await wordPackage('<w:p/>', { parts: { 'word/custom.xml': part }, overrides: { 'word/custom.xml': 'application/xml' }, rels: [['rIdX', `${T}customXml`, 'custom.xml']] });
    const r = await defused(input, { actionOverrides: [{ category: 'XML' as never, detail: 'DOCTYPE' as never, action: 'strip' }] });
    expect(await entryText(r.bytes, 'word/custom.xml')).to.equal(`${XML}<r xmlns="urn:x"/>`);
  });

  it('takes an entity reference inside an element it removes with that element', async () => {
    const body = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>&dde; a b</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const r = await disarmOffice(await wordPackage(body, { doctype: '<!DOCTYPE w:document [<!ENTITY dde "DDEAUTO">]>' }), { actionOverrides: [info('XML', 'INTERNAL_SUBSET')] });
    const doc = must(await entryText(must(r.bytes, 'output'), 'word/document.xml'), 'document');
    expect(doc).to.include('<!DOCTYPE');
    expect(doc).to.not.include('&dde;');
    expect(doc).to.not.include('instrText');
  });

  it('merges two edits of one element, and lets removal cover the rest', async () => {
    const ui = `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="A" loadImage="B"><ribbon/></customUI>`;
    const input = await wordPackage('<w:p/>', {
      parts: { 'customUI/customUI14.xml': ui },
      overrides: { 'customUI/customUI14.xml': 'application/xml' },
      packageRels: [['rIdUi', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility', 'customUI/customUI14.xml']],
    });
    const r = await defused(input);
    expect(await entryText(r.bytes, 'customUI/customUI14.xml')).to.equal('<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon/></customUI>');
  });
});
