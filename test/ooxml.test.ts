import { expect } from 'chai';
import { inspectOffice } from '../src';
import { contentTypesXml, fixture, has, kinds, MS, must, OLE, of, relsXml, S, sheetXml, T, W, WORD_MAIN, withSubset, wordPackage, wordWithPart, wordXml, XML, zipOf } from './helpers/util';

const fld = (instr: string, result = 'r') =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const vmlPart = (inner: string) => `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">${inner}</xml>`;
const PAD = ' '.repeat(5000);

/** An .xlsb package with one external link part holding `link`, the bytes of its records. */
function xlsb(link: Uint8Array): Promise<Buffer> {
  return zipOf([
    ['[Content_Types].xml', contentTypesXml({ 'xl/externalLinks/externalLink1.bin': 'application/vnd.ms-excel.externalLink' }, { bin: 'application/vnd.ms-excel.sheet.binary.macroEnabled.main' })],
    ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'xl/workbook.bin'])],
    ['xl/workbook.bin', Buffer.from([0x83, 0x01, 0x00, 0x84, 0x01, 0x00])],
    ['xl/_rels/workbook.bin.rels', relsXml(['rId1', `${T}externalLink`, 'externalLinks/externalLink1.bin'])],
    ['xl/externalLinks/externalLink1.bin', link],
  ]);
}
const wide = (s: string) => Buffer.concat([Buffer.from(Uint32Array.of(s.length).buffer), Buffer.from(s, 'utf16le')]);
/** BrtBeginSupBook with link type `type` and two names, as one record. */
const supBook = (type: number, a: string, b: string) => {
  const body = Buffer.concat([Buffer.from([type, 0]), wide(a), wide(b)]);
  return Buffer.concat([Buffer.from([0xe8, 0x02, body.length]), body]);
};

describe('OOXML rules', () => {
  it('finds DDE, QUOTE-built, include, macro button and hyperlink fields', async () => {
    const r = await inspectOffice(fixture('docs', 'fields.docx'));
    expect(kinds(r).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['DDE/FIELD', 'DDE/FIELD', 'FIELD/COMPUTED', 'EXTERNAL/INCLUDE', 'FIELD/MACROBUTTON', 'LINK/NETWORK_PATH']);
    expect(of(r, 'DDE/FIELD')[0].data?.code).to.match(/^DDEAUTO /);
    expect(of(r, 'EXTERNAL/INCLUDE')[0].data?.target).to.equal('http://203.0.113.9/payload.docx');
    expect(of(r, 'LINK/NETWORK_PATH')[0].data?.target).to.equal('\\\\fileserver\\share\\notes.docx');
    expect([r.status, r.risk]).to.deep.equal(['strippable', 'CRITICAL']);
  });

  it('reads field codes split over runs, written in any case, hidden in fallbacks or deleted text', async () => {
    const body = [
      `<w:p>${fld('dd')}</w:p>`,
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> dD</w:instrText></w:r><w:r><w:instrText>eAuTo cmd x</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback><w:p><w:fldSimple w:instr="DDE a b"/></w:p></mc:Fallback></mc:AlternateContent>',
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText>DDEAUTO c d</w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      // A field the part never closes still runs.
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MACROBUTTON M x</w:instrText></w:r></w:p>',
    ].join('');
    expect(kinds(await inspectOffice(await wordPackage(body)))).to.deep.equal(['DDE/FIELD', 'DDE/FIELD', 'DDE/FIELD', 'FIELD/MACROBUTTON']);
  });

  it('finds a field name after any amount of white space split over runs', async () => {
    const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${PAD}</w:instrText></w:r><w:r><w:instrText>DDEAUTO c d</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const r = await inspectOffice(await wordPackage(body));
    expect(kinds(r)).to.deep.equal(['DDE/FIELD']);
    expect(r.findings[0].data).to.deep.equal({ code: 'DDEAUTO c d' });
  });

  it('reads Word elements under any prefix bound to the Word namespace, and leaves other namespaces alone', async () => {
    const body = `<q:p xmlns:q="${W}"><q:fldSimple q:instr="DDEAUTO x y"/></q:p><o:p xmlns:o="urn:other"><o:fldSimple o:instr="DDEAUTO x y"/></o:p>`;
    expect(kinds(await inspectOffice(await wordPackage(body)))).to.deep.equal(['DDE/FIELD']);
  });

  it('reports a field built by another field, and fields with nothing to report as nothing', async () => {
    const nested = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>${fld('QUOTE "DDE"')}<w:r><w:instrText> a b</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    expect(kinds(await inspectOffice(await wordPackage(nested)))).to.deep.equal(['FIELD/COMPUTED']);
    const quiet = '<w:p><w:fldSimple/><w:fldSimple w:instr="  "/><w:fldSimple w:instr="PAGE"/><w:r><w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    expect(kinds(await inspectOffice(await wordPackage(quiet)))).to.deep.equal([]);
  });

  it('reports macro buttons with or without a macro name, and includes with or without a target', async () => {
    const body = `<w:p>${fld('MACROBUTTON')}${fld('INCLUDETEXT')}${fld('INCLUDETEXT \\* MERGEFORMAT')}${fld('IMPORT "https://example.com/a.png"')}</w:p>`;
    const r = await inspectOffice(await wordPackage(body));
    expect(r.findings.map(f => [f.category, f.detail, f.data?.macro, f.data?.target])).to.deep.equal([
      ['FIELD', 'MACROBUTTON', '', undefined],
      ['EXTERNAL', 'INCLUDE', undefined, undefined],
      ['EXTERNAL', 'INCLUDE', undefined, undefined],
      ['EXTERNAL', 'INCLUDE', undefined, 'https://example.com/a.png'],
    ]);
  });

  it('reports includes from the machine as local files, and HYPERLINK fields to places in the document not at all', async () => {
    const body = `<w:p>${fld('INCLUDEPICTURE "C:\\\\Pictures\\\\logo.png" \\d')}${fld('LINK Excel.Sheet.12 "\\\\\\\\server\\\\x.xlsx" "Sheet1!R1C1"')}${fld('HYPERLINK \\l "Top"')}${fld('HYPERLINK "#Top"')}${fld('HYPERLINK "https://example.com/"')}</w:p>`;
    const r = await inspectOffice(await wordPackage(body));
    expect(kinds(r)).to.deep.equal(['EXTERNAL/LOCAL_FILE', 'EXTERNAL/INCLUDE', 'LINK/SAFE']);
    expect(r.findings.map(f => f.data?.target)).to.deep.equal(['C:\\Pictures\\logo.png', '\\\\server\\x.xlsx', 'https://example.com/']);
    expect(r.findings[0].data?.kind).to.equal('field');
  });

  it('reports relationships that leave the package by their type, and local ones as local files', async () => {
    const r = await inspectOffice(fixture('docs', 'external.docx'));
    expect(kinds(r).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['EXTERNAL/OBJECT', 'EXTERNAL/IMAGE', 'EXTERNAL/FRAME', 'EXTERNAL/INCLUDE', 'EXTERNAL/OTHER', 'EXTERNAL/TEMPLATE']);
    const template = must(of(r, 'EXTERNAL/TEMPLATE')[0], 'template');
    expect(template.location).to.equal('word/_rels/settings.xml.rels, rIdTpl');
    expect(template.description).to.equal('The document loads its template from https://templates.example.net/report.dotm');
    const local = await inspectOffice(fixture('docs', 'local-template.docx'));
    expect(kinds(local).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['EXTERNAL/LOCAL_FILE']);
    expect(local.risk).to.equal('LOW');
  });

  it('reports workbook links of every Excel path type, and a package linked from outside as an object', async () => {
    const types = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath', `${MS}xlExternalLinkPath/xlStartup`, `${T}package`];
    const r = await inspectOffice(await wordPackage('<w:p/>', { rels: types.map((t, i) => [`rId${i + 2}`, t, `https://example.com/${i}.xlsx`, true]) }));
    expect(kinds(r)).to.deep.equal(['EXTERNAL/WORKBOOK', 'EXTERNAL/WORKBOOK', 'EXTERNAL/OBJECT']);
  });

  it('treats a target with a scheme as external whatever TargetMode says, and reads TargetMode in any case', async () => {
    const r = await inspectOffice(
      await wordPackage('<w:p/>', {
        rels: [
          ['rId5', `${T}attachedTemplate`, 'https://evil.example/t.dotm'],
          ['rId6', `${T}image`, '\\\\host\\share\\a.png'],
        ],
      }),
    );
    expect(kinds(r)).to.deep.equal(['EXTERNAL/TEMPLATE', 'EXTERNAL/IMAGE']);
    const rels = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${T}frame" Target="page.html" TargetMode="external"/></Relationships>`;
    const lower = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/_rels/document.xml.rels': rels } }));
    expect(kinds(lower)).to.deep.equal(['EXTERNAL/LOCAL_FILE']);
    expect(lower.findings[0].data?.kind).to.equal('frame');
  });

  it('reads a percent-encoded scheme, the script moniker trick of CVE-2017-0199, as reaching outside', async () => {
    const target = '%73cript:%68ttp://example.invalid/marker';
    for (const external of [true, false]) {
      const r = await inspectOffice(await wordPackage('<w:p/>', { rels: [['rId2', `${T}oleObject`, target, external]] }));
      expect([kinds(r), r.risk], String(external)).to.deep.equal([['EXTERNAL/OBJECT'], 'HIGH']);
    }
  });

  it('finds VBA projects, embedded objects, packages, stray OLE files and ActiveX controls by type and relationship', async () => {
    const vba = await inspectOffice(fixture('docs', 'vba.docm'));
    expect(kinds(vba).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['MACRO/VBA_PROJECT']);
    expect(of(vba, 'MACRO/VBA_PROJECT')[0].location).to.equal('word/vbaProject.bin');
    const objects = await inspectOffice(fixture('docs', 'objects.docx'));
    const found = objects.findings.filter(f => f.category !== 'LINK');
    expect(found.map(f => `${f.category}/${f.detail} ${f.location}`)).to.deep.equal([
      'EMBEDDED/OLE_OBJECT word/embeddings/oleObject1.bin',
      'EMBEDDED/PACKAGE word/embeddings/Book1.xlsx',
      'EMBEDDED/OLE_FILE word/media/image1.png',
      'CONTROL/ACTIVEX word/activeX/activeX1.xml',
    ]);
  });

  it('finds a VBA project, a macro sheet, an OLE object and a control binary that only a relationship declares', async () => {
    const r = await inspectOffice(
      await wordPackage('<w:p/>', {
        defaults: { bin: 'application/octet-stream' },
        rels: [
          ['rId1', `${MS}vbaProject`, 'a.bin'],
          ['rId2', `${MS}xlMacrosheet`, 'b.xml'],
          ['rId3', `${T}oleObject`, 'c.bin'],
          ['rId4', `${MS}activeXControlBinary`, 'd.bin'],
        ],
        parts: { 'word/a.bin': OLE, 'word/b.xml': '<x/>', 'word/c.bin': OLE, 'word/d.bin': OLE },
      }),
    );
    expect(kinds(r)).to.deep.equal(['MACRO/VBA_PROJECT', 'MACRO/XLM_MACROSHEET', 'EMBEDDED/OLE_OBJECT', 'CONTROL/ACTIVEX']);
  });

  it('reads the relationships of a part that is not there, and reports them as unreferenced', async () => {
    const r = await inspectOffice(
      await wordPackage('<w:p/>', {
        defaults: { bin: 'application/octet-stream' },
        rels: [['rId1', `${T}customXml`, 'a.bin']],
        parts: { 'word/a.bin': OLE, 'word/_rels/missing.xml.rels': relsXml(['rId1', `${MS}vbaProject`, 'a.bin'], ['rId2', `${T}hyperlink`, 'https://example.com/', true]) },
      }),
    );
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.location}`)).to.deep.equal([
      'LINK/SAFE word/_rels/missing.xml.rels, rId2',
      'MACRO/VBA_PROJECT word/a.bin',
      'PACKAGE/UNREFERENCED_PART word/_rels/missing.xml.rels',
    ]);
  });

  it('reports ribbon callbacks', async () => {
    const r = await inspectOffice(fixture('docs', 'customui.docm'));
    expect(kinds(r).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['UI/RIBBON', 'UI/CALLBACK', 'UI/CALLBACK']);
    expect(of(r, 'UI/CALLBACK').map(f => f.data?.callback)).to.deep.equal(['onLoad=RibbonLoaded', 'onAction=DoIt']);
    expect(of(r, 'UI/RIBBON')[0].action).to.equal('info');
  });

  it('reports PowerPoint actions that run programs, macros and object verbs, naming the relationship', async () => {
    const r = await inspectOffice(fixture('docs', 'actions.pptx'));
    expect(kinds(r).filter(k => k !== 'LINK/SAFE')).to.deep.equal(['ACTION/RUN_PROGRAM', 'MACRO/LINK', 'LINK/FILE_URL']);
    expect(of(r, 'ACTION/RUN_PROGRAM')[0].data).to.deep.equal({ trigger: 'hlinkClick', action: 'ppaction://program', relationship: 'rIdRun' });
    const verb = await inspectOffice(
      await wordPackage(
        '<w:p><a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" action="ppaction://ole?verb=0"/><a:hlinkHover action="ppaction://hlinkshowjump?jump=nextslide" xmlns:a="urn:a"/><a:hlinkClick xmlns:a="urn:a"/></w:p>',
      ),
    );
    expect(kinds(verb)).to.deep.equal(['ACTION/OLE_VERB']);
    expect(verb.findings[0].data).to.deep.equal({ trigger: 'hlinkClick', action: 'ppaction://ole?verb=0' });
  });

  it('reports Excel DDE, external workbooks, macro sheets and shapes that run macros', async () => {
    const r = await inspectOffice(fixture('docs', 'excel.xlsm'));
    expect(r.format).to.equal('xlsm');
    expect(kinds(r)).to.deep.equal(['MACRO/EVENT_ON_OPEN', 'DDE/FORMULA', 'DDE/LINK', 'EXTERNAL/WORKBOOK', 'MACRO/LINK', 'MACRO/LINK', 'MACRO/XLM_MACROSHEET']);
    expect(of(r, 'MACRO/EVENT_ON_OPEN').map(f => f.data)).to.deep.equal([{ event: 'Auto_Open', name: '_xlnm.Auto_Open' }]);
    expect(of(r, 'DDE/FORMULA').map(f => f.data?.formula)).to.deep.equal(["cmd|' /c calc'!A0"]);
    expect(of(r, 'MACRO/LINK').map(f => f.data?.macro)).to.deep.equal(['[0]!Module1.Run', '[0]!Module1.Click']);
  });

  it('reports defined names Excel runs as macros when the workbook opens or closes, and skips names without a name', async () => {
    const r = await inspectOffice(await wordWithPart(sheetXml('', '<definedNames><definedName>A1</definedName><definedName name="auto_close_x">A1</definedName></definedNames>')));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data])).to.deep.equal([['MACRO/EVENT', { event: 'Auto_Close', name: 'auto_close_x' }]]);
  });

  it('finds DDE in a formula of any length, outside quotes, and in defined names', async () => {
    const cells = [
      `<c r="A1"><f>${PAD}<![CDATA[ ]]>cmd|' /c calc'!A0</f></c>`,
      `<c r="A2"><f>cmd|'x'!A1&amp;"${'y'.repeat(300)}"<![CDATA[&1]]></f></c>`,
      `<c r="A3"><f>"a|b"&amp;'x|y'!A1<![CDATA[&"say ""|"" here"]]></f></c>`,
      '<c r="A4"><f>SUM(A1:A2)</f></c>',
    ].join('');
    const r = await inspectOffice(await wordWithPart(sheetXml(cells, '<definedNames><definedName name="d">cmd|\'/c x\'!A1</definedName></definedNames>')));
    expect(kinds(r)).to.deep.equal(['DDE/FORMULA', 'DDE/FORMULA', 'DDE/FORMULA']);
    const shown = r.findings.map(f => String(f.data?.formula));
    expect([shown[0], shown[1].length, shown[1].endsWith('...'), shown[2]]).to.deep.equal(["cmd|' /c calc'!A0", 200, true, "cmd|'/c x'!A1"]);
  });

  it('reports an Excel DDE link part with or without its names', async () => {
    const r = await inspectOffice(await wordWithPart(`<externalLink xmlns="${S}"><ddeLink ddeService="cmd" ddeTopic="/c calc"/><ddeLink/></externalLink>`));
    expect(r.findings.map(f => f.data)).to.deep.equal([
      { application: 'cmd', topic: '/c calc' },
      { application: '', topic: '' },
    ]);
  });

  it('finds the DDE link record in an .xlsb external link part, and leaves workbook and OLE link records alone', async () => {
    const r = await inspectOffice(await xlsb(supBook(1, 'inert-app', 'inert topic')));
    expect([r.format, kinds(r), r.findings[0].data]).to.deep.equal(['xlsb', ['DDE/LINK'], { application: 'inert-app', topic: 'inert topic' }]);
    for (const type of [0, 2]) expect(kinds(await inspectOffice(await xlsb(supBook(type, 'a', 'b')))), String(type)).to.deep.equal([]);
  });

  it('reads shapes that run macros in the VML dialect Office writes, whatever pads the macro name', async () => {
    const shape = `<![if gte mso 9]><v:shape><v:textbox><div><b>Multi<br>Line</b></div></v:textbox><x:ClientData><x:FmlaMacro>${PAD}<![CDATA[ ]]>Macro1</x:FmlaMacro></x:ClientData></v:shape><![endif]><x:ClientData><x:FmlaMacro>  </x:FmlaMacro></x:ClientData>`;
    const r = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/v.vml': vmlPart(shape) }, defaults: { vml: VML }, rels: [['rIdV', `${T}vmlDrawing`, 'v.vml']] }));
    expect(kinds(r)).to.deep.equal(['MACRO/LINK']);
    expect(r.findings[0].data?.macro).to.equal('Macro1');
    // A VML part with a NUL is refused even in the dialect.
    const nul = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/v.vml': vmlPart('\u0000') }, defaults: { vml: VML }, rels: [['rIdV', `${T}vmlDrawing`, 'v.vml']] }));
    expect(kinds(nul)).to.deep.equal(['XML/MALFORMED']);
  });

  it('reports Visio event cells that run add-ons, cells that call code, and Visio hyperlinks', async () => {
    const r = await inspectOffice(fixture('docs', 'events.vsdm'));
    expect(r.format).to.equal('vsdm');
    expect(kinds(r)).to.deep.equal(['LINK/NETWORK_PATH', 'MACRO/EVENT', 'MACRO/VBA_PROJECT']);
    expect(of(r, 'MACRO/EVENT')[0].description).to.equal('A macro or script runs on the EventDblClick event');
    const v = 'http://schemas.microsoft.com/office/visio/2012/main';
    const cells = `<Cells xmlns="${v}"><Cell N="Action" F="CALLTHIS(&quot;Module1.Go&quot;)"/><Cell F="RUNADDON(&quot;x&quot;)"/><Cell N="Address" V="#page"/><Cell N="Address" V=" "/><Cell N="Width" F="2"/></Cells>`;
    const parts = await inspectOffice(await wordWithPart(cells));
    expect(parts.findings.map(f => f.data?.cell)).to.deep.equal(['Action', '']);
  });

  it('classifies hyperlinks and skips those to places inside the document', async () => {
    expect(kinds(await inspectOffice(fixture('lo', 'links.docx')))).to.deep.equal(['LINK/SAFE', 'MACRO/LINK', 'LINK/IP_HOST', 'LINK/CREDENTIALS']);
    const inner = await inspectOffice(
      await wordPackage('<w:p/>', {
        rels: [
          ['rId1', `${T}hyperlink`, '#Bookmark', true],
          ['rId2', `${T}hyperlink`, 'styles.xml'],
          ['rId3', `${T}hyperlink`, 'javascript:alert(1)', true],
          ['rId4', `${T}hyperlink`, 'missing.xml'],
        ],
        parts: { 'word/styles.xml': '<w:styles xmlns:w="urn:x"/>' },
      }),
    );
    expect(kinds(inner)).to.deep.equal(['LINK/SCRIPT_URL', 'LINK/RELATIVE']);
  });

  it('reports 200 findings of one kind and counts the rest in one more, as pdf-defuse does', async () => {
    const rels: Array<[string, string, string, boolean]> = Array.from({ length: 250 }, (_, i) => [`rId${i + 1}`, `${T}hyperlink`, `https://example.com/${i}`, true]);
    const safe = of(await inspectOffice(await wordPackage('<w:p/>', { rels })), 'LINK/SAFE');
    const more = safe.filter(f => f.location === 'additional occurrences');
    expect([safe.length - more.length, more.map(f => f.data?.count), more[0].data?.target]).to.deep.equal([200, [50], 'https://example.com/200']);
  });

  it('reports each kind of hyperlink target', async () => {
    const targets = [
      'https://example.com/',
      'file:///C:/x.exe',
      '\\\\server\\share',
      'data:text/html,x',
      'vbscript:x',
      'ms-search:query',
      'https://a@b.example/',
      'http://127.0.0.1/',
      'http://xn--80ak6aa92e.com/',
      'http://%62ank.example/',
      'notes.txt',
    ];
    const r = await inspectOffice(await wordPackage('<w:p/>', { rels: targets.map((t, i): [string, string, string, boolean] => [`rId${i + 1}`, `${T}hyperlink`, t, true]) }));
    expect(kinds(r)).to.deep.equal([
      'LINK/SAFE',
      'LINK/FILE_URL',
      'LINK/NETWORK_PATH',
      'LINK/DATA_URL',
      'LINK/SCRIPT_URL',
      'LINK/OTHER_SCHEME',
      'LINK/CREDENTIALS',
      'LINK/IP_HOST',
      'LINK/LOOKALIKE_HOST',
      'LINK/ENCODED_HOST',
      'LINK/RELATIVE',
    ]);
  });

  it('reports parts without a content type and parts no relationship reaches, and checks directory entries without a finding', async () => {
    const r = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/stray.dat': 'x', 'word/noext': 'x', 'word/orphan.xml': '<a/>', 'word/media/': '' } }));
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.location}`)).to.deep.equal([
      'PACKAGE/UNTYPED_PART word/stray.dat',
      'PACKAGE/UNTYPED_PART word/noext',
      'PACKAGE/UNREFERENCED_PART word/stray.dat',
      'PACKAGE/UNREFERENCED_PART word/noext',
      'PACKAGE/UNREFERENCED_PART word/orphan.xml',
    ]);
    expect([r.status, r.score]).to.deep.equal(['strippable', 10]);
  });

  it('resolves ".." at the package root to the root, and a target that starts with "/" from the root', async () => {
    const r = await inspectOffice(
      await wordPackage('<w:p/>', {
        packageRels: [['rId9', `${T}customXml`, '../customXml/item1.xml']],
        rels: [['rId2', `${T}customXml`, '/customXml/item2.xml']],
        parts: { 'customXml/item1.xml': '<a/>', 'customXml/item2.xml': '<a/>' },
      }),
    );
    expect(kinds(r)).to.deep.equal([]);
  });

  it('rejects a relationship Id or content type given twice, and keeps the score', async () => {
    const zip = await zipOf([
      ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': WORD_MAIN }, { XML: 'application/xml' })],
      ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
      ['word/document.xml', wordXml('')],
      ['word/_rels/document.xml.rels', relsXml(['rId1', `${T}customXml`, 'x.xml'], ['rId1', `${T}hyperlink`, 'https://example.com', true])],
    ]);
    const r = await inspectOffice(zip);
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.data?.kind ?? ''}`)).to.deep.equal([
      'PACKAGE/DUPLICATE_ENTRY content type default',
      'PACKAGE/DUPLICATE_ENTRY relationship Id',
      'LINK/SAFE ',
    ]);
    expect([r.status, r.score, r.risk]).to.deep.equal(['rejected', 0, 'NONE']);
  });

  it('keeps only content types for entries that exist, and declarations it can read', async () => {
    const types = `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="none" ContentType="a/b"/><Default Extension="none" ContentType="a/b"/><Default ContentType="a/b"/><Default Extension="xml"/><Override PartName="/gone.xml" ContentType="a/b"/><Override PartName="/gone.xml" ContentType="a/b"/><Override ContentType="a/b"/><Override PartName="/word/document.xml" ContentType="${WORD_MAIN}"/><x:Other xmlns:x="urn:x"/></Types>`;
    const r = await inspectOffice(
      await zipOf([
        ['[Content_Types].xml', types],
        ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
        ['word/document.xml', wordXml('')],
      ]),
    );
    expect([r.status, kinds(r)]).to.deep.equal(['clean', []]);
    // Two overrides for a part that is there are refused, since readers may take either.
    const twice = types.replace('<x:Other', `<Override PartName="/WORD/document.xml" ContentType="${WORD_MAIN}"/><x:Other`);
    const dup = await inspectOffice(
      await zipOf([
        ['[Content_Types].xml', twice],
        ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
        ['word/document.xml', wordXml('')],
      ]),
    );
    expect([dup.status, dup.findings[0].data]).to.deep.equal(['rejected', { kind: 'content type override', name: '/WORD/document.xml' }]);
  });

  it('skips relationships without a target and reads one without an Id or a type', async () => {
    const rels = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="a" Type="${T}image"/><Relationship Target="https://example.com/x" TargetMode="External"/><x xmlns="urn:x"/></Relationships>`;
    const r = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/_rels/document.xml.rels': rels } }));
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.location}`)).to.deep.equal(['EXTERNAL/OTHER word/_rels/document.xml.rels, ']);
  });

  it('reads strict OOXML namespaces', async () => {
    const strictW = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
    const zip = await zipOf([
      ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': WORD_MAIN })],
      ['_rels/.rels', relsXml(['rId1', 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', 'word/document.xml'])],
      ['word/document.xml', `${XML}<w:document xmlns:w="${strictW}"><w:body><w:p><w:fldSimple w:instr="DDEAUTO a b"/></w:p></w:body></w:document>`],
      ['word/_rels/document.xml.rels', relsXml(['rId1', 'http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate', 'https://x.example/t.dotm', true])],
    ]);
    expect(kinds(await inspectOffice(zip))).to.deep.equal(['DDE/FIELD', 'EXTERNAL/TEMPLATE']);
  });

  it('rejects a package whose main relationship leaves the package or whose main part has no content type', async () => {
    for (const [what, rels, overrides] of [
      ['external main', relsXml(['rId1', `${T}officeDocument`, 'https://example.com/doc.xml', true]), { 'word/document.xml': WORD_MAIN }],
      ['untyped main', relsXml(['rId1', `${T}officeDocument`, 'word/document.bin']), {}],
    ] as const) {
      const zip = await zipOf([
        ['[Content_Types].xml', contentTypesXml(overrides)],
        ['_rels/.rels', rels],
        ['word/document.xml', wordXml('')],
        ['word/document.bin', 'x'],
      ]);
      const r = await inspectOffice(zip);
      expect([r.status, r.format, kinds(r).slice(-1)[0]], what).to.deep.equal(['rejected', '', 'PACKAGE/UNRECOGNIZED']);
    }
  });

  it('reads relationship targets, hyperlinks, field codes and actions that entities build as their real values', async () => {
    const rels = withSubset(relsXml(['rId2', `${T}hyperlink`, 'http://&host;/a', true], ['rId3', `${T}attachedTemplate`, 'https://&host;/t.dotm', true]), '<!ENTITY host "203.0.113.9">');
    const body = `<w:p>${fld('&dde; inert-app x')}${fld('HYPERLINK "&url;"')}<a:hlinkClick xmlns:a="urn:a" action="&act;"/></w:p>`;
    const doctype = '<!DOCTYPE w:document [<!ENTITY dde "DDE&#65;UTO"><!ENTITY url "javascript:alert(1)"><!ENTITY act "ppaction://program">]>';
    const r = await inspectOffice(await wordPackage(body, { doctype, parts: { 'word/_rels/document.xml.rels': rels } }));
    expect(kinds(r)).to.deep.equal(['XML/INTERNAL_SUBSET', 'DDE/FIELD', 'LINK/SCRIPT_URL', 'ACTION/RUN_PROGRAM', 'XML/INTERNAL_SUBSET', 'LINK/IP_HOST', 'EXTERNAL/TEMPLATE']);
    expect(r.findings.map(f => f.data?.target ?? f.data?.code ?? f.data?.action)).to.deep.equal([
      undefined,
      'DDEAUTO inert-app x',
      'javascript:alert(1)',
      'ppaction://program',
      undefined,
      'http://203.0.113.9/a',
      'https://203.0.113.9/t.dotm',
    ]);
    expect(of(r, 'XML/INTERNAL_SUBSET').map(f => [f.location, f.data])).to.deep.equal([
      ['word/document.xml', { part: 'word/document.xml', name: 'w:document', entities: 3, externalEntities: 0, attributeDefaults: 0, skipped: 0 }],
      ['word/_rels/document.xml.rels', { part: 'word/_rels/document.xml.rels', name: 'r', entities: 1, externalEntities: 0, attributeDefaults: 0, skipped: 0 }],
    ]);
    expect([r.status, r.risk, of(r, 'XML/INTERNAL_SUBSET')[0].action]).to.deep.equal(['strippable', 'CRITICAL', 'strip']);
  });

  it('reports an external entity in text, and reads a field, include, hyperlink or formula built from one as a value it cannot read', async () => {
    const doctype = '<!DOCTYPE w:document [<!ENTITY ext SYSTEM "http://example.invalid/x.xml">]>';
    const body = `<w:p><w:r><w:t>&ext;</w:t></w:r>${fld('&ext; x')}${fld('INCLUDETEXT "&ext;"')}${fld('HYPERLINK "&ext;"')}${fld('MACROBUTTON &ext;')}</w:p>`;
    const r = await inspectOffice(await wordPackage(body, { doctype }));
    expect(kinds(r)).to.deep.equal([
      'XML/INTERNAL_SUBSET',
      'XML/EXTERNAL_ENTITY',
      'XML/EXTERNAL_ENTITY',
      'FIELD/COMPUTED',
      'XML/EXTERNAL_ENTITY',
      'EXTERNAL/INCLUDE',
      'XML/EXTERNAL_ENTITY',
      'LINK/RELATIVE',
      'XML/EXTERNAL_ENTITY',
      'FIELD/MACROBUTTON',
    ]);
    const external = of(r, 'XML/EXTERNAL_ENTITY')[0];
    expect([external.action, external.location, external.data, external.description]).to.deep.equal([
      'strip',
      'word/document.xml',
      { part: 'word/document.xml', entity: 'ext', target: 'http://example.invalid/x.xml' },
      'An XML part takes content from http://example.invalid/x.xml',
    ]);
    const computed = of(r, 'FIELD/COMPUTED')[0];
    expect([computed.data, computed.description]).to.deep.equal([{ code: '&ext; x', entity: 'ext' }, "A field's type comes from external entity ext"]);
    expect(r.findings.filter(f => f.category !== 'XML').map(f => [f.detail, f.data?.target ?? f.data?.macro])).to.deep.equal([
      ['COMPUTED', undefined],
      ['INCLUDE', '&ext;'],
      ['RELATIVE', '&ext;'],
      ['MACROBUTTON', '&ext;'],
    ]);
    expect([r.status, r.risk]).to.deep.equal(['strippable', 'CRITICAL']);

    const sheet = (cells: string) => wordWithPart(withSubset(sheetXml(cells), '<!ENTITY ext PUBLIC "-//X//Y" "x.xml">'));
    const formula = await inspectOffice(await sheet('<c r="A1"><f>&ext;</f></c><c r="A2"><f>"&ext;"&amp;cmd|x!A1</f></c><c r="A3"><f>SUM(A1:A2)</f></c>'));
    expect(kinds(formula)).to.deep.equal(['XML/INTERNAL_SUBSET', 'XML/EXTERNAL_ENTITY', 'DDE/FORMULA', 'XML/EXTERNAL_ENTITY', 'DDE/FORMULA']);
    expect(of(formula, 'XML/EXTERNAL_ENTITY')[0].data?.publicId).to.equal('-//X//Y');
    expect(of(formula, 'DDE/FORMULA').map(f => [f.data, f.description])).to.deep.equal([
      [{ formula: '&ext;', entity: 'ext' }, 'A formula takes text from external entity ext, so it may call DDE'],
      [{ formula: '"&ext;"&cmd|x!A1' }, 'A formula asks another program for data through DDE'],
    ]);
  });

  it('refuses an external entity in a relationship target, as XML 1.0 refuses one in any attribute value', async () => {
    const rels = withSubset(relsXml(['rId2', `${T}hyperlink`, '&ext;', true]), '<!ENTITY ext SYSTEM "http://example.invalid/x">');
    const r = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/_rels/document.xml.rels': rels } }));
    expect([r.status, r.score, kinds(r).slice(-1)[0], r.findings.slice(-1)[0].location]).to.deep.equal(['rejected', null, 'XML/MALFORMED', 'word/_rels/document.xml.rels']);
  });

  it('applies attribute-list defaults to the relationships a part leaves them out of', async () => {
    const subset =
      '<!ATTLIST Relationships xmlns CDATA #FIXED "http://schemas.openxmlformats.org/package/2006/relationships"><!ATTLIST Relationship Target CDATA "\\\\example.invalid\\share\\x.docx" TargetMode CDATA "External">';
    const rels = `${XML}<!DOCTYPE Relationships [${subset}]><Relationships><Relationship Id="rId2" Type="${T}hyperlink"/></Relationships>`;
    const r = await inspectOffice(await wordPackage('<w:p/>', { parts: { 'word/_rels/document.xml.rels': rels } }));
    expect(kinds(r)).to.deep.equal(['XML/INTERNAL_SUBSET', 'LINK/NETWORK_PATH']);
    expect(r.findings[0].data?.attributeDefaults).to.equal(3);
    expect(r.findings[1].data?.target).to.equal('\\\\example.invalid\\share\\x.docx');
  });

  it('reports a document type without an internal subset as information, and one with a subset as removable', async () => {
    const plain = await inspectOffice(await wordPackage('<w:p/>', { doctype: '<!DOCTYPE w:document SYSTEM "word.dtd">' }));
    expect([plain.status, plain.risk, kinds(plain), plain.findings[0].action, plain.findings[0].data]).to.deep.equal([
      'clean',
      'NONE',
      ['XML/DOCTYPE'],
      'info',
      { part: 'word/document.xml', name: 'w:document', systemId: 'word.dtd' },
    ]);
    const d = await inspectOffice(fixture('docs', 'doctype.docx'));
    expect([d.status, d.score, d.format, has(d, 'XML/INTERNAL_SUBSET'), has(d, 'DDE/FIELD')]).to.deep.equal(['strippable', 20, 'docx', true, false]);
  });

  it('stops at XML it refuses, keeping what it found before', async () => {
    for (const [name, k] of [
      ['malformed.docx', 'XML/MALFORMED'],
      ['latin1.docx', 'XML/ENCODING'],
    ]) {
      const r = await inspectOffice(fixture('docs', name));
      expect([r.status, r.score, r.format, kinds(r).slice(-1)[0]], name).to.deep.equal(['rejected', null, 'docx', k]);
      expect(must(r.findings.slice(-1)[0], 'finding').location).to.equal('word/document.xml');
    }
    expect(has(await inspectOffice(fixture('docs', 'utf16.docx')), 'XML/ENCODING')).to.equal(false);
  });

  it('cuts values from the file to 200 characters in finding data', async () => {
    const long = 'x'.repeat(500);
    const r = await inspectOffice(await wordPackage(`<w:p><w:fldSimple w:instr="MACROBUTTON ${long}"/></w:p>`));
    expect(String(r.findings[0].data?.macro)).to.have.length(200);
    expect(String(r.findings[0].data?.macro).endsWith('...')).to.equal(true);
  });
});
