import { expect } from 'chai';
import { disarmOffice, inspectOffice } from '../src';
import { entryText } from './helpers/structure';
import { fixture, kinds, must, odfPackage, sheetXml, textBody, wordWithPart } from './helpers/util';

const sheet = (formula: string) => wordWithPart(sheetXml(`<c r="A1"><f>${formula}</f><v>7</v></c>`));
const ods = (formula: string) =>
  odfPackage(
    `<office:body><office:spreadsheet><table:table><table:table-row><table:table-cell table:formula="${formula}" office:value="7"/></table:table-row></table:table></office:spreadsheet></office:body>`,
    {
      mimetype: 'application/vnd.oasis.opendocument.spreadsheet',
    },
  );
const functions = (r: { findings: Array<{ category: string; detail: string; data?: Record<string, string | number> }> }) =>
  r.findings.filter(f => f.detail === 'FORMULA').map(f => `${f.category}/${f.data?.function}`);

describe('formula functions that reach outside the file', () => {
  it('reports functions that fetch or run code in OOXML formulas, and keeps the cached value without the formula', async () => {
    expect(functions(await inspectOffice(await sheet('WEBSERVICE("https://example.invalid/")')))).to.deep.equal(['EXTERNAL/WEBSERVICE']);
    expect(functions(await inspectOffice(await sheet('_xlfn.IMAGE("https://example.invalid/a.png")')))).to.deep.equal(['EXTERNAL/_XLFN.IMAGE']);
    expect(functions(await inspectOffice(await sheet('RTD("inert.progid",,"topic")+CALL ("kernel32","WinExec","JCJ","calc",0)')))).to.deep.equal(['MACRO/RTD', 'MACRO/CALL']);
    // A name inside quotes is text, and a cell reference or another function with the same letters is not the function.
    expect(functions(await inspectOffice(await sheet('"WEBSERVICE(x)"&amp;MYWEBSERVICE(1)&amp;\'RTD(\'!A1')))).to.deep.equal([]);
    const r = await disarmOffice(await sheet('WEBSERVICE("https://example.invalid/")'));
    const part = must(await entryText(must(r.bytes, 'output'), 'word/part.xml'), 'part');
    expect([r.status, part.includes('WEBSERVICE'), part.includes('<v>7</v>')]).to.deep.equal(['defused', false, true]);
  });

  it('reports them in OpenDocument formulas, under the prefixes LibreOffice writes', async () => {
    expect(functions(await inspectOffice(await ods('of:=COM.MICROSOFT.WEBSERVICE(&quot;https://example.invalid/&quot;)')))).to.deep.equal(['EXTERNAL/COM.MICROSOFT.WEBSERVICE']);
    const r = await disarmOffice(await ods('of:=WEBSERVICE(&quot;https://example.invalid/&quot;)'));
    const content = must(await entryText(must(r.bytes, 'output'), 'content.xml'), 'content');
    expect([r.status, content.includes('table:formula'), content.includes('office:value="7"')]).to.deep.equal(['defused', false, true]);
  });

  it('reports them in BIFF cell formulas, by function index and by the name a newer function is called by', async () => {
    const r = await inspectOffice(fixture('ole', 'formula-functions.xls'));
    expect([kinds(r), functions(r)]).to.deep.equal([
      ['MACRO/FORMULA', 'EXTERNAL/FORMULA'],
      ['MACRO/CALL', 'EXTERNAL/_XLFN.WEBSERVICE'],
    ]);
  });
});

describe('data connections, form submit targets and the hyperlink base', () => {
  const CONNECTIONS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml';
  it('reports each source a data connection names, and empties it', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="q" type="1" refreshOnLoad="1"><dbPr connection="Provider=MSDASQL;DSN=inert" command="SELECT 1"/></connection><connection id="2" name="w" type="4"><webPr url="https://example.invalid/data"/></connection></connections>`;
    const input = await wordWithPart(xml, CONNECTIONS);
    const r = await disarmOffice(input);
    expect([r.status, r.before.findings.map(f => `${f.category}/${f.detail}/${f.data?.attribute}`)]).to.deep.equal([
      'defused',
      ['EXTERNAL/DATA_CONNECTION/connection', 'EXTERNAL/DATA_CONNECTION/command', 'EXTERNAL/DATA_CONNECTION/url'],
    ]);
    const part = must(await entryText(must(r.bytes, 'output'), 'word/part.xml'), 'part');
    expect([part.includes('MSDASQL'), part.includes('example.invalid'), part.includes('<connection id="2"'), kinds(must(r.after, 'after'))]).to.deep.equal([false, false, true, []]);
  });

  it("checks Word's hyperlink base, which relative links resolve against", async () => {
    const app = (base: string) =>
      wordWithPart(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><HyperlinkBase>${base}</HyperlinkBase></Properties>`,
        'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      );
    expect(kinds(await inspectOffice(await app('file:///C:/inert/')))).to.deep.equal(['LINK/FILE_URL']);
    expect(kinds(await inspectOffice(await app('https://example.invalid/')))).to.deep.equal(['LINK/SAFE']);
  });

  it('reports where an OpenDocument form sends its fields, and takes the target away', async () => {
    const input = await odfPackage(textBody('<office:forms><form:form form:name="f" xlink:href="https://example.invalid/collect" form:method="post"/></office:forms>'));
    const r = await disarmOffice(input);
    expect([r.status, r.before.findings.map(f => `${f.category}/${f.detail}`)]).to.deep.equal(['defused', ['ACTION/SUBMIT_FORM']]);
    expect(must(await entryText(must(r.bytes, 'output'), 'content.xml'), 'content').includes('example.invalid')).to.equal(false);
  });
});
