import { expect } from 'chai';
import { disarmOffice, inspectOffice } from '../src';
import { entryText } from './helpers/structure';
import { kinds, must, odfPackage, of, relsXml, S, T, textBody, wordPackage, wordWithPart, XML } from './helpers/util';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** A Word document with one hyperlink, whose text and tooltip are given, to `target`. */
const wordLink = (target: string, text: string, tooltip?: string) =>
  wordPackage(`<w:p><w:hyperlink r:id="rIdL"${tooltip === undefined ? '' : ` w:tooltip="${tooltip}"`}><w:r><w:t>${text}</w:t></w:r></w:hyperlink></w:p>`, {
    rels: [['rIdL', `${T}hyperlink`, target, true]],
  });

/** A Word document with a HYPERLINK field whose code and result are given. */
const fieldLink = (code: string, result: string) =>
  wordPackage(
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
  );

/** A part of the document related from it, with its own relationships. */
const partWithLink = (xml: string, target: string) => wordWithPart(xml, 'application/xml', { parts: { 'word/_rels/part.xml.rels': relsXml(['rIdL', `${T}hyperlink`, target, true]) } });

const odt = (body: string) => odfPackage(textBody(body));

describe('link labels', () => {
  it('reports a link whose text names another site, and keeps the text without the link', async () => {
    const input = await wordLink('https://evil.example/login', 'www.mybank.com');
    const r = await disarmOffice(input);
    expect([r.status, kinds(r.before)]).to.deep.equal(['defused', ['LINK/SAFE', 'LINK/TEXT_MISMATCH']]);
    expect(of(r.before, 'LINK/TEXT_MISMATCH')[0].data).to.include({ target: 'https://evil.example/login', label: 'www.mybank.com' });
    const out = must(r.bytes, 'output');
    const doc = must(await entryText(out, 'word/document.xml'), 'document');
    expect([doc.includes('w:hyperlink'), doc.includes('www.mybank.com'), (await entryText(out, 'word/_rels/document.xml.rels'))?.includes('evil.example')]).to.deep.equal([false, true, false]);
  });

  it('accepts text that names the same site, a parent or a subdomain of it, or no site at all', async () => {
    for (const [target, text] of [
      ['https://login.mybank.com/x', 'mybank.com'],
      ['https://mybank.com/x', 'www.mybank.com/help'],
      ['https://www.mybank.com/', 'https://MyBank.com/'],
      ['https://evil.example/', 'click here'],
      ['mailto:help@mybank.com', 'help@mybank.com'],
    ]) {
      expect(kinds(await inspectOffice(await wordLink(target, text))), `${target} ${text}`).to.deep.equal(['LINK/SAFE']);
    }
    // Text that is not itself an address names no site: an Excel function name, a file name, prose, a host in
    // capitals without WWW, and a URL cut short to a host without a dot.
    for (const text of ['BETA.DIST', 'CHIQU.VERT.RE', 'moon.htm', 'see mybank.com for more', 'MYBANK.COM', 'http://cites']) {
      expect(kinds(await inspectOffice(await wordLink('https://evil.example/', text))), text).to.deep.equal(['LINK/SAFE']);
    }
    expect(kinds(await inspectOffice(await wordLink('https://evil.example/', 'WWW.MYBANK.COM')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
    // A subdomain under another site is not that site: evil.co.uk is not under mybank.co.uk.
    expect(kinds(await inspectOffice(await wordLink('https://evil.co.uk/', 'mybank.co.uk')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
    expect(kinds(await inspectOffice(await wordLink('mailto:help@evil.example', 'help@mybank.com')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
  });

  it('compares a tooltip too, and leaves a link that fails on its own to its own finding', async () => {
    expect(kinds(await inspectOffice(await wordLink('https://evil.example/', 'Sign in', 'https://www.mybank.com/signin')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
    expect(kinds(await inspectOffice(await wordLink('http://203.0.113.9/', 'www.mybank.com')))).to.deep.equal(['LINK/IP_HOST']);
  });

  it("compares a HYPERLINK field's result and \\o tooltip, in fldChar and fldSimple form", async () => {
    const complex = await disarmOffice(await fieldLink('HYPERLINK "https://evil.example/"', 'www.mybank.com'));
    expect([complex.status, kinds(complex.before)]).to.deep.equal(['defused', ['LINK/SAFE', 'LINK/TEXT_MISMATCH']]);
    const doc = must(await entryText(must(complex.bytes, 'output'), 'word/document.xml'), 'document');
    expect([doc.includes('HYPERLINK'), doc.includes('www.mybank.com')]).to.deep.equal([false, true]);
    expect(kinds(await inspectOffice(await fieldLink('HYPERLINK "https://evil.example/" \\o "mybank.com"', 'Sign in')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
    expect(kinds(await inspectOffice(await fieldLink('HYPERLINK "https://www.mybank.com/"', 'mybank.com')))).to.deep.equal(['LINK/SAFE']);
    const simple = await wordPackage('<w:p><w:fldSimple w:instr=\'HYPERLINK "https://evil.example/"\'><w:r><w:t>www.mybank.com</w:t></w:r></w:fldSimple></w:p>');
    expect(kinds(await inspectOffice(simple))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
  });

  it('compares the text of a DrawingML run and an Excel display name', async () => {
    const run = `${XML}<root xmlns:a="${A}" xmlns:r="${R}"><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdL"/></a:rPr><a:t>www.mybank.com</a:t></a:r></a:p></root>`;
    expect(kinds(await inspectOffice(await partWithLink(run, 'https://evil.example/')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
    const sheet = `${XML}<worksheet xmlns="${S}" xmlns:r="${R}"><sheetData/><hyperlinks><hyperlink ref="A1" r:id="rIdL" display="www.mybank.com"/></hyperlinks></worksheet>`;
    expect(kinds(await inspectOffice(await partWithLink(sheet, 'https://evil.example/')))).to.deep.equal(['LINK/SAFE', 'LINK/TEXT_MISMATCH']);
  });

  it('compares the text and title of an OpenDocument link', async () => {
    const text = await disarmOffice(await odt('<text:p><text:a xlink:href="https://evil.example/">www.mybank.com</text:a></text:p>'));
    expect([text.status, kinds(text.before)]).to.deep.equal(['defused', ['LINK/SAFE', 'LINK/TEXT_MISMATCH']]);
    const content = must(await entryText(must(text.bytes, 'output'), 'content.xml'), 'content');
    expect([content.includes('text:a'), content.includes('www.mybank.com')]).to.deep.equal([false, true]);
    expect(kinds(await inspectOffice(await odt('<text:p><text:a xlink:href="https://evil.example/" office:title="mybank.com">Sign in</text:a></text:p>')))).to.deep.equal([
      'LINK/SAFE',
      'LINK/TEXT_MISMATCH',
    ]);
    expect(kinds(await inspectOffice(await odt('<text:p><text:a xlink:href="https://www.mybank.com/">mybank.com</text:a></text:p>')))).to.deep.equal(['LINK/SAFE']);
  });

  it("compares an RTF HYPERLINK field's result", async () => {
    const rtf = Buffer.from('{\\rtf1{\\field{\\*\\fldinst HYPERLINK "https://evil.example/"}{\\fldrslt www.mybank.com}}}', 'latin1');
    const r = await disarmOffice(rtf);
    expect([r.status, kinds(r.before)]).to.deep.equal(['defused', ['LINK/SAFE', 'LINK/TEXT_MISMATCH']]);
    const out = Buffer.from(must(r.bytes, 'output')).toString('latin1');
    expect([out.includes('evil.example'), out.includes('www.mybank.com')]).to.deep.equal([false, true]);
  });
});
