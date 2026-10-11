import { expect } from 'chai';
import { callsDde, fieldWords, Quotes } from '../src/formulas';
import { checkLink, isAbsoluteTarget, isFragment, isInsidePackage, isLocalTarget } from '../src/links';
import { UNKNOWN } from '../src/xml';

/** Whether "|" stands outside quotes, as the OOXML rules look for Excel's DDE syntax. */
function hasDdeBar(formula: string): boolean {
  const q = new Quotes();
  let found = false;
  for (const c of formula) if (q.outside(c) && c === '|') found = true;
  return found;
}

describe('link and formula checks', () => {
  it('classifies hyperlink targets as pdf-defuse does', () => {
    const cases: Array<[string, string]> = [
      ['https://example.com/a', 'LINK/SAFE'],
      ['HTTP://Example.COM', 'LINK/SAFE'],
      ['mailto:a@example.com?subject=x', 'LINK/SAFE'],
      ['file:///C:/Windows/notepad.exe', 'LINK/FILE_URL'],
      ['C:\\Windows\\notepad.exe', 'LINK/FILE_URL'],
      ['file://server/share/x', 'LINK/NETWORK_PATH'],
      ['\\\\server\\share\\x', 'LINK/NETWORK_PATH'],
      ['//server/share', 'LINK/NETWORK_PATH'],
      ['smb://server/share', 'LINK/NETWORK_PATH'],
      ['data:text/html,<b>x</b>', 'LINK/DATA_URL'],
      ['javascript:alert(1)', 'LINK/SCRIPT_URL'],
      [' java\tscript:alert(1)', 'LINK/SCRIPT_URL'],
      ['vbscript:msgbox', 'LINK/SCRIPT_URL'],
      ['vnd.sun.star.script:Standard.Module1.Main?language=Basic&location=document', 'MACRO/LINK'],
      ['macro:///Standard.Module1.Main', 'MACRO/LINK'],
      ['ms-msdt:/id PCWDiagnostic', 'LINK/OTHER_SCHEME'],
      ['https://bank.example@evil.example/', 'LINK/CREDENTIALS'],
      ['http://10.0.0.1/', 'LINK/IP_HOST'],
      ['http://0x7f000001/', 'LINK/IP_HOST'],
      ['http://[::1]/', 'LINK/IP_HOST'],
      ['http://xn--e1afmkfd.example/', 'LINK/LOOKALIKE_HOST'],
      ['http://%65vil.example/', 'LINK/ENCODED_HOST'],
      ['other.docx', 'LINK/RELATIVE'],
      ['../x/y.pdf', 'LINK/RELATIVE'],
      ['%6Aavascript:alert(1)', 'LINK/RELATIVE'],
      ['file://%zz/x', 'LINK/FILE_URL'],
      ['mailto:someone', 'LINK/OTHER_SCHEME'],
      ['mailto:%E0%A4%A', 'LINK/OTHER_SCHEME'],
      ['mailto:a@1234', 'LINK/IP_HOST'],
      ['mailto:a@0x7f000001', 'LINK/IP_HOST'],
      ['mailto:a@xn--e1afmkfd.example', 'LINK/LOOKALIKE_HOST'],
      ['http://', 'LINK/OTHER_SCHEME'],
      ['http://1234/', 'LINK/IP_HOST'],
      ['http:\\\\evil.example\\x', 'LINK/SAFE'],
      ['https://user:pass@example.com/', 'LINK/CREDENTIALS'],
      ['https://example.com./', 'LINK/SAFE'],
    ];
    for (const [target, expected] of cases) {
      const v = checkLink(target);
      expect(`${v.category}/${v.detail}`, target).to.equal(expected);
    }
  });

  it('names the host of a safe link as the URL parser reads it', () => {
    expect(checkLink('HTTPS://Example.COM./a').host).to.equal('example.com');
    expect(checkLink('mailto:a@Example.com').host).to.equal('example.com');
  });

  it('tells local targets from those outside the machine', () => {
    for (const t of ['C:\\Users\\x\\t.dotx', 'file:///C:/x.dotx', 'file://localhost/x', '/home/x/t.ott', '../t.ott', 't.ott']) expect(isLocalTarget(t), t).to.equal(true);
    expect(isLocalTarget('file://%zz/x')).to.equal(true);
    for (const t of [
      '\\\\host\\share\\t.dotx',
      '//host/share',
      'file://host/share/t.dotx',
      'https://x.example/t.dotm',
      'smb://host/x',
      'mhtml:http://x!x',
      '%73cript:%68ttp://example.invalid/x',
      '%5c%5chost%5cshare',
    ])
      expect(isLocalTarget(t), t).to.equal(false);
  });

  it('tells package paths from paths outside the package', () => {
    for (const t of ['Pictures/a.png', './Object 1', 'Object 1', 'a/../b']) expect(isInsidePackage(t), t).to.equal(true);
    for (const t of ['../a.png', 'a/../../b', '/a', '\\a', 'C:/a', 'http://x/', '', '%2e%2e/a.png', '%68ttp://x/']) expect(isInsidePackage(t), t).to.equal(false);
    for (const t of ['https://x/', '%68ttp://x/', 'C:\\a', '\\server\\x', '\\a', '//server/x']) expect(isAbsoluteTarget(t), t).to.equal(true);
    for (const t of ['word/x.xml', '/word/x.xml', '../x.xml', 'a%20b.xml']) expect(isAbsoluteTarget(t), t).to.equal(false);
    expect(isFragment(' #Top')).to.equal(true);
    expect(isFragment('x#Top')).to.equal(false);
    // RFC 3986: an empty reference is the document itself.
    expect(isFragment('')).to.equal(true);
    expect(isFragment(' \t')).to.equal(true);
    // A target filled from an external entity is not known to be empty.
    expect(isFragment(UNKNOWN)).to.equal(false);
  });

  it("finds Excel's DDE syntax only outside quotes", () => {
    expect(hasDdeBar("cmd|' /c calc'!A0")).to.equal(true);
    expect(hasDdeBar("=MSEXCEL|'\\..\\..\\..\\Windows\\System32\\cmd.exe /c calc'!A1")).to.equal(true);
    expect(hasDdeBar('"a|b"&\'x|y\'!A1')).to.equal(false);
    expect(hasDdeBar('"say ""|"" here"')).to.equal(false);
    expect(hasDdeBar('SUM(A1:A2)')).to.equal(false);
  });

  it('finds the DDE function in OpenDocument formulas only as a call', () => {
    expect(callsDde('DDE("soffice";"x.ods";"A1")')).to.equal(true);
    expect(callsDde('of:=DDE("soffice";"x.ods";"A1")')).to.equal(true);
    expect(callsDde('of:=dde ("a";"b";"c")')).to.equal(true);
    expect(callsDde('of:="DDE(" & 1')).to.equal(false);
    expect(callsDde('of:=MYDDE(1)')).to.equal(false);
    expect(callsDde('of:=[.DDE(1)]')).to.equal(false);
  });

  it('splits Word field codes into words, switches and quoted arguments', () => {
    expect(fieldWords(' INCLUDETEXT "C:\\\\a b\\\\c.docx" \\* MERGEFORMAT ')).to.deep.equal([
      { text: 'INCLUDETEXT', quoted: false },
      { text: 'C:\\a b\\c.docx', quoted: true },
      { text: '\\*', quoted: false },
      { text: 'MERGEFORMAT', quoted: false },
    ]);
    expect(fieldWords('LINK Excel.Sheet.12 C:\\\\x.xlsx').map(w => w.text)).to.deep.equal(['LINK', 'Excel.Sheet.12', 'C:\\x.xlsx']);
  });
});
