import { expect } from 'chai';
import { inspectOffice } from '../src';
import { fixture, kinds, of } from './helpers/util';

const rtf = (name: string) => fixture('rtf', name);
const text = (s: string) => Buffer.from(s, 'latin1');
const doc = (body: string) => text(`{\\rtf1\\ansi ${body}}`);
/** OLE 1.0 data for an embedded object of `cls` holding `native`, as hex. */
const ole1 = (cls: string, native: Buffer) => {
  const lp = (s: string) => Buffer.concat([Buffer.from([s.length + 1, 0, 0, 0]), Buffer.from(`${s}\0`, 'latin1')]);
  const size = Buffer.alloc(4);
  size.writeUInt32LE(native.length);
  return Buffer.concat([Buffer.from('0105000002000000', 'hex'), lp(cls), lp(''), lp(''), size, native]).toString('hex');
};

describe('RTF', () => {
  it('reads plain RTF, and any file that starts with {\\rt, as Word does', async () => {
    for (const name of ['plain.rtf', 'header-only.rtf']) {
      const r = await inspectOffice(rtf(name));
      expect([r.status, r.format, r.mimeType, kinds(r)]).to.deep.equal(['clean', 'rtf', 'application/rtf', []]);
    }
    expect(kinds(await inspectOffice(fixture('lo', 'clean.rtf')))).to.deep.equal(['LINK/SAFE', 'LINK/SAFE']);
  });

  it('decodes objects and reports what they hold: a packaged file, Equation Editor, a link, a control', async () => {
    const pkg = await inspectOffice(rtf('package.rtf'));
    expect(pkg.findings.map(f => [`${f.category}/${f.detail}`, f.data?.fileName, f.data?.className])).to.deep.equal([['EMBEDDED/OLE_OBJECT', 'inert.txt', 'Package']]);
    expect(kinds(await inspectOffice(rtf('equation.rtf')))).to.deep.equal(['ACTION/OBJECT_UPDATE', 'EMBEDDED/OLE_OBJECT', 'EMBEDDED/EQUATION']);
    const link = await inspectOffice(rtf('ole2link.rtf'));
    expect(link.findings.map(f => [`${f.category}/${f.detail}`, f.data?.target])).to.deep.equal([
      ['ACTION/OBJECT_UPDATE', undefined],
      ['EXTERNAL/OBJECT', 'https://example.invalid/inert.hta'],
    ]);
    expect(of(await inspectOffice(rtf('linked.rtf')), 'EXTERNAL/OBJECT')[0].data?.target).to.equal('\\\\inert.example.invalid\\share\\book.xls');
    expect(kinds(await inspectOffice(rtf('linked-no-data.rtf')))).to.deep.equal(['EXTERNAL/OBJECT']);
    expect(kinds(await inspectOffice(rtf('objocx.rtf')))).to.deep.equal(['CONTROL/ACTIVEX']);
    expect(of(await inspectOffice(rtf('not-ole1.rtf')), 'EMBEDDED/OLE_OBJECT')[0].data?.reason).to.equal('its data is not an OLE 1.0 object');
    expect(of(await inspectOffice(rtf('bare-objdata.rtf')), 'EMBEDDED/OLE_OBJECT')[0].data?.fileName).to.equal('inert.txt');
  });

  it("reads object data as Word does: junk, white space, a digit dropped before \\'hh, \\bin data and groups inside it", async () => {
    const r = await inspectOffice(rtf('obfuscated.rtf'));
    // The data holds a {\*\inertdest 4242} group. Read with its digits the data is no object; read without them it is
    // the equation. Both readings are reported.
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.className, f.data?.reading])).to.deep.equal([
      ['ACTION/OBJECT_UPDATE', undefined, undefined],
      ['EMBEDDED/OLE_OBJECT', undefined, undefined],
      ['EMBEDDED/OLE_OBJECT', 'Equation.3', 'without the groups \\* marks'],
      ['EMBEDDED/EQUATION', 'Equation.3', 'without the groups \\* marks'],
    ]);
  });

  it("reads a group that \\* marks in object data as the data around it: \\'hh, \\bin and such a group inside it", async () => {
    const h = ole1('Inert.Class.1', Buffer.from('inert marker'));
    const bin = String.fromCharCode(Number.parseInt(h.slice(4, 6), 16));
    const r = await inspectOffice(doc(`{\\object\\objemb{\\*\\objdata {\\*\\par1 ${h.slice(0, 2)}7\\'41{\\*\\par2 ${h.slice(2, 4)}}\\bin1 ${bin}}${h.slice(6)}}}`));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.className, f.data?.reading])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'Inert.Class.1', undefined],
      ['EMBEDDED/OLE_OBJECT', undefined, 'without the groups \\* marks'],
    ]);
  });

  it('reads fields built across groups and from \\u characters, the template, and shape hyperlinks', async () => {
    const r = await inspectOffice(rtf('fields.rtf'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.code ?? f.data?.target])).to.deep.equal([
      ['EXTERNAL/TEMPLATE', '\\\\inert.example.invalid\\share\\inert.dotm'],
      ['DDE/FIELD', 'DDEAUTO inert-app "inert topic"'],
      ['EXTERNAL/INCLUDE', 'INCLUDEPICTURE "https://example.invalid/inert.png" \\d'],
      ['LINK/SAFE', 'HYPERLINK "https://example.invalid/page"'],
      ['FIELD/MACROBUTTON', 'MACROBUTTON InertMacro click'],
      ['FIELD/COMPUTED', 'inert-app topic'],
      ['LINK/FILE_URL', 'file:///C:/inert/page.html'],
    ]);
  });

  it('reads past the document group, as rtfobj does, and reports what lies there as trailing data', async () => {
    const r = await inspectOffice(rtf('trailing.rtf'));
    expect(kinds(r)).to.deep.equal(['EMBEDDED/OLE_OBJECT', 'CONTAINER/TRAILING_BYTES']);
    // White space and NULs after the last brace are what writers leave, not data.
    expect(kinds(await inspectOffice(text('{\\rtf1 x}\r\n\0\0  ')))).to.deep.equal([]);
    expect(kinds(await inspectOffice(text('{\\rtf1 x}}}junk')))).to.deep.equal(['CONTAINER/TRAILING_BYTES']);
    expect(kinds(await inspectOffice(text('{\\rtf1 x}\\info{\\title t}')))).to.deep.equal(['CONTAINER/TRAILING_BYTES']);
  });

  it('reads a file that ends inside its groups, as Word does', async () => {
    expect(kinds(await inspectOffice(rtf('unbalanced.rtf')))).to.deep.equal(['DDE/FIELD']);
    expect(kinds(await inspectOffice(text('{\\rtf1 x\\')))).to.deep.equal([]);
    expect(kinds(await inspectOffice(text("{\\rtf1 x\\'4")))).to.deep.equal([]);
  });

  it('stops at the group and destination nesting limits', async () => {
    expect(kinds(await inspectOffice(rtf('deep.rtf')))).to.deep.equal(['LIMIT/RTF_DEPTH']);
    expect(kinds(await inspectOffice(rtf('many-destinations.rtf')))).to.deep.equal(['LIMIT/RTF_DEPTH']);
    expect(kinds(await inspectOffice(rtf('deep.rtf'), { limits: { rtfDepth: 400 } }))).to.deep.equal([]);
  });

  it('refuses RTF that is also a ZIP archive', async () => {
    expect(kinds(await inspectOffice(rtf('zip-polyglot.rtf')))).to.deep.equal(['PACKAGE/AMBIGUOUS_FORMAT']);
  });

  it('skips destinations marked with \\* that it does not know, and keeps the text around unknown control words', async () => {
    const r = await inspectOffice(doc('{\\field{\\*\\fldinst DD{\\*\\inert hidden}E\\inertword AUTO a b}{\\fldrslt x}}'));
    expect(r.findings.map(f => f.data?.code)).to.deep.equal(['DDEAUTO a b']);
    // Control symbols: literal braces and backslashes, a nonbreaking space, and a line end as a paragraph mark.
    const symbols = await inspectOffice(doc('{\\field{\\*\\fldinst INCLUDETEXT "C:\\\\\\\\x\\{y\\}"\\~\\\n}}'));
    expect(symbols.findings.map(f => f.data?.target)).to.deep.equal(['C:\\x{y}']);
  });

  it("reads \\u characters with the fallback count \\uc sets, and \\'hh characters", async () => {
    const r = await inspectOffice(doc("{\\uc2\\field{\\*\\fldinst \\u68\\'3f\\'3fDE\\u-65471??UTO a}}{\\field{\\*\\fldinst \\'4d\\'41CROBUTTON m}}"));
    expect(r.findings.map(f => f.data?.code)).to.deep.equal(['DDEAUTO a', 'MACROBUTTON m']);
  });

  it('reads an object from its class word when its data names none, and Equation Editor by its class id alone', async () => {
    const r = await inspectOffice(doc(`{\\object\\objemb{\\*\\objclass Equation.3}{\\*\\objdata ${ole1('', Buffer.from('x'))}}}`));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.className])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'Equation.3'],
      ['EMBEDDED/EQUATION', 'Equation.3'],
    ]);
    const native = fixture('ole', 'bare-object.bin');
    const stored = await inspectOffice(doc(`{\\object\\objemb{\\*\\objdata ${ole1('Inert.Class', native)}}}`));
    expect(stored.findings.map(f => [`${f.category}/${f.detail}`, f.data?.fileName])).to.deep.equal([['EMBEDDED/OLE_OBJECT', 'inert.txt']]);
    const broken = await inspectOffice(doc(`{\\object\\objemb{\\*\\objdata ${ole1('Inert.Class', Buffer.concat([native.subarray(0, 600)]))}}}`));
    expect(String(broken.findings[0].data?.reason)).to.match(/^its storage does not read/);
    const empty = await inspectOffice(doc('{\\object\\objemb{\\*\\objclass Inert}{\\*\\objdata }}'));
    expect(kinds(empty)).to.deep.equal([]);
  });

  it("decodes object data whose control words and \\bin data cross the reader's 256 KiB chunks", async () => {
    const data = ole1('Package', Buffer.alloc(0));
    // The bytes of the format id go in as \bin data.
    const object = `{\\object\\objemb\\objupdate{\\*\\objdata ${data.slice(0, 8)}\\bin2 \u0002\u0000${data.slice(12)}}}`;
    const head = '{\\rtf1\\ansi '.length;
    for (let at = 0; at < 70; at++) {
      const r = await inspectOffice(doc(`${'x'.repeat(262144 - head - at)}${object}`));
      expect({ at, kinds: kinds(r) }).to.deep.equal({ at, kinds: ['ACTION/OBJECT_UPDATE', 'EMBEDDED/OLE_OBJECT'] });
    }
  });

  it('applies the decompressed size limits to object data', async () => {
    const body = `{\\object\\objemb{\\*\\objdata ${ole1('Inert', Buffer.alloc(2000))}}}`;
    expect(kinds(await inspectOffice(doc(body), { limits: { decompressedBytes: 1000 } }))).to.deep.equal(['LIMIT/DECOMPRESSED_SIZE']);
    expect(kinds(await inspectOffice(doc(body + body), { limits: { totalDecompressedBytes: 3000 } }))).to.deep.equal(['LIMIT/TOTAL_DECOMPRESSED_SIZE']);
  });

  it('reads linked objects of the Equation class or of no class, nested files with hidden data, and a link with no topic', async () => {
    const eq = await inspectOffice(rtf('linked-equation.rtf'));
    expect(eq.findings.map(f => [`${f.category}/${f.detail}`, f.data?.target ?? f.data?.clsid])).to.deep.equal([
      ['EXTERNAL/OBJECT', 'https://example.invalid/inert.hta'],
      ['EMBEDDED/EQUATION', '0002CE02-0000-0000-C000-000000000046'],
    ]);
    expect(kinds(await inspectOffice(rtf('linked-no-class.rtf')))).to.deep.equal(['EXTERNAL/OBJECT']);
    const nested = await inspectOffice(rtf('nested-hidden.rtf'));
    expect(nested.findings.map(f => [`${f.category}/${f.detail}`, f.location])).to.deep.equal([
      ['CONTAINER/HIDDEN_DATA', 'object at offset 62'],
      ['EMBEDDED/OLE_OBJECT', 'object at offset 62'],
    ]);
    expect(kinds(await inspectOffice(rtf('linked-topic-only.rtf')))).to.deep.equal(['EXTERNAL/OBJECT']);
    const lp = (v: string) => Buffer.concat([Buffer.from([v.length + 1, 0, 0, 0]), Buffer.from(`${v}\0`, 'latin1')]);
    const linked = Buffer.concat([Buffer.from('0105000001000000', 'hex'), lp(''), lp('\\\\inert.example.invalid\\a.xls'), lp(''), lp(''), Buffer.alloc(8)]).toString('hex');
    const r = await inspectOffice(doc(`{\\object\\objlink{\\*\\objdata ${linked}}}`));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data])).to.deep.equal([['EXTERNAL/OBJECT', { object: 'object at offset 13', target: '\\\\inert.example.invalid\\a.xls' }]]);
    // A nested storage over a limit stops the run.
    expect(kinds(await inspectOffice(rtf('nested-hidden.rtf'), { limits: { entries: 2 } }))).to.deep.equal(['LIMIT/ENTRY_COUNT']);
  });

  it("reads \\bin without a length or with a negative one as no data, \\uc without a count as one, and a \\'hh the file cuts short", async () => {
    const r = await inspectOffice(doc("{\\field{\\*\\fldinst \\bin DD\\bin-5 E\\uc\\u65?UTO a}}x\\'"));
    expect(r.findings.map(f => f.data?.code)).to.deep.equal(['DDEAUTO a']);
    expect(kinds(await inspectOffice(text("{\\rtf1 x\\'")))).to.deep.equal([]);
  });

  it("keeps a field or object open when a destination follows in its own group, and reads a \\'hh across the reader's chunks", async () => {
    const r = await inspectOffice(doc('{\\field\\fldinst DDEAUTO a b}{\\object\\objocx\\objdata 01}'));
    expect(kinds(r)).to.deep.equal(['DDE/FIELD', 'CONTROL/ACTIVEX']);
    for (let at = 0; at < 6; at++) {
      const pad = 'x'.repeat(262144 - '{\\rtf1\\ansi '.length - '{\\field{\\*\\fldinst '.length - at);
      const field = await inspectOffice(doc(`${pad}{\\field{\\*\\fldinst \\'44DEAUTO a}}`));
      expect({ at, codes: field.findings.map(f => f.data?.code) }).to.deep.equal({ at, codes: ['DDEAUTO a'] });
    }
  });
});
