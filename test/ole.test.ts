import { expect } from 'chai';
import { inspectOffice } from '../src';
import { fixture, kinds, of } from './helpers/util';

const ole = (name: string) => fixture('ole', name);
/** The kinds of what a file holds, without the information findings. */
const actionable = (r: { findings: Array<{ category: string; detail: string; action: string }> }) => r.findings.filter(f => f.action !== 'info').map(f => `${f.category}/${f.detail}`);

describe('binary Office formats', () => {
  it('names each format by its content', async () => {
    const formats: Array<[string, string, string]> = [
      ['word.doc', 'doc', 'application/msword'],
      ['word-8bit.doc', 'doc', 'application/msword'],
      ['template.dot', 'dot', 'application/msword'],
      ['word6.doc', 'doc', 'application/msword'],
      ['excel.xls', 'xls', 'application/vnd.ms-excel'],
      ['excel5.xls', 'xls', 'application/vnd.ms-excel'],
      ['template.xlt', 'xlt', 'application/vnd.ms-excel'],
      ['addin.xla', 'xla', 'application/vnd.ms-excel'],
      ['powerpoint.ppt', 'ppt', 'application/vnd.ms-powerpoint'],
      ['visio.vsd', 'vsd', 'application/vnd.visio'],
      ['template.oft', 'oft', 'application/vnd.ms-outlook'],
      ['publisher.pub', 'pub', 'application/x-mspublisher'],
      ['publisher95.pub', 'pub', 'application/x-mspublisher'],
    ];
    for (const [name, format, mimeType] of formats) {
      const r = await inspectOffice(ole(name));
      expect({ name, status: r.status, format: r.format, mimeType: r.mimeType, findings: kinds(r) }).to.deep.equal({ name, status: 'clean', format, mimeType, findings: [] });
    }
    const project = await inspectOffice(ole('project.mpp'));
    expect([project.format, project.mimeType, kinds(project)]).to.deep.equal(['mpp', 'application/vnd.ms-project', ['MACRO/VBA_PROJECT']]);
  });

  it('refuses a compound file of two formats, one of no Office format, and a polyglot with a ZIP archive', async () => {
    const stop = async (name: string) => {
      const r = await inspectOffice(ole(name));
      return [r.status, r.format, kinds(r)];
    };
    expect(await stop('two-formats.doc')).to.deep.equal(['rejected', '', ['PACKAGE/AMBIGUOUS_FORMAT']]);
    expect(await stop('bare-object.bin')).to.deep.equal(['rejected', '', ['PACKAGE/UNRECOGNIZED']]);
    expect(await stop('zip-polyglot.doc')).to.deep.equal(['rejected', '', ['PACKAGE/AMBIGUOUS_FORMAT']]);
  });

  it('reads a ZIP archive whose first bytes are a compound file header that does not read, with the header as leading data', async () => {
    const r = await inspectOffice(ole('broken-cfb-zip.docx'));
    expect([r.status, r.format, r.findings.map(f => [`${f.category}/${f.detail}`, f.data])]).to.deep.equal(['strippable', 'docx', [['CONTAINER/LEADING_BYTES', { bytes: 512 }]]]);
  });

  it('reads LibreOffice exports, which leave an old directory sector in a free sector', async () => {
    for (const name of ['clean.doc', 'clean.xls', 'clean.ppt']) {
      const r = await inspectOffice(fixture('lo', name));
      expect({ name, status: r.status, findings: actionable(r) }).to.deep.equal({ name, status: 'strippable', findings: ['CONTAINER/HIDDEN_DATA'] });
    }
    const links = await inspectOffice(fixture('lo', 'links.doc'));
    expect(actionable(links)).to.deep.equal(['CONTAINER/HIDDEN_DATA', 'MACRO/LINK', 'LINK/IP_HOST', 'LINK/CREDENTIALS']);
    // LibreOffice writes a hyperlink's field code with a 0x01 character after the address, which is not text.
    expect(of(links, 'LINK/SAFE').map(f => f.data?.code)).to.deep.equal(['HYPERLINK "https://example.com/page"']);
    expect(actionable(await inspectOffice(fixture('lo', 'dde.xls')))).to.deep.equal(['CONTAINER/HIDDEN_DATA', 'DDE/LINK']);
  });
});

describe('VBA, Excel 4.0 macros and embedded objects', () => {
  it('reports VBA projects in Word, Excel and Project files, and an Excel project storage without its dir stream', async () => {
    const locations = async (name: string) => of(await inspectOffice(ole(name)), 'MACRO/VBA_PROJECT').map(f => f.location);
    expect(await locations('vba.doc')).to.deep.equal(['Macros/VBA']);
    expect(await locations('workbook-vba.xls')).to.deep.equal(['_VBA_PROJECT_CUR/VBA']);
    expect(await locations('vba-incomplete.xls')).to.deep.equal(['_VBA_PROJECT_CUR']);
  });

  it('reports Excel 4.0 macro sheets by their BoundSheet record or their BOF, and the names Excel runs on open and close', async () => {
    const r = await inspectOffice(ole('xlm.xls'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data])).to.deep.equal([
      ['MACRO/EVENT_ON_OPEN', { event: 'Auto_Open', name: 'Auto_Open' }],
      ['MACRO/EVENT', { event: 'Auto_Close', name: 'auto_close_inert' }],
      ['MACRO/XLM_MACROSHEET', { sheet: 'Macro1', state: 'hidden' }],
    ]);
    expect(of(await inspectOffice(ole('xlm-bof.xls')), 'MACRO/XLM_MACROSHEET').map(f => f.data)).to.deep.equal([{ sheet: 'Macro1', state: 'visible' }]);
    // Excel 97 reads Workbook and Excel 5.0 reads Book, and a macro sheet in either counts.
    expect(of(await inspectOffice(ole('both-streams.xls')), 'MACRO/XLM_MACROSHEET').map(f => f.location)).to.deep.equal(['Book']);
  });

  it('reports WordBasic macros in a Word 6.0 template', async () => {
    const r = await inspectOffice(ole('word6-macros.dot'));
    expect([r.format, r.findings.map(f => [`${f.category}/${f.detail}`, f.data])]).to.deep.equal(['dot', [['MACRO/WORDBASIC', { stream: 'WordDocument', bytes: 64 }]]]);
  });

  it('reports a Packager object with its file, an Equation Editor object, a control and linked objects in a Word ObjectPool', async () => {
    const r = await inspectOffice(ole('objects.doc'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.location])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1000'],
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1001'],
      ['EMBEDDED/EQUATION', 'ObjectPool/_1001'],
      ['CONTROL/ACTIVEX', 'ObjectPool/_1002'],
      ['EXTERNAL/OBJECT', 'ObjectPool/_1003'],
      ['EXTERNAL/LOCAL_FILE', 'ObjectPool/_1004'],
    ]);
    expect(r.findings[0].data).to.deep.include({ fileName: 'inert.txt', sourcePath: 'C:\\inert\\inert.txt', tempPath: 'C:\\Temp\\inert.txt', size: 12, progId: 'Package' });
    expect(r.findings[0].description).to.equal('An embedded OLE object, which holds the file inert.txt');
    expect(r.findings[2].data).to.deep.include({ clsid: '0002CE02-0000-0000-C000-000000000046', progId: 'Equation.3' });
    expect(r.findings[3].data).to.deep.include({ name: 'InertControl1', clsid: 'D27CDB6E-AE6D-11CF-96B8-444553540000' });
    expect(r.findings[4].data?.target).to.equal('https://example.invalid/inert.hta');
    expect(r.findings[5].data?.target).to.equal('C:\\inert\\linked.xls');
    expect(r.risk).to.equal('CRITICAL');
  });

  it('reports an Excel MBD storage and the Ctls stream of ActiveX controls', async () => {
    expect(kinds(await inspectOffice(ole('objects.xls')))).to.deep.equal(['EMBEDDED/OLE_OBJECT', 'CONTROL/ACTIVEX']);
  });

  it('reads PowerPoint objects, compressed or not, by the persist directory, and the actions, links and media of the slides', async () => {
    const r = await inspectOffice(ole('objects.ppt'));
    expect(kinds(r)).to.deep.equal([
      'LINK/SAFE',
      'EXTERNAL/LOCAL_FILE',
      'ACTION/RUN_PROGRAM',
      'MACRO/LINK',
      'ACTION/OLE_VERB',
      'EMBEDDED/OLE_OBJECT',
      'EMBEDDED/EQUATION',
      'MACRO/VBA_PROJECT',
      'CONTROL/ACTIVEX',
    ]);
    expect(of(r, 'ACTION/RUN_PROGRAM')[0].data).to.deep.equal({ trigger: 'click', action: 'C:\\Windows\\inert.exe' });
    expect(of(r, 'MACRO/LINK')[0].data).to.deep.equal({ trigger: 'mouse over', action: 'InertMacro' });
    expect(of(r, 'MACRO/VBA_PROJECT')[0].location).to.match(/^PowerPoint Document, offset \d+ > VBA$/);
  });

  it('reports a PowerPoint object whose data does not inflate or is no compound file', async () => {
    const r = await inspectOffice(ole('broken-object.ppt'));
    expect(r.findings.map(f => f.data?.reason)).to.deep.equal(['its compressed data does not inflate to its declared size', 'its storage does not read: not a compound file']);
  });
});

describe('fields, links and attachments', () => {
  const FIELDS = ['DDE/FIELD', 'EXTERNAL/INCLUDE', 'EXTERNAL/LOCAL_FILE', 'LINK/SAFE', 'FIELD/MACROBUTTON', 'FIELD/COMPUTED'];

  it('reads the fields of a Word document through its piece table, as 16-bit or 8-bit text, and its attached template', async () => {
    expect(kinds(await inspectOffice(ole('fields.doc')))).to.deep.equal([...FIELDS, 'EXTERNAL/TEMPLATE']);
    expect(kinds(await inspectOffice(ole('fields-8bit.doc')))).to.deep.equal(FIELDS);
    const local = await inspectOffice(ole('local-template.doc'));
    expect(local.findings.map(f => [`${f.category}/${f.detail}`, f.data])).to.deep.equal([['EXTERNAL/LOCAL_FILE', { source: 'SttbfAssoc', target: 'C:\\Templates\\Inert.dot', kind: 'template' }]]);
    expect(kinds(await inspectOffice(ole('word6-dde.doc')))).to.deep.equal(['DDE/FIELD']);
  });

  it('reads Excel DDE, OLE and workbook links, shapes that run macros and hyperlinks', async () => {
    const r = await inspectOffice(ole('links.xls'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.target ?? f.data?.application ?? f.data?.sheet])).to.deep.equal([
      ['DDE/LINK', 'inert-app'],
      ['EXTERNAL/LOCAL_FILE', 'C:\\inert\\object.bin'],
      ['EXTERNAL/WORKBOOK', 'https://example.invalid/books\\inert.xls'],
      ['EXTERNAL/LOCAL_FILE', 'C:\\inert\\local.xls'],
      ['EXTERNAL/LOCAL_FILE', 'relative.xls'],
      ['MACRO/LINK', 'Sheet1'],
      ['LINK/SAFE', 'https://example.invalid/page'],
      ['LINK/FILE_URL', 'C:\\inert\\file.txt'],
    ]);
  });

  it('reports each attachment of a message, an embedded message and its own attachments, and attachments by reference', async () => {
    const r = await inspectOffice(ole('message.msg'));
    expect(r.format).to.equal('msg');
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.name ?? f.data?.progId, f.data?.type ?? f.data?.target])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'Inert.Object', undefined],
      ['EMBEDDED/ATTACHMENT', 'inert.txt', 'text/plain'],
      ['EXTERNAL/OTHER', 'linked.txt', '\\\\inert.example.invalid\\share\\linked.txt'],
      ['EXTERNAL/OTHER', 'web.txt', 'https://example.invalid/web.txt'],
      ['EMBEDDED/ATTACHMENT', 'Forwarded', 'application/vnd.ms-outlook'],
      ['EMBEDDED/ATTACHMENT', 'nested.txt', 'text/plain'],
      ['EMBEDDED/ATTACHMENT', 'object', undefined],
    ]);
    expect(r.findings[1].description).to.equal('An attached file "inert.txt" (text/plain)');
    expect(r.findings[1].data).to.deep.include({ method: 'by value', size: 12 });
  });
});

describe('encryption', () => {
  it('opens encrypted packages and workbooks with VelvetSweatshop, and reads what they hold', async () => {
    const agile = await inspectOffice(ole('agile.xlsx'));
    expect([agile.status, agile.format, agile.findings.map(f => [`${f.category}/${f.detail}`, f.action])]).to.deep.equal([
      'strippable',
      'xlsx',
      [
        ['ENCRYPTED/DEFAULT_PASSWORD', 'strip'],
        ['ENCRYPTED/AES_256', 'info'],
      ],
    ]);
    const standard = await inspectOffice(ole('standard.docx'));
    expect([standard.format, kinds(standard)]).to.deep.equal(['docx', ['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/AES_128']]);
    expect(kinds(await inspectOffice(ole('rc4.xls')))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/RC4_40', 'MACRO/EVENT_ON_OPEN', 'MACRO/XLM_MACROSHEET']);
    expect(kinds(await inspectOffice(ole('cryptoapi.xls')))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/RC4_128', 'DDE/LINK']);
  });

  it('opens a file with the password the caller passes, and asks for one when no password at hand opens it', async () => {
    const xls = await inspectOffice(ole('rc4-password.xls'), { password: 'inert-password' });
    // As in pdf-defuse, a rewrite drops the encryption the supplied password opens, so the action is strip at weight 0.
    expect([xls.status, xls.score, xls.findings.map(f => [`${f.category}/${f.detail}`, f.action])]).to.deep.equal([
      'strippable',
      0,
      [
        ['ENCRYPTED/USER_PASSWORD', 'strip'],
        ['ENCRYPTED/RC4_40', 'info'],
      ],
    ]);
    expect(kinds(await inspectOffice(ole('password.docx'), { password: 'inert-password' }))).to.deep.equal(['ENCRYPTED/USER_PASSWORD', 'ENCRYPTED/AES_256']);
    // A wrong password falls back to VelvetSweatshop.
    expect(kinds(await inspectOffice(ole('standard.docx'), { password: 'wrong' }))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/AES_128']);
    const stopped: Array<[string, { password?: string }, string]> = [
      ['password.docx', {}, 'the package is encrypted with a password that is not VelvetSweatshop'],
      ['password.docx', { password: 'wrong' }, 'the package is encrypted with a password that is not VelvetSweatshop'],
      ['rc4-password.xls', {}, 'the workbook is encrypted with a password that is not VelvetSweatshop'],
      ['word-password.doc', {}, 'The document is encrypted'],
      ['presentation-password.ppt', {}, 'The presentation is encrypted'],
      ['filepass-cut.xls', {}, 'a FILEPASS record cut short'],
      ['package-cut.docx', {}, 'an EncryptedPackage stream cut short'],
      ['package-short.docx', {}, 'an EncryptedPackage stream shorter than the size it gives'],
      ['standard-cut.docx', {}, 'a Standard EncryptionInfo cut short'],
      ['agile-no-key.docx', {}, 'an Agile descriptor without key data or a password key'],
      ['agile-spin.docx', {}, 'an Agile spin count past 10,000,000'],
    ];
    for (const [name, options, reason] of stopped) {
      const r = await inspectOffice(ole(name), options);
      expect({ name, status: r.status, risk: r.risk, kinds: kinds(r), reason: r.findings[0].data?.reason }).to.deep.equal({
        name,
        status: 'rejected',
        risk: 'UNKNOWN',
        kinds: ['ENCRYPTED/PASSWORD_REQUIRED'],
        reason,
      });
    }
  });

  it('refuses a file encrypted with an algorithm it does not decrypt, and never hands such a name to OpenSSL', async () => {
    const algorithms: Array<[string, string]> = [
      ['xor.xls', 'XOR obfuscation'],
      ['extensible.docx', 'encryption version 4.3'],
      ['info-cut.docx', 'encryption version 0.0'],
      ['filepass-version.xls', 'encryption version 9.9'],
      ['cryptoapi-aes.xls', 'CryptoAPI algorithm 0x660e with a 128-bit key'],
      ['standard-rc4.docx', 'Standard encryption with algorithm 0x6801, hash 0x8004 and a 128-bit key'],
      ['agile-cipher.docx', 'Agile encryption with DES in ChainingModeCBC, a 256-bit key and SHA512'],
      ['agile-ripemd.docx', 'Agile encryption with AES in ChainingModeCBC, a 256-bit key and RIPEMD160'],
      ['agile-constructor.docx', 'Agile encryption with AES in ChainingModeCBC, a 256-bit key and constructor'],
      ['agile-no-bits.docx', 'Agile encryption with AES in ChainingModeCBC, a none-bit key and SHA512'],
      ['standard-bits.docx', 'Standard encryption with algorithm 0x660e, hash 0x8004 and a 64-bit key'],
      ['standard-sha256.docx', 'Standard encryption with algorithm 0x660e, hash 0x800c and a 128-bit key'],
    ];
    for (const [name, algorithm] of algorithms) {
      const r = await inspectOffice(ole(name));
      expect({ name, status: r.status, kinds: kinds(r), algorithm: r.findings[0].data?.algorithm }).to.deep.equal({ name, status: 'rejected', kinds: ['ENCRYPTED/UNSUPPORTED_ALGORITHM'], algorithm });
    }
  });

  it('reads RC4 CryptoAPI keys of 40 bits, given as 40 or as 0, and of other sizes, Agile with SHA-1, and data after a decrypted archive', async () => {
    expect(kinds(await inspectOffice(ole('cryptoapi-40.xls')))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/RC4_40']);
    const other = await inspectOffice(ole('cryptoapi-56.xls'));
    expect(other.findings.map(f => [`${f.category}/${f.detail}`, f.description])).to.deep.equal([
      ['ENCRYPTED/DEFAULT_PASSWORD', 'The file is encrypted with VelvetSweatshop, the password Excel opens it with without asking'],
      ['ENCRYPTED/RC4_OTHER', 'Encryption is RC4 with a 56-bit key'],
    ]);
    expect(kinds(await inspectOffice(ole('agile-sha1.docx')))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/AES_256']);
    // Without the key data's salt, Agile segments decrypt with the wrong IVs, and the archive does not read.
    expect(kinds(await inspectOffice(ole('agile-no-salt.docx')))).to.deep.equal(['ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/AES_256', 'CONTAINER/HEADER_MISMATCH']);
    const trailing = await inspectOffice(ole('agile-trailing.docx'));
    expect(trailing.findings.map(f => [`${f.category}/${f.detail}`, f.location])).to.deep.equal([
      ['ENCRYPTED/DEFAULT_PASSWORD', 'EncryptedPackage'],
      ['ENCRYPTED/AES_256', 'EncryptedPackage'],
      ['CONTAINER/TRAILING_BYTES', 'EncryptedPackage'],
    ]);
  });
});

describe('less common records and storages', () => {
  it('refuses a WordDocument stream without a FIB, and reads a FIB or piece table cut short, the 0Table stream and formatting before the pieces', async () => {
    for (const name of ['short-fib.doc', 'not-fib.doc']) expect(kinds(await inspectOffice(ole(name)))).to.deep.equal(['PACKAGE/UNRECOGNIZED']);
    for (const name of ['fib-cut.doc', 'clx-other.doc', 'clx-cut.doc', 'assoc-one.doc', 'assoc-cut.doc'])
      expect({ name, kinds: kinds(await inspectOffice(ole(name))) }).to.deep.equal({ name, kinds: [] });
    for (const name of ['0table.doc', 'clx-variants.doc']) expect(kinds(await inspectOffice(ole(name)))).to.deep.equal(['DDE/FIELD']);
    expect(of(await inspectOffice(ole('assoc-8bit.doc')), 'EXTERNAL/LOCAL_FILE')[0].data?.target).to.equal('C:\\Templates\\8bit');
  });

  it('reads object storages without object streams, with a CompObj past 64 KiB, Packager links, a control without a name and a link without a moniker', async () => {
    const r = await inspectOffice(ole('objects-variants.doc'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.location, f.data?.target ?? f.data?.fileName ?? f.data?.clsid])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1000', undefined],
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1001', undefined],
      ['EXTERNAL/LOCAL_FILE', 'ObjectPool/_1002', 'C:\\inert\\linked.txt'],
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1002', 'linked.txt'],
      ['CONTROL/ACTIVEX', 'ObjectPool/_1003', undefined],
      ['EXTERNAL/OBJECT', 'ObjectPool/_1004', '00000300-0000-0000-C000-000000000046'],
      ['EMBEDDED/OLE_OBJECT', 'ObjectPool/_1005', 'linked.txt'],
    ]);
  });

  it('refuses a workbook stream without a BOF, and reads an OLE link without a file, sheets past their BOF, charts in sheets, Obj records without macros and a record cut short', async () => {
    expect(kinds(await inspectOffice(ole('no-bof.xls')))).to.deep.equal(['PACKAGE/UNRECOGNIZED']);
    const r = await inspectOffice(ole('xls-variants.xls'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data])).to.deep.equal([['EXTERNAL/OBJECT', { source: 'SupBook', className: 'Inert.Class' }]]);
  });

  it('reads PowerPoint objects whose data is cut short, controls without a class id, storages that do not read, and records and persist entries out of place', async () => {
    const r = await inspectOffice(ole('ppt-variants.ppt'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.clsid ?? f.data?.reason])).to.deep.equal([
      ['EMBEDDED/OLE_OBJECT', 'its compressed data does not inflate to its declared size'],
      ['CONTROL/ACTIVEX', undefined],
      ['CONTROL/ACTIVEX', 'its storage does not read: not a compound file'],
      ['MACRO/VBA_PROJECT', 'its compressed data does not inflate to its declared size'],
    ]);
    expect(kinds(await inspectOffice(ole('huge-object.ppt')))).to.deep.equal(['LIMIT/DECOMPRESSED_SIZE']);
    for (const name of ['bad-edit.ppt', 'record-past-end.ppt']) expect(kinds(await inspectOffice(ole(name)))).to.deep.equal([]);
    // A nested storage over the entry limit stops the run as the file's own would.
    expect(kinds(await inspectOffice(ole('objects.ppt'), { limits: { entries: 6 } })).pop()).to.equal('LIMIT/ENTRY_COUNT');
  });

  it('names attachments from the short name or the display name, and reads attachments without properties or data', async () => {
    const r = await inspectOffice(ole('msg-variants.msg'));
    expect(r.findings.map(f => [`${f.category}/${f.detail}`, f.data?.name ?? f.data?.target, f.data?.method])).to.deep.equal([
      ['EMBEDDED/ATTACHMENT', 'short.txt', 'by value'],
      ['EMBEDDED/ATTACHMENT', 'Display', '0'],
      ['EXTERNAL/LOCAL_FILE', 'C:\\inert\\ref.txt', undefined],
      ['EMBEDDED/ATTACHMENT', undefined, '0'],
      ['EXTERNAL/OTHER', 'https://example.invalid/w', undefined],
    ]);
  });
});
