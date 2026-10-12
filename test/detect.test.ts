import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { expect } from 'chai';
import { type ByteSource, type ContainedFile, detectOffice, detectOfficeSource, type FormatDetection, inspectOffice, officePlugin } from '../src';
import { type CfbEntry, CompoundFile } from '../src/cfb';
import { type CfbNode, writeCfb } from '../src/cfb-writer';
import { ODF_FORMATS, OOXML_FORMATS, XPS_FORMATS, XPS_SEQUENCE } from '../src/formats';
import { bufferSource } from '../src/io';
import { XPS_MAIN } from '../src/ooxml';
import {
  bufferSink,
  contentTypesXml,
  failure,
  fixture,
  fixtures,
  flatOdf,
  flatOpc,
  flatWord,
  kinds,
  must,
  odfPackage,
  partRels,
  RELS,
  relsXml,
  T,
  textBody,
  wordPackage,
  XML,
  zipOf,
} from './helpers/util';

/** Every file under test/fixtures, as its path there. */
const allFixtures = fs.readdirSync(fixtures).flatMap(dir =>
  fs
    .readdirSync(path.join(fixtures, dir))
    .filter(f => fs.statSync(path.join(fixtures, dir, f)).isFile())
    .map(f => `${dir}/${f}`),
);

const claim = (f: { format: string; mimeType: string }, offset = 0): FormatDetection => ({ mimeType: f.mimeType, extension: f.format, offset });

/** Detects `bytes`, and checks that the inspection names the format detection names, or none when detection claims nothing. */
async function agreed(bytes: Uint8Array, what: string): Promise<FormatDetection | undefined> {
  const d = await detectOffice(bytes);
  const r = await inspectOffice(bytes);
  // An encrypted package no password at hand opens is claimed as a compound file. The inspection names no format and
  // rejects it as encrypted.
  const locked = !r.format && r.findings.some(f => f.category === 'ENCRYPTED' && f.action === 'reject');
  const expected = r.format ? { mimeType: r.mimeType, extension: r.format } : locked ? { mimeType: 'application/x-ole-storage', extension: '' } : undefined;
  expect(d && { mimeType: d.mimeType, extension: d.extension }, what).to.deep.equal(expected);
  return d;
}

/** A compound file with `edit` applied to the streams at its root, written again by the compound file writer. */
async function editRoot(bytes: Uint8Array, edit: (name: string, data: Buffer) => Buffer): Promise<Buffer> {
  const cfb = await CompoundFile.open(bufferSource(bytes), { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 });
  const node = async (e: CfbEntry, top: boolean): Promise<CfbNode> => {
    if (e.kind !== 'stream') return { kind: 'storage', name: e.name, clsid: e.clsid, children: await Promise.all(e.children.map(c => node(c, false))) };
    const data = Buffer.from(await cfb.bytes(e));
    const out = top ? edit(e.name, data) : data;
    return { kind: 'stream', name: e.name, size: out.length, source: bufferSource(out) };
  };
  const sink = bufferSink();
  await writeCfb({ clsid: cfb.root.clsid, children: await Promise.all(cfb.root.children.map(c => node(c, true))) }, sink);
  return sink.result();
}

/** A workbook stream with a Template record after its FILEPASS record. */
function withTemplateRecord(stream: Buffer): Buffer {
  for (let at = 0; at + 4 <= stream.length; at += 4 + stream.readUInt16LE(at + 2)) {
    if (stream.readUInt16LE(at) !== 0x002f) continue;
    const end = at + 4 + stream.readUInt16LE(at + 2);
    return Buffer.concat([stream.subarray(0, end), Buffer.from([0x60, 0x00, 0x00, 0x00]), stream.subarray(end)]);
  }
  throw new Error('no FILEPASS record');
}

/** A minimal XPS or OpenXPS package, whose FixedDocumentSequence the package relationship of `kind` names. */
const xpsPackage = (kind: 'xps' | 'oxps') =>
  zipOf([
    ['[Content_Types].xml', contentTypesXml({ 'Doc.fdseq': XPS_SEQUENCE })],
    ['_rels/.rels', relsXml(['R0', XPS_MAIN[kind], 'Doc.fdseq'])],
    ['Doc.fdseq', `${XML}<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06"/>`],
  ]);

describe('detectOffice', () => {
  it('names every fixture the inspection names, with its format and MIME type, and claims no other', async () => {
    const seen = new Set<string>();
    for (const name of allFixtures) {
      const d = await agreed(fixture(name), name);
      if (d) seen.add(d.extension);
    }
    expect([...seen]).to.include.members(['doc', 'dot', 'xls', 'xlt', 'xla', 'ppt', 'vsd', 'msg', 'oft', 'pub', 'mpp', 'rtf']);
    expect([...seen]).to.include.members(['docx', 'docm', 'dotx', 'dotm', 'xlsx', 'xlsm', 'xltx', 'xlam', 'xlsb', 'pptx', 'potx', 'ppsm', 'ppam', 'vsdx', 'vsdm', 'xps']);
    expect([...seen]).to.include.members(['odt', 'ott', 'ods', 'ots', 'odp', 'otp', 'odg', 'otg', 'odf', 'odm', 'odb']);
  });

  it('names every OOXML, XPS and OpenDocument format by the markers the inspection reads', async () => {
    for (const [type, f] of OOXML_FORMATS) expect(await agreed(await wordPackage('<w:p/>', { main: type }), type)).to.deep.equal(claim(f));
    // XPS and OpenXPS share their content types, so the relationship type tells them apart.
    for (const kind of ['xps', 'oxps'] as const) expect(await agreed(await xpsPackage(kind), kind)).to.deep.equal(claim(XPS_FORMATS[kind]));
    for (const [type, f] of ODF_FORMATS) expect(await agreed(await odfPackage(textBody(''), { mimetype: type }), type)).to.deep.equal(claim(f));
    // Without a mimetype entry, the manifest names the format.
    const manifestOnly = await odfPackage(textBody(''), { mimetype: null, manifestType: 'application/vnd.oasis.opendocument.spreadsheet' });
    expect(await agreed(manifestOnly, 'manifest only')).to.deep.equal(claim({ format: 'ods', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' }));
  });

  it('names flat OpenDocument by its root, and Flat OPC once the package its parts make names a format', async () => {
    for (const [media, format] of [
      ['text', 'fodt'],
      ['spreadsheet', 'fods'],
      ['presentation', 'fodp'],
      ['graphics', 'fodg'],
    ]) {
      expect(await agreed(flatOdf('<office:body/>', media), media)).to.deep.equal(claim({ format, mimeType: `application/vnd.oasis.opendocument.${media}-flat-xml` }));
    }
    const flatPackage = claim({ format: 'xml', mimeType: 'application/xml' });
    expect(await agreed(flatWord('<w:p/>'), 'Flat OPC')).to.deep.equal(flatPackage);
    // Package relationships held as binary data are read as the inspection reads them.
    const binaryRels = flatOpc([
      { name: '/_rels/.rels', type: RELS, binary: Buffer.from(relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])) },
      {
        name: '/word/document.xml',
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        xml: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      },
    ]);
    expect(await agreed(binaryRels, 'binary package relationships')).to.deep.equal(flatPackage);
  });

  it('claims an encrypted package no password at hand opens as a compound file, which the password names later', async () => {
    expect(await detectOffice(fixture('ole', 'password.docx'))).to.deep.equal({ mimeType: 'application/x-ole-storage', extension: '', offset: 0 });
    expect((await detectOffice(fixture('ole', 'agile.xlsx')))?.extension).to.equal('xlsx');
  });

  it('does not claim a file whose format it cannot name', async () => {
    const installer = bufferSink();
    await writeCfb(
      {
        clsid: '000C1084-0000-0000-C000-000000000046',
        children: [
          { kind: 'stream', name: '\x05SummaryInformation', size: 12, source: bufferSource(Buffer.from('inert marker')) },
          { kind: 'stream', name: 'InertTable', size: 12, source: bufferSource(Buffer.from('inert marker')) },
        ],
      },
      installer,
    );
    const manifestRel = 'http://schemas.microsoft.com/packaging/2010/07/manifest';
    const unnamed: Array<[string, Uint8Array]> = [
      ['a ZIP archive without an Office marker', fixture('cases', 'plain-archive.zip')],
      [
        'a package with [Content_Types].xml and no Office main part, as a NuGet package has',
        await zipOf([
          ['[Content_Types].xml', contentTypesXml({}, { nuspec: 'application/octet', rels: 'application/vnd.openxmlformats-package.relationships+xml' })],
          ['_rels/.rels', relsXml(['R0', manifestRel, '/Inert.nuspec'])],
          ['Inert.nuspec', `${XML}<package/>`],
        ]),
      ],
      ['an OOXML package whose main part has a content type of no Office format', await wordPackage('<w:p/>', { main: 'application/vnd.inert.main+xml' })],
      ['a StarOffice package, whose media type is no OpenDocument one', fixture('cases', 'sun-xml-calc.ods')],
      ['a compound file of another kind, with an installer class id', installer.result()],
      ['a bare OLE object', fixture('ole', 'bare-object.bin')],
      ['a compound file of two Office formats', fixture('ole', 'two-formats.doc')],
      ['a file that is a compound file and a ZIP archive at once', fixture('ole', 'zip-polyglot.doc')],
      ['a file that is RTF and a ZIP archive at once', fixture('rtf', 'zip-polyglot.rtf')],
      ['XML of another kind', Buffer.from(`${XML}<svg xmlns="http://www.w3.org/2000/svg"/>`)],
      [
        'a Flat OPC package of no Office format',
        flatOpc([
          { name: '/_rels/.rels', type: RELS, xml: partRels(['R0', manifestRel, 'inert.xml']) },
          { name: '/inert.xml', type: 'application/xml', xml: '<inert/>' },
        ]),
      ],
      ['a flat OpenDocument file of a media type it does not read', flatOdf('<office:body/>', 'chart')],
      ['a PDF', Buffer.from('%PDF-1.7\n1 0 obj\n<<>>\nendobj\n')],
      ['an empty file', Buffer.alloc(0)],
    ];
    for (const [what, bytes] of unnamed) expect(await agreed(bytes, what), what).to.equal(undefined);
  });

  it('claims a file it names that it cannot read further on, which the inspection then rejects', async () => {
    const damaged: Array<[string, Uint8Array, string]> = [
      ['a WordDocument stream without a Word FIB', fixture('cases', 'ole-other-word-identifier.doc'), 'doc'],
      ['a Workbook stream without a BOF record', fixture('ole', 'no-bof.xls'), 'xls'],
      ['an entry whose CRC is wrong', fixture('cases', 'wrong-crc.docx'), 'docx'],
      ['an empty main part', fixture('cases', 'empty-main-part.ppam'), 'ppam'],
      ['an encrypted OpenDocument package', fixture('docs', 'encrypted.odt'), 'odt'],
      ['a flat OpenDocument file cut short', flatOdf(textBody('<text:p>x</text:p>')).subarray(0, -30), 'fodt'],
      [
        'a Flat OPC package whose main part is not XML',
        flatOpc([
          { name: '/_rels/.rels', type: RELS, xml: partRels(['rId1', `${T}officeDocument`, 'word/document.xml']) },
          { name: '/word/document.xml', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', binary: Buffer.from('<w:document') },
        ]),
        'xml',
      ],
    ];
    for (const [what, bytes, extension] of damaged) {
      expect((await agreed(bytes, what))?.extension, what).to.equal(extension);
      expect((await inspectOffice(bytes)).status, what).to.equal('rejected');
    }
  });

  it('names a template, and an encrypted Word or Excel template, without a password, as the inspection does', async () => {
    const dot = await editRoot(fixture('ole', 'word-password.doc'), (name, data) => {
      if (name === 'WordDocument') data[10] |= 1;
      return data;
    });
    const xlt = await editRoot(fixture('ole', 'rc4-password.xls'), (name, data) => (name === 'Workbook' ? withTemplateRecord(data) : data));
    for (const [bytes, extension] of [
      [dot, 'dot'],
      [xlt, 'xlt'],
    ] as const) {
      expect((await agreed(bytes, extension))?.extension).to.equal(extension);
      // The password the file needs is not at hand, and the inspection stops there, having named the format.
      expect(kinds(await inspectOffice(bytes))).to.include('ENCRYPTED/PASSWORD_REQUIRED');
    }
    // An encrypted OOXML package names its format once VelvetSweatshop opens it.
    expect((await agreed(fixture('ole', 'agile.xlsx'), 'agile.xlsx'))?.extension).to.equal('xlsx');
  });

  it('gives the offset where the ZIP data starts when other data comes first, as the ZIP reader finds it', async () => {
    for (const name of ['zip/leading-bytes.docx', 'ole/broken-cfb-zip.docx']) {
      const bytes = fixture(name);
      const leading = must(
        (await inspectOffice(bytes)).findings.find(f => f.detail === 'LEADING_BYTES'),
        `${name} leading bytes`,
      ).data?.bytes;
      expect((await detectOffice(bytes))?.offset, name)
        .to.equal(leading)
        .and.to.be.above(0);
    }
    expect((await detectOffice(fixture('lo', 'clean.docx')))?.offset).to.equal(0);
  });

  it('decides by the content alone, whatever the name and declared type say', async () => {
    const docx = fixture('lo', 'clean.docx');
    expect(await detectOffice(docx, { name: 'report.pdf', declaredType: 'application/pdf' })).to.deep.equal(await detectOffice(docx));
    const word = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    expect(await detectOffice(fixture('cases', 'plain-archive.zip'), { name: 'report.docx', declaredType: word })).to.equal(undefined);
  });

  it('leaves the source open, stops at the deadline, passes I/O errors up and cleans up its temporary files', async () => {
    let closed = 0;
    const inner = bufferSource(fixture('lo', 'clean.docx'));
    const source: ByteSource = { size: () => inner.size(), read: (o, n) => inner.read(o, n), close: async () => void closed++ };
    expect((await detectOfficeSource(source))?.extension).to.equal('docx');
    expect(closed).to.equal(0);
    for (const name of ['lo/clean.docx', 'lo/clean.rtf', 'ole/word.doc']) {
      const e = await failure(detectOffice(fixture(name), { deadline: Date.now() - 1 }));
      expect((e as { code?: string })?.code, name).to.equal('DEFUSE_LIMIT');
    }
    const broken = Object.assign(new Error('the disk failed'), { code: 'EIO' });
    const failing: ByteSource = { size: () => inner.size(), read: async () => Promise.reject(broken) };
    expect(await failure(detectOfficeSource(failing))).to.equal(broken);
    // A Flat OPC package is packed into a ZIP archive to be named, which spills to the temporary folder.
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-detect-'));
    try {
      expect((await detectOffice(flatWord('<w:p/>'), { tempDir, memoryThreshold: 0 }))?.extension).to.equal('xml');
      expect(fs.readdirSync(tempDir)).to.deep.equal([]);
    } finally {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it('lets officePlugin take exactly what detectOffice claims', async () => {
    const plugin = officePlugin();
    const file = (bytes: Uint8Array, name: string): ContainedFile => ({ name, size: bytes.length, location: 'x', depth: 1, source: bufferSource(bytes) });
    const built = [flatWord('<w:p/>'), flatOdf('<office:body/>'), flatOdf('<office:body/>', 'chart'), Buffer.from(`${XML}<svg/>`)];
    for (const [name, bytes] of [...allFixtures.map(n => [n, fixture(n)] as const), ...built.map((b, i) => [`built ${i}`, b] as const)]) {
      expect(await plugin.accepts(file(bytes, name)), name).to.equal((await detectOffice(bytes)) !== undefined);
    }
  });
});
