import * as zlib from 'node:zlib';
import { expect } from 'chai';
import { type ContainedFile, type ContainedFilePlugin, disarmOffice, inspectOffice, officePlugin } from '../src';
import { crc32 } from '../src/crc32';
import { bufferSource } from '../src/io';
import { flatOdf, flatOpc, flatWord, kinds, must, OLE, of, partRels, RELS, T, W } from './helpers/util';

/** A 1 x 1 gray PNG. */
function png(): Buffer {
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.from([0, 0, 0, 1, 0, 0, 0, 1, 8, 0, 0, 0, 0]);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.from([0, 0x80]))), chunk('IEND', Buffer.alloc(0))]);
}

const textBody = (inner: string) => `<office:body><office:text>${inner}</office:text></office:body>`;
const picture = (data: Buffer) => `<draw:frame draw:name="p"><draw:image draw:mime-type="image/png"><office:binary-data>${data.toString('base64')}</office:binary-data></draw:image></draw:frame>`;

/** A plugin that takes every file and answers `answer`, writing `bytes` when given. It records what it saw. */
function recorder(answer: 'passed' | 'scrubbed' | 'removed', bytes?: Uint8Array) {
  const seen: ContainedFile[] = [];
  const plugin: ContainedFilePlugin = {
    kind: 'file',
    name: 'test',
    accepts: () => true,
    async process(file, sink) {
      seen.push(file);
      if (bytes) await sink.write(bytes);
      return answer;
    },
  };
  return { plugin, seen };
}

const text = (b: Uint8Array | undefined) => Buffer.from(must(b, 'output')).toString('utf8');

describe('flat OpenDocument', () => {
  it('reads each flat format by its media type', async () => {
    for (const [kind, ext] of [
      ['text', 'fodt'],
      ['spreadsheet', 'fods'],
      ['presentation', 'fodp'],
      ['graphics', 'fodg'],
    ]) {
      const r = await inspectOffice(flatOdf('<office:body/>', kind));
      expect([r.status, r.format, r.mimeType]).to.deep.equal(['clean', ext, `application/vnd.oasis.opendocument.${kind}-flat-xml`]);
    }
    expect(kinds(await inspectOffice(flatOdf('<office:body/>', 'chart')))).to.deep.equal(['PACKAGE/UNRECOGNIZED']);
    expect(kinds(await inspectOffice(Buffer.from('<?xml version="1.0"?><html><body/></html>')))).to.deep.equal(['PACKAGE/UNRECOGNIZED']);
  });

  it('removes Basic and event listeners, and the output is flat OpenDocument again', async () => {
    const basic =
      '<office:script script:language="ooo:Basic"><ooo:libraries xmlns:ooo="http://openoffice.org/2004/office"><ooo:library-embedded ooo:name="Standard"><ooo:module ooo:name="M"><ooo:source-code>Sub Main\nEnd Sub</ooo:source-code></ooo:module></ooo:library-embedded></ooo:libraries></office:script>';
    const events =
      '<office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:load" xlink:href="vnd.sun.star.script:Standard.M.Main?language=Basic&amp;location=document"/></office:event-listeners>';
    const input = flatOdf(`<office:scripts>${basic}${events}</office:scripts>${textBody('<text:p>kept</text:p>')}`);
    const r = await disarmOffice(input);
    expect([r.status, kinds(r.before)]).to.deep.equal(['defused', ['MACRO/ODF_BASIC', 'MACRO/EVENT_ON_OPEN']]);
    const out = text(r.bytes);
    expect([out.includes('Sub Main'), out.includes('vnd.sun.star.script'), out.includes('<text:p>kept</text:p>'), kinds(must(r.after, 'after')), r.after?.format]).to.deep.equal([
      false,
      false,
      true,
      [],
      'fodt',
    ]);
  });

  it('reads every relative reference as a file beside the document', async () => {
    const r = await inspectOffice(
      flatOdf(
        textBody(
          '<draw:frame><draw:image xlink:href="Pictures/a.png"/></draw:frame><draw:frame><draw:image xlink:href="http://example.com/b.png"/></draw:frame><text:a xlink:href="javascript:x()">x</text:a>',
        ),
      ),
    );
    expect(kinds(r)).to.deep.equal(['EXTERNAL/LOCAL_FILE', 'EXTERNAL/IMAGE', 'LINK/SCRIPT_URL']);
    // A chart's ".." names the document that holds it, and a vnd.sun.star.Package: URL the document's own storage.
    const inside = await inspectOffice(
      flatOdf(
        textBody(
          '<draw:frame><draw:object><office:document office:mimetype="application/vnd.oasis.opendocument.chart"><office:body><office:chart><chart:chart xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xlink:href=".."/></office:chart></office:body></office:document></draw:object></draw:frame><draw:frame><draw:image xlink:href="vnd.sun.star.Package:Pictures/c.png"/></draw:frame>',
        ),
      ),
    );
    expect(kinds(inside)).to.deep.equal([]);
  });

  it('shows pictures held as base64 to the plugins, and applies their verdicts', async () => {
    const image = png();
    const input = flatOdf(textBody(`<text:p>${picture(image)}</text:p>`));
    const passed = recorder('passed');
    const kept = await disarmOffice(input, { filePlugins: [passed.plugin] });
    expect([passed.seen.map(f => [f.declaredType, f.sniffedType, f.size, f.depth]), kept.status, kinds(must(kept.after, 'after'))]).to.deep.equal([
      [['image/png', 'image/png', image.length, 1]],
      'clean',
      ['EMBEDDED/PLUGIN_PASSED'],
    ]);
    expect(text(kept.bytes)).to.include(image.toString('base64'));
    const removed = await disarmOffice(input, { filePlugins: [recorder('removed').plugin] });
    expect([removed.status, text(removed.bytes).includes('draw:image'), text(removed.bytes).includes('<draw:frame')]).to.deep.equal(['defused', false, true]);
    // A plugin's copy goes back in as base64, and the check of the output finds it there.
    const copy = Buffer.concat([image, Buffer.from('rewritten')]);
    const scrubbed = await disarmOffice(input, { filePlugins: [recorder('scrubbed', copy).plugin] });
    expect([scrubbed.status, text(scrubbed.bytes).includes(copy.toString('base64')), kinds(must(scrubbed.after, 'after'))]).to.deep.equal(['defused', true, ['EMBEDDED/PLUGIN_SCRUBBED']]);
    // With no plugins, a picture stays, since no rule names it.
    expect((await disarmOffice(input)).status).to.equal('clean');
  });

  it('removes an OLE object held as base64 unless a plugin keeps it', async () => {
    const input = flatOdf(textBody(`<draw:frame><draw:object-ole><office:binary-data>${OLE.toString('base64')}</office:binary-data></draw:object-ole>${picture(png())}</draw:frame>`));
    const r = await disarmOffice(input);
    expect([r.status, kinds(r.before), text(r.bytes).includes('object-ole'), text(r.bytes).includes('draw:image')]).to.deep.equal(['defused', ['EMBEDDED/OLE_OBJECT'], false, true]);
    const kept = await inspectOffice(input, { filePlugins: [recorder('passed').plugin] });
    // The plugin passes the object and its picture.
    expect([of(kept, 'EMBEDDED/OLE_OBJECT').length, of(kept, 'EMBEDDED/PLUGIN_PASSED').length]).to.deep.equal([0, 2]);
  });

  it('writes out internal entities when the internal subset goes', async () => {
    const input = flatOdf(textBody('<text:p>&greeting;</text:p>'), 'text', '<!DOCTYPE office:document [<!ENTITY greeting "hello">]>');
    const r = await disarmOffice(input);
    expect([r.status, kinds(r.before)]).to.deep.equal(['defused', ['XML/INTERNAL_SUBSET']]);
    expect([text(r.bytes).includes('DOCTYPE'), text(r.bytes).includes('<text:p>hello</text:p>')]).to.deep.equal([false, true]);
  });
});

describe('Flat OPC', () => {
  it('reads a flat Word document as a package of its parts', async () => {
    const r = await inspectOffice(flatWord('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
    expect([r.status, r.format, r.mimeType, kinds(r)]).to.deep.equal(['clean', 'xml', 'application/xml', []]);
    // A part may use a namespace that the package element declares.
    const inherited = flatOpc(
      [
        { name: '/_rels/.rels', type: RELS, xml: partRels(['rId1', `${T}officeDocument`, 'word/document.xml']) },
        { name: '/word/document.xml', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', xml: '<w:document><w:body/></w:document>' },
      ],
      ` xmlns:w="${W}"`,
    );
    expect(kinds(await inspectOffice(inherited))).to.deep.equal([]);
  });

  it('removes a DDE field and a VBA project, and writes the package back as one XML file', async () => {
    const field =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DDEAUTO cmd "/c calc"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const image = png();
    const input = flatWord(
      `${field}<w:p><w:r><w:t>kept</w:t></w:r></w:p>`,
      [
        ['rId2', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin'],
        ['rId3', `${T}image`, 'media/image1.png'],
        ['rId4', `${T}attachedTemplate`, 'file:///C:/evil.dotm', true],
      ],
      [
        { name: '/word/vbaProject.bin', type: 'application/vnd.ms-office.vbaProject', binary: OLE },
        { name: '/word/media/image1.png', type: 'image/png', binary: image },
      ],
    );
    const r = await disarmOffice(input);
    expect([r.status, kinds(r.before)]).to.deep.equal(['defused', ['DDE/FIELD', 'EXTERNAL/LOCAL_FILE', 'MACRO/VBA_PROJECT']]);
    const out = text(r.bytes);
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n<pkg:package')).to.equal(true);
    expect([
      out.includes('DDEAUTO'),
      out.includes('vbaProject'),
      out.includes('evil.dotm'),
      out.includes('<w:t>result</w:t>'),
      out.includes('<w:t>kept</w:t>'),
      out.replace(/\n/g, '').includes(image.toString('base64')),
    ]).to.deep.equal([false, false, false, true, true, true]);
    expect([kinds(must(r.after, 'after')), r.after?.format]).to.deep.equal([[], 'xml']);
  });

  it('refuses parts that share a name or have none', async () => {
    const twice = flatWord('<w:p/>', [], [{ name: '/word/document.xml', type: 'application/xml', xml: '<x/>' }]);
    expect(kinds(await inspectOffice(twice))).to.deep.equal(['CONTAINER/DUPLICATE_NAME']);
    const nameless = flatWord('<w:p/>', [], [{ name: '', type: 'application/xml', xml: '<x/>' }]);
    expect(kinds(await inspectOffice(nameless))).to.deep.equal(['PACKAGE/UNRECOGNIZED']);
  });

  it('is accepted by officePlugin, as flat OpenDocument is', async () => {
    const plugin = officePlugin();
    const file = (bytes: Buffer): ContainedFile => ({ sniffedType: 'application/xml', size: bytes.length, location: 'x', depth: 1, source: bufferSource(bytes) });
    expect([
      await plugin.accepts(file(flatWord('<w:p/>'))),
      await plugin.accepts(file(flatOdf('<office:body/>'))),
      await plugin.accepts(file(Buffer.from('<?xml version="1.0"?><svg/>'))),
    ]).to.deep.equal([true, true, false]);
  });
});
