import * as zlib from 'node:zlib';
import { expect } from 'chai';
import { type ContainedFile, type ContainedFilePlugin, disarmOffice, inspectOffice } from '../src';
import { crc32 } from '../src/crc32';
import { entryNames, entryText, structureProblems } from './helpers/structure';
import { kinds, must, of, zipOf } from './helpers/util';

const XPS = 'http://schemas.microsoft.com/xps/2005/06';
const OXPS = 'http://schemas.openxps.org/oxps/v1.0';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

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

interface XpsOptions {
  /** OpenXPS, ECMA-388, instead of Microsoft XPS. */
  open?: boolean;
  /** Markup inside the FixedPage element. */
  page?: string;
  /** Attributes of the DocumentReference, which hold its Source by default. */
  documentSource?: string;
  /** Further parts, by name. */
  parts?: Record<string, string | Uint8Array>;
}

/** An XPS package of one document with one page, and an image at /Resources/image.png. */
function xpsPackage(o: XpsOptions = {}): Promise<Buffer> {
  const ns = o.open ? OXPS : XPS;
  const main = o.open ? 'http://schemas.openxps.org/oxps/v1.0/fixedrepresentation' : 'http://schemas.microsoft.com/xps/2005/06/fixedrepresentation';
  const types =
    '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="fdseq" ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>' +
    '<Default Extension="fdoc" ContentType="application/vnd.ms-package.xps-fixeddocument+xml"/>' +
    '<Default Extension="fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Default Extension="odttf" ContentType="application/vnd.ms-package.obfuscated-opentype"/>' +
    '<Default Extension="struct" ContentType="application/vnd.ms-package.xps-documentstructure+xml"/>' +
    '</Types>';
  return zipOf([
    ['[Content_Types].xml', types],
    ['_rels/.rels', `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${RELS}"><Relationship Id="R0" Type="${main}" Target="/FixedDocSeq.fdseq"/></Relationships>`],
    ['FixedDocSeq.fdseq', `<FixedDocumentSequence xmlns="${ns}"><DocumentReference ${o.documentSource ?? 'Source="/Documents/1/FixedDoc.fdoc"'}/></FixedDocumentSequence>`],
    ['Documents/1/FixedDoc.fdoc', `<FixedDocument xmlns="${ns}"><PageContent Source="Pages/1.fpage"/></FixedDocument>`],
    ['Documents/1/Pages/1.fpage', `<FixedPage xmlns="${ns}" Width="816" Height="1056" xml:lang="en-US">${o.page ?? ''}</FixedPage>`],
    ['Resources/image.png', png()],
    ...Object.entries(o.parts ?? {}),
  ]);
}

/** A path filled with the image at `source`. */
const imagePath = (source: string) =>
  `<Path Data="M 0,0 L 10,0 10,10 0,10 z"><Path.Fill><ImageBrush ImageSource="${source}" Viewbox="0,0,1,1" Viewport="0,0,10,10" ViewboxUnits="Absolute" ViewportUnits="Absolute"/></Path.Fill></Path>`;

/** A plugin that takes every file and answers `answer`. It records the names it saw. */
function recorder(answer: 'passed' | 'removed') {
  const seen: ContainedFile[] = [];
  const plugin: ContainedFilePlugin = {
    kind: 'file',
    name: 'test',
    accepts: () => true,
    async process(file) {
      seen.push(file);
      return answer;
    },
  };
  return { plugin, seen };
}

describe('XPS', () => {
  it('reads XPS and OpenXPS packages, whose pages and resources the markup reaches', async () => {
    for (const open of [false, true]) {
      const r = await inspectOffice(await xpsPackage({ open, page: imagePath('/Resources/image.png') }));
      expect([r.status, r.format, r.mimeType, kinds(r)]).to.deep.equal(['clean', open ? 'oxps' : 'xps', open ? 'application/oxps' : 'application/vnd.ms-xpsdocument', []]);
    }
    // A relative reference resolves against the folder of the part that makes it.
    expect(kinds(await inspectOffice(await xpsPackage({ page: imagePath('../../../Resources/image.png') })))).to.deep.equal([]);
    // A part no relationship and no markup reaches is still reported.
    expect(kinds(await inspectOffice(await xpsPackage({ page: imagePath('/Resources/image.png'), parts: { 'Resources/extra.png': png() } })))).to.deep.equal(['PACKAGE/UNREFERENCED_PART']);
  });

  it('checks links with a scheme or a network path, and leaves links inside the document alone', async () => {
    const page = (uri: string) => `<Canvas FixedPage.NavigateUri="${uri}"><Path Data="M 0,0 L 1,1"/></Canvas>${imagePath('/Resources/image.png')}`;
    const detail = async (uri: string) => kinds(await inspectOffice(await xpsPackage({ page: page(uri) })));
    expect(await detail('https://example.com/')).to.deep.equal(['LINK/SAFE']);
    expect(await detail('javascript:alert(1)')).to.deep.equal(['LINK/SCRIPT_URL']);
    expect(await detail('file:///C:/Windows/notepad.exe')).to.deep.equal(['LINK/FILE_URL']);
    expect(await detail('\\\\host\\share\\x')).to.deep.equal(['LINK/NETWORK_PATH']);
    expect(await detail('../FixedDoc.fdoc#Page2')).to.deep.equal([]);
    expect(await detail('#Page2')).to.deep.equal([]);
    const outline = await inspectOffice(
      await xpsPackage({
        page: imagePath('/Resources/image.png'),
        parts: {
          'Documents/1/Structure/DocStructure.struct': `<DocumentStructure xmlns="${XPS}"><DocumentStructure.Outline><DocumentOutline><OutlineEntry OutlineTarget="http://203.0.113.9/x" Description="x"/></DocumentOutline></DocumentStructure.Outline></DocumentStructure>`,
        },
      }),
    );
    // No relationship reaches the structure part here, so it is reported for that too.
    expect(kinds(outline)).to.deep.equal(['LINK/IP_HOST', 'PACKAGE/UNREFERENCED_PART']);
  });

  it('reports resources that load from outside the package', async () => {
    const external = await inspectOffice(
      await xpsPackage({
        page: `${imagePath('/Resources/image.png')}${imagePath('http://example.com/a.png')}<Glyphs FontUri="\\\\host\\fonts\\f.odttf" OriginX="0" OriginY="10" FontRenderingEmSize="10" UnicodeString="x"/><Canvas><Canvas.Resources><ResourceDictionary Source="file:///C:/dict.xaml"/></Canvas.Resources></Canvas>`,
      }),
    );
    expect(external.findings.map(f => `${f.category}/${f.detail}/${f.data?.attribute}`)).to.deep.equal(['EXTERNAL/IMAGE/ImageSource', 'EXTERNAL/OTHER/FontUri', 'EXTERNAL/LOCAL_FILE/Source']);
    const document = await inspectOffice(await xpsPackage({ documentSource: 'Source="http://example.com/FixedDoc.fdoc"' }));
    expect(kinds(document)).to.include('EXTERNAL/INCLUDE');
  });

  it('defuses links and external resources, and leaves no empty property element', async () => {
    const input = await xpsPackage({
      page: `<Canvas FixedPage.NavigateUri="javascript:alert(1)"><Path Data="M 0,0 L 1,1"/></Canvas>${imagePath('http://example.com/a.png')}${imagePath('/Resources/image.png')}`,
    });
    const r = await disarmOffice(input);
    expect(r.status).to.equal('defused');
    const out = must(r.bytes, 'output');
    expect([kinds(must(r.after, 'after')), await structureProblems(out)]).to.deep.equal([[], []]);
    const page = must(await entryText(out, 'Documents/1/Pages/1.fpage'), 'page');
    expect([page.includes('NavigateUri'), page.includes('example.com'), (page.match(/<Path.Fill>/g) ?? []).length, page.includes('/Resources/image.png')]).to.deep.equal([false, false, 1, true]);
  });

  it('shows images to the plugins, and takes away the markup that names one a plugin removed', async () => {
    const font = Buffer.alloc(64, 1);
    const input = await xpsPackage({ page: imagePath('/Resources/image.png'), parts: { 'Resources/font.odttf': font } });
    const passed = recorder('passed');
    await inspectOffice(input, { filePlugins: [passed.plugin] });
    // The obfuscated font is part of the format, not a file of its own. It is unreferenced here, as no markup names it.
    expect(passed.seen.map(f => [f.name, f.sniffedType])).to.deep.equal([['Resources/image.png', 'image/png']]);
    const removed = recorder('removed');
    const r = await disarmOffice(await xpsPackage({ page: imagePath('/Resources/image.png') }), { filePlugins: [removed.plugin] });
    expect([r.status, of(r.before, 'EMBEDDED/PLUGIN_REMOVED').length]).to.deep.equal(['defused', 1]);
    const out = must(r.bytes, 'output');
    expect(await entryNames(out)).to.not.include('Resources/image.png');
    const page = must(await entryText(out, 'Documents/1/Pages/1.fpage'), 'page');
    expect([page.includes('ImageBrush'), page.includes('Path.Fill'), page.includes('<Path ')]).to.deep.equal([false, false, true]);
    expect(await structureProblems(out)).to.deep.equal([]);
  });
});
