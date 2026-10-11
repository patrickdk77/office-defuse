import { expect } from 'chai';
import { inspectOffice } from '../src';
import { fixture, kinds, livePeak, must, ODF_NS, OLE, odfPackage, of, rawZip, textBody, zipOf } from './helpers/util';

const noSafe = (k: string) => k !== 'LINK/SAFE';
const presentation = (inner: string) =>
  odfPackage(`<office:body><office:presentation><draw:page>${inner}</draw:page></office:presentation></office:body>`, { mimetype: 'application/vnd.oasis.opendocument.presentation' });
const MANIFEST_NS = 'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"';

describe('OpenDocument rules', () => {
  it('reports Basic, dialogs, scripts and menu settings by folder without parsing them, and script events', async () => {
    const r = await inspectOffice(fixture('docs', 'macros.odt'));
    expect(kinds(r).filter(noSafe)).to.deep.equal(['MACRO/EVENT_ON_OPEN', 'MACRO/EVENT', 'MACRO/ODF_BASIC', 'MACRO/ODF_DIALOGS', 'MACRO/ODF_SCRIPTS', 'UI/ODF_CONFIGURATION']);
    expect(of(r, 'MACRO/ODF_BASIC')[0]).to.include({ location: 'Basic/' });
    expect(of(r, 'MACRO/ODF_BASIC')[0].data).to.deep.equal({ entries: 3 });
    expect(of(r, 'MACRO/EVENT')[0].data?.target).to.match(/^vnd\.sun\.star\.script:LibreLogo/);
    expect([r.status, r.risk]).to.deep.equal(['strippable', 'CRITICAL']);
  });

  it('counts only files with content in the code and settings folders, so the empty accelerator file LibreOffice saves is not settings', async () => {
    const files = { 'Configurations2/accelerator/current.xml': '', 'Basic/Standard/Module1.xml': '', 'Object 1/Scripts/': '' };
    expect(kinds(await inspectOffice(await odfPackage(textBody(''), { files, unlisted: ['Object 1/Scripts/'] })))).to.deep.equal([]);
    const menu = await inspectOffice(await odfPackage(textBody(''), { files: { ...files, 'Configurations2/menubar/menubar.xml': '<menu/>' } }));
    expect([kinds(menu), menu.findings[0].location, menu.findings[0].data]).to.deep.equal([['UI/ODF_CONFIGURATION'], 'Configurations2/', { entries: 1 }]);
  });

  it('reports embedded objects, applets, plug-ins, frames and Impress actions', async () => {
    const r = await inspectOffice(fixture('docs', 'objects.odp'));
    expect(kinds(r).filter(noSafe)).to.deep.equal(['EMBEDDED/OLE_OBJECT', 'EMBEDDED/APPLET', 'EMBEDDED/PLUGIN', 'EXTERNAL/FRAME', 'ACTION/RUN_PROGRAM', 'ACTION/OLE_VERB', 'EMBEDDED/OLE_FILE']);
    expect(of(r, 'EMBEDDED/OLE_FILE')[0].location).to.equal('Pictures/stray.bin');
  });

  it('reports objects, applets and plug-ins without their optional attributes', async () => {
    const r = await inspectOffice(await presentation('<draw:frame><draw:applet/><draw:plugin/><draw:object-ole/><draw:object-ole xlink:href="./Object 9"/></draw:frame>'));
    expect(r.findings.map(f => [f.category, f.detail, f.data ?? {}])).to.deep.equal([
      ['EMBEDDED', 'APPLET', { code: '' }],
      ['EMBEDDED', 'PLUGIN', { mimeType: '' }],
      ['EMBEDDED', 'OLE_OBJECT', {}],
      ['EMBEDDED', 'OLE_OBJECT', { target: './Object 9' }],
    ]);
  });

  it('reports an OLE file as an object when an object element names it, and as a stray OLE file when none does', async () => {
    const named = await inspectOffice(await odfPackage(textBody('<draw:frame><draw:object-ole xlink:href="./Object%201"/></draw:frame>'), { files: { 'Object 1': OLE } }));
    expect(kinds(named)).to.deep.equal(['EMBEDDED/OLE_OBJECT']);
    const stray = await inspectOffice(await odfPackage(textBody('<draw:frame><draw:object xlink:href="./Object 1"/></draw:frame>'), { files: { 'Object 1': OLE } }));
    expect(kinds(stray)).to.deep.equal(['EMBEDDED/OLE_FILE']);
  });

  it('reports references that load on their own from outside, and a linked section as a local file', async () => {
    const r = await inspectOffice(fixture('docs', 'external.odt'));
    expect(kinds(r).filter(noSafe)).to.deep.equal(['EXTERNAL/IMAGE', 'EXTERNAL/LOCAL_FILE', 'EXTERNAL/TEMPLATE', 'EXTERNAL/RELOAD']);
    expect(of(r, 'EXTERNAL/LOCAL_FILE')[0].data).to.deep.equal({ target: '../shared/chapter.odt', element: 'text:section-source', kind: 'section-source' });
  });

  it('reads a percent-encoded path that climbs out of the package as outside it', async () => {
    const r = await inspectOffice(await odfPackage(textBody('<draw:frame><draw:image xlink:href="Pictures/%2e%2e/%2e%2e/x.png"/><draw:image xlink:href="%68ttp://example.com/x.png"/></draw:frame>')));
    expect(kinds(r)).to.deep.equal(['EXTERNAL/LOCAL_FILE', 'EXTERNAL/IMAGE']);
  });

  it('reports DDE links and DDE formulas, but not "DDE(" inside a string', async () => {
    expect(kinds(await inspectOffice(fixture('lo', 'dde.odt')))).to.deep.equal(['DDE/LINK']);
    const ods = await inspectOffice(fixture('lo', 'dde.ods'));
    expect(kinds(ods)).to.deep.equal(['DDE/FORMULA', 'DDE/LINK']);
    expect(of(ods, 'DDE/FORMULA')[0].data?.formula).to.equal('of:=DDE("cmd";"/c calc";"x")');
    const bare = await inspectOffice(await odfPackage(textBody('<text:dde-connection-decls><text:dde-connection-decl/></text:dde-connection-decls>')));
    expect(bare.findings.map(f => f.data)).to.deep.equal([{ application: '', topic: '' }]);
  });

  it('classifies links, including links that run a script', async () => {
    expect(kinds(await inspectOffice(fixture('lo', 'links.odt')))).to.deep.equal(['LINK/SAFE', 'MACRO/LINK', 'LINK/IP_HOST', 'LINK/CREDENTIALS']);
  });

  it('reads the document type OpenOffice.org 2.x wrote into the manifest as information, and entities in content as their values', async () => {
    const manifest = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">\n<manifest:manifest ${MANIFEST_NS}>\n <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>\n <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>\n</manifest:manifest>`;
    const content = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE office:document-content [<!ENTITY u "javascript:alert(1)">]><office:document-content ${ODF_NS} office:version="1.0">${textBody('<text:p><text:a xlink:href="&u;">x</text:a></text:p>')}</office:document-content>`;
    const odt = (m: string, c = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${ODF_NS} office:version="1.0">${textBody('')}</office:document-content>`) =>
      zipOf(
        [
          ['mimetype', 'application/vnd.oasis.opendocument.text'],
          ['content.xml', c],
          ['META-INF/manifest.xml', m],
        ],
        { stored: ['mimetype'] },
      );
    const plain = await inspectOffice(await odt(manifest));
    expect([plain.status, plain.risk, plain.format, kinds(plain), plain.findings[0].location]).to.deep.equal(['clean', 'NONE', 'odt', ['XML/DOCTYPE'], 'META-INF/manifest.xml']);
    expect(plain.findings[0].data).to.deep.equal({ part: 'META-INF/manifest.xml', name: 'manifest:manifest', publicId: '-//OpenOffice.org//DTD Manifest 1.0//EN', systemId: 'Manifest.dtd' });
    const linked = await inspectOffice(await odt(manifest, content));
    expect([linked.status, kinds(linked), of(linked, 'LINK/SCRIPT_URL')[0].data?.target]).to.deep.equal([
      'strippable',
      ['XML/DOCTYPE', 'XML/INTERNAL_SUBSET', 'LINK/SCRIPT_URL'],
      'javascript:alert(1)',
    ]);
    // A manifest whose prefix only the external DTD declared, as in an OpenOffice.org database file, stays refused.
    const undeclared = await inspectOffice(await odt(manifest.replace(` ${MANIFEST_NS}`, '')));
    expect([undeclared.status, undeclared.format, kinds(undeclared)]).to.deep.equal(['rejected', '', ['XML/DOCTYPE', 'XML/MALFORMED']]);
  });

  it('rejects an encrypted package without a score', async () => {
    const r = await inspectOffice(fixture('docs', 'encrypted.odt'));
    expect([r.status, r.score, r.format, kinds(r)]).to.deep.equal(['rejected', null, 'odt', ['ENCRYPTED/PASSWORD_REQUIRED']]);
  });

  it('reports a mimetype entry that is not first, compressed, padded or carrying an extra field, and a missing manifest', async () => {
    const r = await inspectOffice(fixture('docs', 'layout.odt'));
    expect(r.findings.filter(f => noSafe(`${f.category}/${f.detail}`)).map(f => f.data?.reason)).to.deep.equal(['mimetype is not the first entry', 'mimetype is compressed']);
    const padded = await inspectOffice(
      await odfPackage(textBody(''), { mimetype: ' application/vnd.oasis.opendocument.text\n', manifestType: 'application/vnd.oasis.opendocument.text', deflateMimetype: true }),
    );
    expect([padded.format, padded.findings.map(f => f.data?.reason)]).to.deep.equal(['odt', ['mimetype is compressed', 'mimetype has white space around the media type']]);
    const extra = await inspectOffice(
      rawZip([
        { name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', method: 0, localExtra: Buffer.from([0xfe, 0xca, 0, 0]) },
        { name: 'content.xml', data: `<office:document-content ${ODF_NS}/>` },
        {
          name: 'META-INF/manifest.xml',
          data: `<manifest:manifest ${MANIFEST_NS}><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml"/></manifest:manifest>`,
        },
      ]),
    );
    expect(extra.findings.map(f => f.data?.reason)).to.deep.equal(['mimetype has an extra field']);
    // First in the file is what counts, wherever the central directory lists it.
    const listedLate = await inspectOffice(
      rawZip(
        [
          { name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', method: 0 },
          { name: 'content.xml', data: `<office:document-content ${ODF_NS}/>` },
          {
            name: 'META-INF/manifest.xml',
            data: `<manifest:manifest ${MANIFEST_NS}><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml"/></manifest:manifest>`,
          },
        ],
        { centralOrder: [1, 0, 2] },
      ),
    );
    expect([listedLate.status, kinds(listedLate)]).to.deep.equal(['clean', []]);
    const bare = await inspectOffice(
      await zipOf(
        [
          ['mimetype', 'application/vnd.oasis.opendocument.text'],
          ['content.xml', `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>`],
        ],
        { stored: ['mimetype'] },
      ),
    );
    expect([bare.format, bare.findings.map(f => f.data?.reason)]).to.deep.equal(['odt', ['there is no manifest']]);
  });

  it('reads a long mimetype entry in bounded memory, and reads it as no media type', async () => {
    const zip = await odfPackage(textBody(''), {
      mimetype: `application/vnd.oasis.opendocument.text${' '.repeat(128 * 1024 * 1024)}`,
      manifestType: 'application/vnd.oasis.opendocument.text',
      deflateMimetype: true,
    });
    let r: Awaited<ReturnType<typeof inspectOffice>> | undefined;
    const peak = await livePeak(async () => {
      r = await inspectOffice(zip);
    });
    const result = must(r, 'result');
    expect(zip.length).to.be.lessThan(1024 * 1024);
    expect([result.status, kinds(result).slice(-1)[0]]).to.deep.equal(['rejected', 'PACKAGE/UNRECOGNIZED']);
    expect(String(result.findings.slice(-1)[0].data?.mediaType)).to.match(/^application\/vnd\.oasis\.opendocument\.text +\.\.\.$/);
    // Holding the entry would take 128 MB.
    expect(peak).to.be.lessThan(32 * 1024 * 1024);
  });

  it('reads OpenDocument namespaces with another version or committee name, as LibreOffice does', async () => {
    const ns = 'xmlns:s="urn:oasis:names:tc:other:xmlns:script:1.2"';
    const body = textBody(`<office:event-listeners ${ns}><s:event-listener s:event-name="dom:load" xmlns:l="http://www.w3.org/1999/xlink" l:href="macro:///x"/></office:event-listeners>`);
    expect(kinds(await inspectOffice(await odfPackage(body)))).to.deep.equal(['MACRO/EVENT_ON_OPEN']);
    const v2 = textBody('<office:event-listeners xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:2.0"><s:event-listener s:event-name="dom:load"/></office:event-listeners>');
    expect(kinds(await inspectOffice(await odfPackage(v2)))).to.deep.equal([]);
  });

  it('counts focus as an opening event only for the document itself, and skips a listener that names no script', async () => {
    const doc = `<office:scripts><office:event-listeners><script:event-listener script:event-name="office:focus" xlink:href="macro:///a"/><script:event-listener/><script:event-listener script:event-name="dom:load" xlink:href=""/><script:event-listener script:event-name="office:print" xlink:href="" script:macro-name="c"/><script:event-listener script:event-name="office:save" xlink:href=" "/></office:event-listeners></office:scripts>${textBody(
      '<text:p><text:a xlink:href="https://example.com/"><office:event-listeners><script:event-listener script:event-name="dom:focus" script:macro-name="b"/></office:event-listeners>x</text:a></text:p>',
    )}`;
    const r = await inspectOffice(await odfPackage(doc));
    // LibreOffice checks for an empty value, not a blank one, so a space still binds the event.
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.data?.target ?? ''}`)).to.deep.equal([
      'MACRO/EVENT_ON_OPEN macro:///a',
      'MACRO/EVENT c',
      'MACRO/EVENT  ',
      'LINK/SAFE https://example.com/',
      'MACRO/EVENT b',
    ]);
  });

  it('reports inline scripts by language, and empty script elements not at all', async () => {
    const body = `<office:scripts><office:script script:language="ooo:Basic"><x>Sub Main</x></office:script><office:script script:language="ooo:Python">print(1)</office:script><office:script script:language="ooo:Basic"/><office:script> </office:script><office:script>x</office:script></office:scripts>`;
    expect(kinds(await inspectOffice(await odfPackage(body)))).to.deep.equal(['MACRO/ODF_BASIC', 'MACRO/ODF_SCRIPTS', 'MACRO/ODF_SCRIPTS']);
  });

  it('reports files the manifest leaves out, paths it lists twice, and OLE files no object uses', async () => {
    const r = await inspectOffice(
      await odfPackage(textBody(''), {
        files: { 'extra.xml': '<a/>', 'Object 2': OLE },
        unlisted: ['extra.xml'],
        manifestExtra:
          '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="gone.xml"/><manifest:file-entry manifest:full-path="gone.xml"/><manifest:file-entry/><manifest:other/>',
      }),
    );
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.location}`)).to.deep.equal([
      'PACKAGE/DUPLICATE_ENTRY META-INF/manifest.xml',
      'EMBEDDED/OLE_FILE Object 2',
      'PACKAGE/UNREFERENCED_PART extra.xml',
    ]);
  });

  it('rejects a manifest that names the package root twice, and reads a root without a media type as unrecognized', async () => {
    const twice = await inspectOffice(await odfPackage(textBody(''), { manifestExtra: '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>' }));
    expect([twice.status, kinds(twice)]).to.deep.equal(['rejected', ['PACKAGE/DUPLICATE_ENTRY']]);
    const manifest = `<manifest:manifest ${MANIFEST_NS}><manifest:file-entry manifest:full-path="/"/><x:other xmlns:x="urn:x"/></manifest:manifest>`;
    const r = await inspectOffice(await zipOf([['META-INF/manifest.xml', manifest]]));
    expect([kinds(r), r.findings[1].data?.mediaType]).to.deep.equal([['PACKAGE/ODF_LAYOUT', 'PACKAGE/UNRECOGNIZED'], '']);
    const rootless = await inspectOffice(await zipOf([['META-INF/manifest.xml', `<manifest:manifest ${MANIFEST_NS}/>`]]));
    expect([kinds(rootless), rootless.findings[1].data?.mediaType]).to.deep.equal([['PACKAGE/ODF_LAYOUT', 'PACKAGE/UNRECOGNIZED'], '']);
  });

  it('reports links in form controls and Impress events that open documents or run programs', async () => {
    const r = await inspectOffice(
      await presentation(
        '<draw:rect><office:event-listeners><presentation:event-listener presentation:action="show" xlink:href="https://example.com/next"/><presentation:event-listener presentation:action="execute" xlink:href="calc.exe"/><presentation:event-listener presentation:action="verb"/><presentation:event-listener xlink:href="#slide2"/><presentation:event-listener presentation:action="next-page"/></office:event-listeners></draw:rect><form:button xlink:href="file:///etc/passwd"/><form:form xlink:href="https://example.com/submit"/>',
      ),
    );
    expect(r.findings.map(f => `${f.category}/${f.detail} ${f.data?.action ?? ''}`)).to.deep.equal(['LINK/SAFE show', 'ACTION/RUN_PROGRAM execute', 'ACTION/OLE_VERB verb', 'LINK/FILE_URL ']);
  });

  it('leaves an embedded database connection alone and reports one to a server', async () => {
    expect(kinds(await inspectOffice(fixture('docs', 'base.odb')))).to.deep.equal([]);
    const body =
      '<office:body><office:database xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><db:data-source><db:connection-resource xlink:href="sdbc:mysql:jdbc:db.example.net:3306/sales"/></db:data-source></office:database></office:body>';
    expect(kinds(await inspectOffice(await odfPackage(body, { mimetype: 'application/vnd.oasis.opendocument.base' })))).to.deep.equal(['EXTERNAL/OTHER']);
  });
});
