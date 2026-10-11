import * as crypto from 'node:crypto';
import { expect } from 'chai';
import { inspectOffice } from '../src';
import { ODF_FORMATS, OOXML_FORMATS } from '../src/formats';
import { contentTypesXml, fixture, kinds, odfPackage, relsXml, T, textBody, wordPackage, wordXml, zipOf } from './helpers/util';

describe('format detection', () => {
  const exports: Array<[string, string, string]> = [
    ['clean.docx', 'docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    ['clean.docm', 'docm', 'application/vnd.ms-word.document.macroEnabled.12'],
    ['clean.dotx', 'dotx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'],
    ['clean.xlsx', 'xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    ['clean.xltx', 'xltx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'],
    ['clean.pptx', 'pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
    ['clean.potx', 'potx', 'application/vnd.openxmlformats-officedocument.presentationml.template'],
    ['clean.odt', 'odt', 'application/vnd.oasis.opendocument.text'],
    ['clean.ott', 'ott', 'application/vnd.oasis.opendocument.text-template'],
    ['clean.ods', 'ods', 'application/vnd.oasis.opendocument.spreadsheet'],
    ['clean.ots', 'ots', 'application/vnd.oasis.opendocument.spreadsheet-template'],
    ['clean.odp', 'odp', 'application/vnd.oasis.opendocument.presentation'],
    ['clean.otp', 'otp', 'application/vnd.oasis.opendocument.presentation-template'],
    ['clean.odg', 'odg', 'application/vnd.oasis.opendocument.graphics'],
    ['clean.otg', 'otg', 'application/vnd.oasis.opendocument.graphics-template'],
    ['clean.odf', 'odf', 'application/vnd.oasis.opendocument.formula'],
  ];
  for (const [name, format, mimeType] of exports) {
    it(`reads LibreOffice's ${name} as clean ${format}`, async () => {
      const r = await inspectOffice(fixture('lo', name));
      expect({ status: r.status, format: r.format, mimeType: r.mimeType, score: r.score, risk: r.risk }).to.deep.equal({ status: 'clean', format, mimeType, score: 0, risk: 'NONE' });
      expect(kinds(r).filter(k => k !== 'LINK/SAFE')).to.deep.equal([]);
    });
  }

  const built: Array<[string, string]> = [
    ['addin.xlam', 'xlam'],
    ['show.ppsm', 'ppsm'],
    ['visio.vsdx', 'vsdx'],
    ['base.odb', 'odb'],
    ['infozip.odt', 'odt'],
    ['utf16.docx', 'docx'],
  ];
  for (const [name, format] of built) {
    it(`reads ${name} as clean ${format}`, async () => {
      const r = await inspectOffice(fixture('docs', name));
      expect([r.status, r.format]).to.deep.equal(['clean', format]);
    });
  }

  it('reads a master document by its media type', async () => {
    const r = await inspectOffice(fixture('docs', 'master.odm'));
    expect([r.format, r.mimeType]).to.deep.equal(['odm', 'application/vnd.oasis.opendocument.text-master']);
  });

  it('takes every OOXML format from the main part content type, whatever the extension would say', async () => {
    for (const [type, f] of OOXML_FORMATS) {
      const r = await inspectOffice(await wordPackage('<w:p/>', { main: type }));
      expect([r.status, r.format, r.mimeType], type).to.deep.equal(['clean', f.format, f.mimeType]);
    }
  });

  it('takes the main part from the strict and the Visio relationship types too', async () => {
    for (const type of ['http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', 'http://schemas.microsoft.com/visio/2010/relationships/document']) {
      const zip = await zipOf([
        ['[Content_Types].xml', contentTypesXml({ 'main.xml': 'application/vnd.ms-visio.drawing.main+xml' })],
        ['_rels/.rels', relsXml(['rId1', type, 'main.xml'])],
        ['main.xml', '<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main"/>'],
      ]);
      expect((await inspectOffice(zip)).format, type).to.equal('vsdx');
    }
  });

  it('takes every OpenDocument format from the mimetype entry', async () => {
    for (const [type, f] of ODF_FORMATS) {
      const r = await inspectOffice(await odfPackage(textBody('<text:p/>'), { mimetype: type }));
      expect([r.status, r.format, r.mimeType], type).to.deep.equal(['clean', f.format, f.mimeType]);
    }
  });

  it('falls back on the manifest when there is no mimetype entry, and reports the layout', async () => {
    const r = await inspectOffice(await odfPackage(textBody(''), { mimetype: null, manifestType: 'application/vnd.oasis.opendocument.spreadsheet' }));
    expect(r.format).to.equal('ods');
    expect(kinds(r)).to.deep.equal(['PACKAGE/ODF_LAYOUT']);
    expect(r.findings[0].data).to.deep.equal({ reason: 'there is no mimetype entry' });
  });

  const refused: Array<[string, () => Promise<Buffer>, string]> = [
    ['OOXML and ODF markers together', async () => fixture('docs', 'ambiguous.docx'), 'PACKAGE/AMBIGUOUS_FORMAT'],
    ['bytes that are not a ZIP', async () => crypto.randomBytes(5000), 'CONTAINER/NOT_ZIP'],
    [
      'a main relationship to a missing part',
      () =>
        zipOf([
          ['[Content_Types].xml', contentTypesXml({})],
          ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'])],
        ]),
      'PACKAGE/UNRECOGNIZED',
    ],
    [
      'two main relationships',
      () =>
        zipOf([
          ['[Content_Types].xml', contentTypesXml({ 'word/document.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' })],
          ['_rels/.rels', relsXml(['rId1', `${T}officeDocument`, 'word/document.xml'], ['rId2', `${T}officeDocument`, 'word/document.xml'])],
          ['word/document.xml', wordXml('')],
        ]),
      'PACKAGE/AMBIGUOUS_FORMAT',
    ],
    ['an unknown ODF media type', () => odfPackage('', { mimetype: 'application/vnd.oasis.opendocument.chart' }), 'PACKAGE/UNRECOGNIZED'],
    [
      'a mimetype and a manifest that name different formats',
      () => odfPackage('', { mimetype: 'application/vnd.oasis.opendocument.text', manifestType: 'application/vnd.oasis.opendocument.spreadsheet' }),
      'PACKAGE/AMBIGUOUS_FORMAT',
    ],
  ];
  for (const [what, make, expected] of refused) {
    it(`rejects ${what} without a score`, async () => {
      const r = await inspectOffice(await make());
      expect({ status: r.status, score: r.score, risk: r.risk, last: kinds(r).slice(-1)[0] }).to.deep.equal({ status: 'rejected', score: null, risk: 'UNKNOWN', last: expected });
    });
  }
});
