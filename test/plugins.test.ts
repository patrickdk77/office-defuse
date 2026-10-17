import { expect } from 'chai';
import { type ContainedFile, type ContainedFilePlugin, disarmOffice, inspectOffice, officePlugin, type PluginContext } from '../src';
import { entryNames, entryText, structureProblems } from './helpers/structure';
import { fixture, kinds, must, OLE, odfPackage, of, T, textBody, wordPackage } from './helpers/util';

const SHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** A Word package holding `inner` as an embedded package, the way Word embeds a workbook. */
const withPackage = (inner: Uint8Array) =>
  wordPackage('<w:p/>', {
    parts: { 'word/embeddings/Book1.xlsx': inner },
    overrides: { 'word/embeddings/Book1.xlsx': SHEET },
    rels: [['rIdPkg', `${T}package`, 'embeddings/Book1.xlsx']],
  });

/** A plugin that takes every file and answers with `answer`, writing `bytes` when given. It records what it saw. */
function plugin(answer: (file: ContainedFile, context: PluginContext) => Promise<'passed' | 'scrubbed' | 'removed'> | 'passed' | 'scrubbed' | 'removed', bytes?: Uint8Array) {
  const seen: Array<{ file: ContainedFile; context: PluginContext }> = [];
  const p: ContainedFilePlugin = {
    kind: 'file',
    name: 'test',
    accepts: () => true,
    async process(file, sink, context) {
      seen.push({ file, context });
      if (bytes) await sink.write(bytes);
      return answer(file, context);
    },
  };
  return { plugin: p, seen };
}

describe('file plugins', () => {
  it('shows an embedded package to the plugins with its name, type, sniffed type, size, depth and the deadline', async () => {
    const inner = fixture('lo', 'clean.xlsx');
    const p = plugin(() => 'passed');
    const r = await inspectOffice(await withPackage(inner), { filePlugins: [p.plugin], limits: { timeMs: 60_000 } });
    expect(p.seen).to.have.length(1);
    const { file, context } = p.seen[0];
    expect([file.name, file.declaredType, file.sniffedType, file.size, file.location, file.depth, context.depth]).to.deep.equal([
      'word/embeddings/Book1.xlsx',
      SHEET.toLowerCase(),
      'application/zip',
      inner.length,
      'word/embeddings/Book1.xlsx',
      1,
      1,
    ]);
    expect(context.deadline).to.be.a('number');
    // The plugin kept the file, so the finding for its being there only reports.
    expect(r.findings.map(f => `${f.category}/${f.detail}/${f.action}`)).to.deep.equal(['EMBEDDED/PLUGIN_PASSED/info']);
    expect([r.status, r.risk]).to.deep.equal(['clean', 'NONE']);
  });

  it('keeps a file a plugin passed, and checks the output against what the plugin returned', async () => {
    const input = await withPackage(fixture('lo', 'clean.xlsx'));
    const r = await disarmOffice(input, { filePlugins: [plugin(() => 'passed').plugin] });
    expect(r.status).to.equal('clean');
    const out = must(r.bytes, 'output');
    expect(await entryNames(out)).to.include('word/embeddings/Book1.xlsx');
    expect(kinds(must(r.after, 'after'))).to.deep.equal(['EMBEDDED/PLUGIN_PASSED']);
    expect(await structureProblems(out)).to.deep.equal([]);
  });

  it('puts the bytes a plugin wrote in place of the file it scrubbed', async () => {
    const cleaned = fixture('lo', 'clean.xlsx');
    const r = await disarmOffice(await withPackage(fixture('docs', 'excel.xlsm')), { filePlugins: [plugin(() => 'scrubbed', cleaned).plugin] });
    expect([r.status, r.before.status]).to.deep.equal(['defused', 'strippable']);
    const out = must(r.bytes, 'output');
    const { ZipArchive } = await import('../src/zip');
    const { bufferSource } = await import('../src/io');
    const zip = await ZipArchive.open(bufferSource(out), { entries: 100, decompressedBytes: 1 << 30, totalDecompressedBytes: 1 << 30 });
    const chunks: Buffer[] = [];
    for await (const c of zip.read(must(zip.get('word/embeddings/Book1.xlsx'), 'part'))) chunks.push(Buffer.from(c));
    expect(Buffer.concat(chunks).equals(cleaned)).to.equal(true);
    expect(kinds(must(r.after, 'after'))).to.deep.equal(['EMBEDDED/PLUGIN_SCRUBBED']);
  });

  it('drops a file a plugin removed, failed on, or scrubbed without writing anything, whatever the overrides say', async () => {
    const input = await withPackage(fixture('lo', 'clean.xlsx'));
    const info = { actionOverrides: [{ category: 'EMBEDDED' as never, action: 'info' as const }] };
    const failing: ContainedFilePlugin = { kind: 'file', name: 'broken', accepts: () => true, process: async () => Promise.reject(new Error('plugin bug')) };
    for (const [p, detail] of [
      [plugin(() => 'removed').plugin, 'PLUGIN_REMOVED'],
      [failing, 'PLUGIN_FAILED'],
      [plugin(() => 'scrubbed').plugin, 'PLUGIN_FAILED'],
    ] as const) {
      const r = await disarmOffice(input, { ...info, filePlugins: [p] });
      expect([r.status, of(r.before, `EMBEDDED/${detail}`)[0]?.action], detail).to.deep.equal(['defused', 'strip']);
      expect(await entryNames(must(r.bytes, 'output'))).to.not.include('word/embeddings/Book1.xlsx');
      expect(await entryText(must(r.bytes, 'output'), 'word/_rels/document.xml.rels')).to.not.include('Book1');
    }
  });

  it('labels the findings a plugin reports from inside a file, and turns a reject inside it into strip', async () => {
    const p: ContainedFilePlugin = {
      kind: 'file',
      name: 'reporter',
      accepts: () => true,
      process: async () => ({ result: 'removed', findings: [{ category: 'X', detail: 'Y', description: 'inside', action: 'reject', location: 'page 1', weight: 45 }] }),
    };
    const r = await inspectOffice(await withPackage(fixture('lo', 'clean.xlsx')), { filePlugins: [p] });
    const inside = must(of(r, 'X/Y')[0], 'inside finding');
    expect([inside.action, inside.location, inside.attachment]).to.deep.equal(['strip', 'word/embeddings/Book1.xlsx > page 1', 'word/embeddings/Book1.xlsx']);
    // A kind office-defuse does not make scores at the weight it carries.
    expect(r.score).to.be.at.least(45);
  });

  it('passes a time limit or an I/O error from a plugin up instead of counting it as a failed plugin', async () => {
    const input = await withPackage(fixture('lo', 'clean.xlsx'));
    const limit: ContainedFilePlugin = {
      kind: 'file',
      name: 'slow',
      accepts: () => true,
      process: async () => Promise.reject(Object.assign(new Error('late'), { code: 'DEFUSE_LIMIT', limit: 'time' })),
    };
    const r = await disarmOffice(input, { filePlugins: [limit] });
    expect([r.status, r.before.score, kinds(r.before).at(-1)]).to.deep.equal(['rejected', null, 'LIMIT/TIME']);
    const cause = Object.assign(new Error('disk gone'), { code: 'EIO' });
    const io: ContainedFilePlugin = { kind: 'file', name: 'io', accepts: () => true, process: async () => Promise.reject(Object.assign(new Error('nested'), { code: 'DEFUSE_IO', cause })) };
    let thrown: unknown;
    try {
      await disarmOffice(input, { filePlugins: [io] });
    } catch (e) {
      thrown = e;
    }
    expect(thrown).to.equal(cause);
    const accepts: ContainedFilePlugin = {
      kind: 'file',
      name: 'picky',
      accepts: () => Promise.reject(Object.assign(new Error('late'), { code: 'DEFUSE_LIMIT', limit: 'nesting' })),
      process: async () => 'passed',
    };
    expect(kinds((await disarmOffice(input, { filePlugins: [accepts] })).before).at(-1)).to.equal('LIMIT/NESTING_DEPTH');
  });

  it('gives a file no plugin accepts the rule it gets without plugins', async () => {
    const none: ContainedFilePlugin = { kind: 'file', name: 'none', accepts: () => false, process: async () => 'passed' };
    const r = await disarmOffice(await withPackage(fixture('lo', 'clean.xlsx')), { filePlugins: [none] });
    expect([r.status, r.before.findings.map(f => `${f.category}/${f.detail}/${f.action}`)]).to.deep.equal(['defused', ['EMBEDDED/PACKAGE/strip']]);
  });

  it('shows the OLE parts of a package to the plugins: VBA projects, objects and OLE files', async () => {
    const p = plugin(() => 'passed');
    await inspectOffice(fixture('docs', 'objects.docx'), { filePlugins: [p.plugin] });
    await inspectOffice(fixture('docs', 'vba.docm'), { filePlugins: [p.plugin] });
    expect(p.seen.map(s => [s.file.name, s.file.sniffedType])).to.deep.equal([
      ['word/embeddings/oleObject1.bin', 'application/x-ole-storage'],
      ['word/embeddings/Book1.xlsx', 'application/zip'],
      ['word/media/image1.png', 'application/x-ole-storage'],
      ['word/activeX/activeX1.bin', 'application/x-ole-storage'],
      ['word/vbaProject.bin', 'application/x-ole-storage'],
    ]);
  });

  it('keeps an OpenDocument object and its entry when a plugin keeps the entry, and drops both when it removes it', async () => {
    const body = textBody('<text:p><draw:frame><draw:object-ole xlink:href="./Object 1"/><draw:image xlink:href="./ObjectReplacements/Object 1"/></draw:frame></text:p>');
    const input = await odfPackage(body, { files: { 'Object 1': OLE, 'ObjectReplacements/Object 1': Buffer.from('preview') } });
    const kept = await disarmOffice(input, { filePlugins: [plugin(() => 'passed').plugin] });
    // The preview image is a whole file too, and the plugin passes it as well.
    expect([kept.status, kept.before.findings.map(f => `${f.category}/${f.detail}/${f.action} ${f.location}`)]).to.deep.equal([
      'clean',
      ['EMBEDDED/PLUGIN_PASSED/info Object 1', 'EMBEDDED/PLUGIN_PASSED/info ObjectReplacements/Object 1'],
    ]);
    expect(await entryText(must(kept.bytes, 'output'), 'content.xml')).to.include('draw:object-ole');
    // A plugin that removes the object and passes its preview.
    const removed = await disarmOffice(input, { filePlugins: [plugin(f => (f.sniffedType === 'application/x-ole-storage' ? 'removed' : 'passed')).plugin] });
    const out = must(removed.bytes, 'output');
    expect(await entryNames(out)).to.not.include('Object 1');
    const content = must(await entryText(out, 'content.xml'), 'content');
    expect(content).to.not.include('draw:object-ole');
    expect(content).to.include('ObjectReplacements/Object 1');
  });
});

describe('officePlugin', () => {
  it('accepts what detectOffice names, by content alone', async () => {
    const p = officePlugin();
    const { bufferSource } = await import('../src/io');
    const { sniffType } = await import('../src/sniff');
    const file = (bytes: Uint8Array, name: string): ContainedFile => ({
      name,
      size: bytes.length,
      location: 'x',
      depth: 1,
      source: bufferSource(bytes),
      sniffedType: sniffType(bytes.subarray(0, 1024)),
    });
    expect(await p.accepts(file(fixture('lo', 'clean.docx'), 'picture.png'))).to.equal(true);
    expect(await p.accepts(file(fixture('lo', 'clean.odt'), 'x'))).to.equal(true);
    expect(await p.accepts(file(fixture('ole', 'word.doc'), 'notes.txt'))).to.equal(true);
    expect(await p.accepts(file(fixture('rtf', 'plain.rtf'), 'a.doc'))).to.equal(true);
    expect(await p.accepts(file(fixture('cases', 'plain-archive.zip'), 'report.docx'))).to.equal(false);
    expect(await p.accepts(file(Buffer.from('%PDF-1.7\n'), 'x.docx'))).to.equal(false);
    expect(await p.accepts(file(Buffer.from('PK\x03\x04 cut short'), 'x.docx'))).to.equal(false);
  });

  it('defuses a package inside a package, and reports what it removed there in the outer report', async () => {
    const input = await withPackage(fixture('docs', 'excel.xlsm'));
    const r = await disarmOffice(input, { filePlugins: [officePlugin()] });
    expect([r.status, r.before.status]).to.deep.equal(['defused', 'strippable']);
    const inner = r.before.findings.filter(f => f.attachment === 'word/embeddings/Book1.xlsx');
    expect(inner.map(f => `${f.category}/${f.detail}`)).to.include.members(['MACRO/XLM_MACROSHEET', 'MACRO/EVENT_ON_OPEN', 'DDE/LINK']);
    expect(inner.every(f => typeof f.weight === 'number')).to.equal(true);
    expect(of(r.before, 'EMBEDDED/PLUGIN_SCRUBBED')[0]?.data?.plugin).to.equal('office');
    const out = must(r.bytes, 'output');
    expect(await structureProblems(out)).to.deep.equal([]);
    // The package inside the output is the defused one.
    const nested = await inspectOffice(await entryBytes(out, 'word/embeddings/Book1.xlsx'));
    expect([nested.status, nested.format]).to.deep.equal(['clean', 'xlsm']);
  });

  it('runs with the options of the run that found the file, or its own', async () => {
    // An encrypted OOXML package, which detectOffice claims as a compound file until a password opens it, and a
    // workbook whose records name it before its password opens it.
    for (const encrypted of [fixture('ole', 'password.docx'), fixture('ole', 'rc4-password.xls')]) {
      const input = await withPackage(encrypted);
      const inherited = await inspectOffice(input, { password: ['nope', 'inert-password'], filePlugins: [officePlugin()] });
      expect(kinds(inherited)).to.include('ENCRYPTED/USER_PASSWORD');
      const own = await inspectOffice(input, { password: 'inert-password', filePlugins: [officePlugin({})] });
      expect(kinds(own)).to.include('ENCRYPTED/PASSWORD_REQUIRED');
    }
  });

  it('stops the whole upload at the nesting limit, counted from the upload', async () => {
    const twice = await withPackage(await withPackage(fixture('lo', 'clean.xlsx')));
    const r = await disarmOffice(twice, { filePlugins: [officePlugin()], limits: { nestingDepth: 1 } });
    expect([r.status, r.before.score, kinds(r.before).at(-1)]).to.deep.equal(['rejected', null, 'LIMIT/NESTING_DEPTH']);
    const ok = await disarmOffice(twice, { filePlugins: [officePlugin()], limits: { nestingDepth: 2 } });
    expect(ok.status).to.equal('clean');
    expect(kinds(await inspectOffice(fixture('lo', 'clean.docx'), { limits: { nestingDepth: -1 } }))).to.deep.equal(['LIMIT/NESTING_DEPTH']);
  });
});

async function entryBytes(bytes: Uint8Array, name: string): Promise<Buffer> {
  const { ZipArchive } = await import('../src/zip');
  const { bufferSource } = await import('../src/io');
  const zip = await ZipArchive.open(bufferSource(bytes), { entries: 1000, decompressedBytes: 1 << 30, totalDecompressedBytes: 1 << 30 });
  const chunks: Buffer[] = [];
  for await (const c of zip.read(must(zip.get(name), name))) chunks.push(Buffer.from(c));
  return Buffer.concat(chunks);
}

describe('whole files inside a file', () => {
  const names = (p: { seen: Array<{ file: ContainedFile }> }) => p.seen.map(s => s.file.location);

  it('shows pictures and other whole-file parts to the plugins, and not the binary parts of the format itself', async () => {
    const p = plugin(() => 'passed');
    const docx = await inspectOffice(fixture('media', 'pictures.docx'), { filePlugins: [p.plugin] });
    expect(names(p)).to.deep.equal(['word/media/image1.png', 'word/media/image2.jpeg', 'docProps/thumbnail.jpeg']);
    expect(p.seen.map(s => [s.file.declaredType, s.file.sniffedType])).to.deep.equal([
      ['image/png', 'image/png'],
      ['image/jpeg', 'image/jpeg'],
      ['image/jpeg', 'image/jpeg'],
    ]);
    // No rule names a picture, so a plugin's verdict on one is all the report says of it.
    expect([docx.status, docx.risk, kinds(docx)]).to.deep.equal(['clean', 'NONE', ['EMBEDDED/PLUGIN_PASSED', 'EMBEDDED/PLUGIN_PASSED', 'EMBEDDED/PLUGIN_PASSED']]);
    const odt = plugin(() => 'passed');
    await inspectOffice(fixture('media', 'pictures.odt'), { filePlugins: [odt.plugin] });
    expect(names(odt)).to.deep.equal(['Pictures/1.png', 'Pictures/2.jpeg', 'Thumbnails/thumbnail.png']);
    const rtf = plugin(() => 'passed');
    await inspectOffice(fixture('media', 'pictures.rtf'), { filePlugins: [rtf.plugin] });
    // The PNG and JPEG pictures, in hex or binary. The metafile is no file of its own.
    expect(rtf.seen.map(s => s.file.declaredType)).to.deep.equal(['image/png', 'image/jpeg', 'image/png']);
    expect(rtf.seen.every(s => s.file.sniffedType === s.file.declaredType)).to.equal(true);
  });

  it('keeps a picture no plugin takes, and drops one a plugin removes with every reference to it', async () => {
    const none: ContainedFilePlugin = { kind: 'file', name: 'none', accepts: () => false, process: async () => 'passed' };
    const kept = await disarmOffice(fixture('media', 'pictures.docx'), { filePlugins: [none] });
    expect([kept.status, (await entryNames(must(kept.bytes, 'output'))).filter(n => /media|thumbnail/.test(n))]).to.deep.equal([
      'clean',
      ['word/media/image1.png', 'word/media/image2.jpeg', 'docProps/thumbnail.jpeg'],
    ]);
    const removed = await disarmOffice(fixture('media', 'pictures.docx'), { filePlugins: [plugin(() => 'removed').plugin] });
    const out = must(removed.bytes, 'output');
    expect(await structureProblems(out)).to.deep.equal([]);
    expect((await entryNames(out)).filter(n => /media|thumbnail/.test(n))).to.deep.equal([]);
    const doc = must(await entryText(out, 'word/document.xml'), 'document');
    expect(doc).to.not.match(/a:blip|v:imagedata/);
    expect(await entryText(out, '_rels/.rels')).to.not.include('thumbnail');
    // The printer settings and the font, which no plugin saw, stay.
    expect(await entryNames(out)).to.include.members(['word/printerSettings/printerSettings1.bin', 'word/fonts/font1.odttf']);
    const odt = await disarmOffice(fixture('media', 'pictures.odt'), { filePlugins: [plugin(() => 'removed').plugin] });
    const odtOut = must(odt.bytes, 'output');
    expect(await structureProblems(odtOut)).to.deep.equal([]);
    expect((await entryNames(odtOut)).filter(n => /^(Pictures|Thumbnails)\//.test(n))).to.deep.equal([]);
    expect(await entryText(odtOut, 'content.xml')).to.not.include('draw:image');
  });

  it('puts the bytes a plugin wrote in place of a picture, in packages and in RTF', async () => {
    // Cleans PNG pictures and passes everything else as it is.
    const scrub = (): { plugin: ContainedFilePlugin } => ({
      plugin: {
        kind: 'file',
        name: 'png',
        accepts: () => true,
        async process(file, sink) {
          if (file.sniffedType !== 'image/png') return 'passed';
          await sink.write(Buffer.from('cleaned picture'));
          return 'scrubbed';
        },
      },
    });
    const docx = await disarmOffice(fixture('media', 'pictures.docx'), { filePlugins: [scrub().plugin] });
    const out = must(docx.bytes, 'output');
    expect([docx.status, await entryText(out, 'word/media/image1.png')]).to.deep.equal(['defused', 'cleaned picture']);
    expect(kinds(must(docx.after, 'after')).filter(k => k === 'EMBEDDED/PLUGIN_SCRUBBED')).to.have.length(1);
    const rtf = await disarmOffice(fixture('media', 'pictures.rtf'), { filePlugins: [scrub().plugin] });
    const text = Buffer.from(must(rtf.bytes, 'output')).toString('latin1');
    const hex = Buffer.from('cleaned picture').toString('hex');
    // Both PNG pictures, the one in hex and the one after \bin, now hold the plugin's bytes in hex.
    expect(text.split(hex).length - 1).to.equal(2);
    expect(text).to.not.include('\\bin');
    expect(text).to.include('\\wmetafile8');
    const removed = await disarmOffice(fixture('media', 'pictures.rtf'), { filePlugins: [plugin(() => 'removed').plugin] });
    const left = Buffer.from(must(removed.bytes, 'output')).toString('latin1');
    expect(left.match(/\\pict\\/g)).to.deep.equal(['\\pict\\']);
  });
});
