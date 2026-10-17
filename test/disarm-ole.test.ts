import { expect } from 'chai';
import { type ContainedFilePlugin, disarmOffice, inspectOffice, type OfficeDisarmResult, officePlugin } from '../src';
import { type CfbEntry, CompoundFile, entryPath } from '../src/cfb';
import { type CfbNode, writeCfb } from '../src/cfb-writer';
import { bufferSink, bufferSource } from '../src/io';
import { decompressRtf } from '../src/rtfcomp';
import { fixture, kinds, must, of, zipOf } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const open = (bytes: Uint8Array) => CompoundFile.open(bufferSource(bytes), LIMITS);
const paths = async (bytes: Uint8Array) => (await open(bytes)).entries.slice(1).map(entryPath);

/** Disarms `bytes` and checks that the output reads as a compound file with nothing hidden in it. */
async function defused(bytes: Uint8Array, options = {}): Promise<OfficeDisarmResult & { bytes: Uint8Array }> {
  const r = await disarmOffice(bytes, options);
  const out = must(r.bytes, `output (${r.status}: ${r.before.findings.map(f => `${f.category}/${f.detail} ${JSON.stringify(f.data ?? {})}`).join('; ')})`);
  if (out[0] === 0xd0) expect((await open(out)).notes).to.deep.equal([]);
  return { ...r, bytes: out };
}

/** A plugin that takes every file it is shown and answers as `answer` says, writing `bytes` when given. */
function plugin(answer: 'passed' | 'scrubbed' | 'removed', bytes?: (seen: Uint8Array) => Uint8Array): ContainedFilePlugin & { seen: Array<{ name?: string; bytes: Buffer }> } {
  const seen: Array<{ name?: string; bytes: Buffer }> = [];
  return {
    kind: 'file',
    name: 'test',
    seen,
    accepts: () => true,
    async process(file, sink) {
      const b = Buffer.from(await file.source.read(0, file.size));
      seen.push({ name: file.name, bytes: b });
      if (bytes) await sink.write(bytes(b));
      return answer;
    },
  };
}

describe('compound file writer', () => {
  it('writes trees the reader reads back exactly, with the mini stream, large streams and a DIFAT', async () => {
    const stream = (name: string, size: number, fill: number): CfbNode => ({ kind: 'stream', name, size, source: bufferSource(Buffer.alloc(size, fill)) });
    const many = Array.from({ length: 40 }, (_, i) => stream(`S${String(i).padStart(2, '0')}`, i * 37, i));
    const tree = {
      clsid: '00020906-0000-0000-C000-000000000046',
      children: [stream('Empty', 0, 0), stream('Big', 9 * 1024 * 1024, 7), { kind: 'storage' as const, name: 'Store', clsid: '', children: many }, stream('Small', 100, 1)],
    };
    const sink = bufferSink();
    await writeCfb(tree, sink);
    const bytes = Buffer.from(sink.result());
    const cfb = await open(bytes);
    expect(cfb.notes).to.deep.equal([]);
    expect(cfb.root.clsid).to.equal('00020906-0000-0000-C000-000000000046');
    expect(cfb.root.children.map(c => [c.name, c.size])).to.have.deep.members([
      ['Empty', 0],
      ['Big', 9 * 1024 * 1024],
      ['Store', 0],
      ['Small', 100],
    ]);
    const store = must(cfb.child(cfb.root, 'Store'), 'Store');
    for (const [i, c] of [...store.children].sort((a, b) => a.name.localeCompare(b.name)).entries()) {
      expect(Buffer.from(await cfb.bytes(c)).equals(Buffer.alloc(i * 37, i)), c.name).to.equal(true);
    }
    expect(Buffer.from(await cfb.bytes(must(cfb.child(cfb.root, 'Big'), 'Big'))).equals(Buffer.alloc(9 * 1024 * 1024, 7))).to.equal(true);
    // The FAT of a 9 MiB file takes more sectors than the header lists.
    expect(bytes.readUInt32LE(0x48)).to.be.greaterThan(0);
  });

  it('keeps stream data out of the first 1024 bytes and ZIP end records out of the last 65,557', async () => {
    // A PDF in a small stream would sit where PDF readers look for a header, and a stream that ends with a ZIP end
    // record where ZIP readers look for one. Either would make the copy open as a second format.
    const pdf = Buffer.from('%PDF-1.7\n% an attachment\n');
    const pkg = await zipOf([['word/document.xml', 'x']]);
    const bigPkg = Buffer.concat([Buffer.alloc(8192, 1), pkg]);
    const endRecord = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
    for (const [name, data] of [
      ['Attachment', pdf],
      ['Package', pkg],
      ['Package', bigPkg],
    ] as const) {
      const sink = bufferSink();
      await writeCfb({ clsid: '', children: [{ kind: 'stream', name, size: data.length, source: bufferSource(data) }] }, sink);
      const bytes = Buffer.from(sink.result());
      expect([bytes.subarray(0, 1024).includes('%PDF-'), bytes.subarray(Math.max(0, bytes.length - 65557)).includes(endRecord)], `${name} of ${data.length}`).to.deep.equal([false, false]);
      const cfb = await open(bytes);
      expect(cfb.notes).to.deep.equal([]);
      expect(Buffer.from(await cfb.bytes(cfb.root.children[0])).equals(data)).to.equal(true);
    }
    // Free sectors are added only when an end record needs them.
    const sink = bufferSink();
    await writeCfb({ clsid: '', children: [{ kind: 'stream', name: 'Plain', size: 100, source: bufferSource(Buffer.alloc(100, 7)) }] }, sink);
    expect(sink.result().length).to.equal(5 * 512);
  });

  it('colors each storage as a valid red-black tree in MS-CFB name order', async () => {
    for (const n of [1, 2, 3, 6, 7, 8, 31]) {
      const children: CfbNode[] = Array.from({ length: n }, (_, i) => ({
        kind: 'stream',
        name: 'x'.repeat((i % 5) + 1) + String.fromCharCode(65 + i),
        size: 0,
        source: bufferSource(new Uint8Array(0)),
      }));
      const sink = bufferSink();
      await writeCfb({ clsid: '', children }, sink);
      const b = Buffer.from(sink.result());
      const entry = (id: number) => b.subarray(512 + 512 * b.readUInt32LE(0x30) + id * 128);
      const name = (id: number) =>
        entry(id)
          .subarray(0, entry(id).readUInt16LE(0x40) - 2)
          .toString('utf16le');
      const order = (a: string, c: string) => a.length - c.length || (a.toUpperCase() < c.toUpperCase() ? -1 : a.toUpperCase() > c.toUpperCase() ? 1 : 0);
      // Returns the black height, checking order and colors on the way.
      const walk = (id: number, lo?: string, hi?: string, parentRed = false): number => {
        if (id === 0xffffffff) return 1;
        const e = entry(id);
        const red = e[0x43] === 0;
        expect(red && parentRed, `red under red, ${n}`).to.equal(false);
        if (lo !== undefined) expect(order(lo, name(id))).to.be.lessThan(0);
        if (hi !== undefined) expect(order(name(id), hi)).to.be.lessThan(0);
        const left = walk(e.readUInt32LE(0x44), lo, name(id), red);
        const right = walk(e.readUInt32LE(0x48), name(id), hi, red);
        expect(left, `black height, ${n}`).to.equal(right);
        return left + (red ? 0 : 1);
      };
      walk(entry(0).readUInt32LE(0x4c));
    }
  });
});

describe('disarm: compound files', () => {
  it('drops VBA projects, objects, controls and attachments with their storages', async () => {
    const vba = await defused(fixture('ole', 'vba.doc'));
    expect([vba.status, (await paths(vba.bytes)).filter(p => /^Macros/.test(p)), kinds(must(vba.after, 'after'))]).to.deep.equal(['defused', [], []]);
    const objects = await defused(fixture('ole', 'objects.doc'));
    expect((await paths(objects.bytes)).filter(p => p.startsWith('ObjectPool/'))).to.deep.equal([]);
    expect((await paths(objects.bytes)).includes('ObjectPool')).to.equal(true);
    const xls = await defused(fixture('ole', 'objects.xls'));
    expect((await paths(xls.bytes)).filter(p => /^(MBD|Ctls)/.test(p))).to.deep.equal([]);
  });

  it('drops attachments and lowers the count the message keeps of them', async () => {
    // The message, with its property stream counting its five attachments as Outlook counts them.
    const original = await open(fixture('ole', 'message.msg'));
    const props = must(original.child(original.root, '__properties_version1.0'), 'properties');
    const counted = Buffer.from(await original.bytes(props));
    counted.writeUInt32LE(5, 20);
    const copy = (e: CfbEntry): CfbNode =>
      e.kind === 'stream'
        ? { kind: 'stream', name: e.name, size: e === props ? counted.length : e.size, source: e === props ? bufferSource(counted) : original.source(e) }
        : { kind: 'storage', name: e.name, clsid: '', children: e.children.map(copy) };
    const sink = bufferSink();
    await writeCfb({ clsid: '', children: original.root.children.map(copy) }, sink);
    const input = sink.result();
    const count = async (b: Uint8Array) => {
      const cfb = await open(b);
      return Buffer.from(await cfb.bytes(must(cfb.child(cfb.root, '__properties_version1.0'), 'properties'))).readUInt32LE(20);
    };
    const r = await defused(input);
    expect([(await paths(r.bytes)).filter(p => /^__attach/.test(p)), await count(r.bytes), must(r.after, 'after').format]).to.deep.equal([[], 0, 'msg']);
    // A plugin that keeps the attachment held by value keeps one in the count.
    const kept = await defused(input, { filePlugins: [plugin('passed')] });
    expect([(await paths(kept.bytes)).filter(p => /^__attach_version1\.0_#[0-9A-F]{8}$/.test(p)), await count(kept.bytes)]).to.deep.equal([['__attach_version1.0_#00000000'], 1]);
  });

  it('leaves out hidden data and data past the last sector, and writes a version 4 file as version 3', async () => {
    for (const file of ['hidden-data.doc', 'trailing.doc', 'mini-slack.doc', 'header-padding.doc']) {
      const r = await defused(fixture('cfb', file));
      expect([file, r.status, kinds(must(r.after, 'after'))]).to.deep.equal([file, 'defused', []]);
    }
    const v4 = await defused(fixture('cfb', 'version-4.doc'));
    expect([v4.status, Buffer.from(v4.bytes).readUInt16LE(0x1a), await paths(v4.bytes)]).to.deep.equal(['clean', 3, await paths(fixture('cfb', 'version-4.doc'))]);
  });

  it('rejects a file whose findings lie inside a stream, which the rewrite does not change', async () => {
    // Excel's and PowerPoint's records are rewritten, but not Word's fields, Excel's cell formulas, nor a workbook whose
    // sheet starts where no record does.
    for (const file of ['fields.doc', 'formula-functions.xls', 'xls-variants.xls']) {
      const r = await disarmOffice(fixture('ole', file));
      expect([file, r.status, r.bytes, of(r.before, 'PROCESSING/VERIFICATION_FAILED').length]).to.deep.equal([file, 'rejected', undefined, 1]);
      // The upload's own verdict keeps its score.
      expect(r.before.score).to.be.a('number');
    }
  });

  it('shows embedded objects, Packager files and attachments to the plugins, and applies their verdicts', async () => {
    const seen = plugin('passed');
    const kept = await defused(fixture('ole', 'objects.doc'), { filePlugins: [seen] });
    expect(seen.seen.map(s => s.name)).to.deep.equal(['inert.txt', 'Microsoft Equation 3.0']);
    // The plugin passed both objects, and its answer is final, so both stay and office-defuse reports nothing of its own
    // about them, not even that one is an Equation Editor object.
    expect((await paths(kept.bytes)).filter(p => /^ObjectPool\/_100[01]$/.test(p))).to.deep.equal(['ObjectPool/_1000', 'ObjectPool/_1001']);
    expect([...of(kept.before, 'EMBEDDED/OLE_OBJECT'), ...of(kept.before, 'EMBEDDED/EQUATION')]).to.deep.equal([]);
    expect(of(kept.before, 'EMBEDDED/PLUGIN_PASSED').length).to.equal(2);

    const scrubbed = await defused(fixture('ole', 'objects.doc'), { filePlugins: [plugin('scrubbed', b => (b[0] === 0xd0 ? b : Buffer.from('cleaned')))] });
    const cfb = await open(scrubbed.bytes);
    const native = must(
      cfb.entries.find(e => entryPath(e) === 'ObjectPool/_1000/\\x01Ole10Native'),
      'native',
    );
    expect(Buffer.from(await cfb.bytes(native)).includes(Buffer.from('cleaned'))).to.equal(true);
    expect(kinds(must(scrubbed.after, 'after'))).to.include('EMBEDDED/PLUGIN_SCRUBBED');

    const attachments = await defused(fixture('ole', 'msg-variants.msg'), { filePlugins: [plugin('scrubbed', () => Buffer.from('replaced'))] });
    const msg = await open(attachments.bytes);
    const data = msg.entries.filter(e => /__substg1\.0_37010102$/i.test(e.name));
    expect(data.length).to.be.greaterThan(0);
    for (const d of data) expect(Buffer.from(await msg.bytes(d)).toString()).to.equal('replaced');

    const removed = await defused(fixture('ole', 'objects.doc'), { filePlugins: [plugin('removed')] });
    expect((await paths(removed.bytes)).filter(p => p.startsWith('ObjectPool/_'))).to.deep.equal([]);
  });

  it('puts the storages of the compound file a plugin returns in place of an object', async () => {
    const r = await defused(fixture('cases', 'ole-embedded-package.doc'), { filePlugins: [officePlugin()] });
    // A bare object storage holds no Office document, so officePlugin does not take it, and the object's own rule
    // removes it.
    expect(r.before.findings.map(f => [`${f.category}/${f.detail}`, f.action, f.attachment])).to.deep.equal([['EMBEDDED/OLE_OBJECT', 'strip', undefined]]);
    const graft: ContainedFilePlugin = {
      kind: 'file',
      name: 'graft',
      accepts: f => f.sniffedType === 'application/x-ole-storage',
      async process(_file, sink) {
        await writeCfb({ clsid: '', children: [{ kind: 'stream', name: 'Replaced', size: 3, source: bufferSource(Buffer.from('new')) }] }, sink);
        return 'scrubbed';
      },
    };
    const grafted = await defused(fixture('cases', 'ole-embedded-package.doc'), { filePlugins: [graft] });
    expect((await paths(grafted.bytes)).filter(p => p.startsWith('ObjectPool/_1000'))).to.deep.equal(['ObjectPool/_1000', 'ObjectPool/_1000/Replaced']);
    expect(kinds(must(grafted.after, 'after'))).to.include('EMBEDDED/PLUGIN_SCRUBBED');
  });
});

describe('disarm: RTF', () => {
  const text = (b: Uint8Array) => Buffer.from(b).toString('latin1');

  it('drops field instructions and keeps results, and drops the template, shape hyperlinks and trailing text', async () => {
    const r = await defused(fixture('rtf', 'fields.rtf'));
    const out = text(r.bytes);
    expect(out).to.not.match(/fldinst \{DDE|INCLUDEPICTURE|MACROBUTTON|\\template/);
    expect(out).to.include('{\\field{\\fldrslt x}}');
    expect(out).to.include('{\\*\\fldinst HYPERLINK "https://example.invalid/page"}');
    const trailing = await defused(fixture('rtf', 'trailing.rtf'));
    expect(text(trailing.bytes).endsWith('Inert text.\\par}')).to.equal(true);
  });

  it('puts an object result in place of the object, and drops \\objupdate alone from an object a plugin kept', async () => {
    const r = await defused(fixture('rtf', 'equation.rtf'));
    expect(text(r.bytes)).to.include('{{\\pict');
    expect(text(r.bytes)).to.not.include('objdata');
    const kept = await disarmOffice(fixture('rtf', 'package.rtf'), { filePlugins: [plugin('passed')] });
    expect([kept.status, of(kept.before, 'EMBEDDED/OLE_OBJECT').length, of(kept.before, 'EMBEDDED/PLUGIN_PASSED').length]).to.deep.equal(['clean', 0, 1]);
    expect(text(must(kept.bytes, 'output'))).to.include('\\objdata');
  });

  it('writes the bytes a plugin returned into the object data, and reads them back from the output', async () => {
    const p = plugin('scrubbed', () => Buffer.from('cleaned file'));
    const r = await defused(fixture('rtf', 'package.rtf'), { filePlugins: [p] });
    expect(p.seen.map(s => [s.name, s.bytes.toString()])).to.deep.equal([['inert.txt', 'inert marker']]);
    expect(text(r.bytes)).to.include(Buffer.from('cleaned file').toString('hex'));
    expect(kinds(must(r.after, 'after'))).to.deep.equal(['EMBEDDED/PLUGIN_SCRUBBED']);
    const again = await inspectOffice(r.bytes);
    expect(of(again, 'EMBEDDED/OLE_OBJECT')[0]?.data?.size).to.equal(12);
  });
});

describe('disarm: Outlook message bodies', () => {
  /** The stream at `path` in a compound file, or undefined. */
  const stream = async (bytes: Uint8Array, path: string) => {
    const cfb = await open(bytes);
    const e = cfb.entries.find(x => entryPath(x).toUpperCase() === path.toUpperCase());
    return e ? Buffer.from(await cfb.bytes(e)) : undefined;
  };
  /** The size the root property stream gives a property held in a stream, or undefined when it has no entry. */
  const declared = async (bytes: Uint8Array, tag: number) => {
    const props = must(await stream(bytes, '__properties_version1.0'), 'property stream');
    for (let at = 32; at + 16 <= props.length; at += 16) if (props.readUInt32LE(at) === tag) return props.readUInt32LE(at + 8);
    return undefined;
  };
  const RTF = '__substg1.0_10090102';
  const HTML = '__substg1.0_10130102';

  it('reads the RTF body with the RTF rules, rewrites it without what they strip, and stores it uncompressed', async () => {
    const r = await defused(fixture('ole', 'message-rtf-body.msg'));
    expect([r.status, r.before.findings.map(f => [`${f.category}/${f.detail}`, f.location, f.attachment])]).to.deep.equal(['defused', [['DDE/FIELD', `${RTF} > document`, 'RTF body']]]);
    const body = must(decompressRtf(must(await stream(r.bytes, RTF), 'RTF body'), 1 << 20), 'decompressed body').toString();
    expect([body.includes('DDEAUTO'), body.includes('result'), body.includes('kept text')]).to.deep.equal([false, true, true]);
    expect(await declared(r.bytes, 0x10090102)).to.equal(must(await stream(r.bytes, RTF), 'RTF body').length);
    expect(kinds(must(r.after, 'after'))).to.deep.equal([]);
  });

  it('removes an RTF body that does not decompress, with its property entry', async () => {
    const r = await defused(fixture('ole', 'message-rtf-crc.msg'));
    expect(kinds(r.before)).to.deep.equal(['PACKAGE/UNREADABLE_BODY']);
    expect([await stream(r.bytes, RTF), await declared(r.bytes, 0x10090102)]).to.deep.equal([undefined, undefined]);
  });

  it('shows the RTF and HTML bodies to the plugins, and applies their verdicts', async () => {
    const seen = plugin('removed');
    const removed = await defused(fixture('ole', 'message-rtf-body.msg'), { filePlugins: [seen] });
    expect([seen.seen.map(s => [s.name, s.bytes.subarray(0, 6).toString()]), await stream(removed.bytes, RTF), await declared(removed.bytes, 0x10090102)]).to.deep.equal([
      [['RTF body', '{\\rtf1']],
      undefined,
      undefined,
    ]);
    // A plugin's copy of the RTF body is stored uncompressed, which MS-OXRTFCP allows.
    const copy = Buffer.from('{\\rtf1\\ansi plain}');
    const scrubbed = await defused(fixture('ole', 'message-rtf-body.msg'), { filePlugins: [plugin('scrubbed', () => copy)] });
    const stored = must(await stream(scrubbed.bytes, RTF), 'RTF body');
    expect([stored.readUInt32LE(8), decompressRtf(stored, 1 << 20)?.toString(), await declared(scrubbed.bytes, 0x10090102), kinds(must(scrubbed.after, 'after'))]).to.deep.equal([
      0x414c454d,
      copy.toString(),
      stored.length,
      ['EMBEDDED/PLUGIN_SCRUBBED'],
    ]);
    const html = plugin('scrubbed', () => Buffer.from('<p>clean</p>'));
    const r = await defused(fixture('ole', 'message-html-body.msg'), { filePlugins: [html] });
    expect([html.seen.map(s => s.name), (await stream(r.bytes, HTML))?.toString(), await declared(r.bytes, 0x10130102)]).to.deep.equal([['HTML body'], '<p>clean</p>', 12]);
  });

  it("reads an embedded message's RTF body too", async () => {
    const r = await inspectOffice(fixture('ole', 'message-embedded-body.msg'));
    expect(of(r, 'DDE/FIELD').map(f => [f.location, f.attachment])).to.deep.equal([['__attach_version1.0_#00000000/__substg1.0_3701000D/__substg1.0_10090102 > document', 'RTF body']]);
  });
});
