import * as zlib from 'node:zlib';
import { expect } from 'chai';
import { detectOffice, disarmOffice, type ScriptPlugin } from '../src';
import { CompoundFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-writer';
import { bufferSink, bufferSource } from '../src/io';
import { PURGED_VBA_PROJECT } from '../src/vba';
import { fixture, kinds, must, of } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const passAll: ScriptPlugin = { kind: 'script', name: 'test', accepts: () => true, process: async () => ({ result: 'passed' }) };

interface Rec {
  type: number;
  instance: number;
  offset: number;
  body: Buffer;
  kids: Rec[];
}

/** A record and, for a container, the records inside it. */
function parse(b: Buffer, offset: number, end = b.length): Rec[] {
  const out: Rec[] = [];
  for (let pos = offset; pos + 8 <= end; ) {
    const verInst = b.readUInt16LE(pos);
    const length = b.readUInt32LE(pos + 4);
    const body = b.subarray(pos + 8, pos + 8 + length);
    out.push({ type: b.readUInt16LE(pos + 2), instance: verInst >> 4, offset: pos, body, kids: (verInst & 0x0f) === 0x0f ? parse(b, pos + 8, pos + 8 + length) : [] });
    pos += 8 + length;
  }
  return out;
}
const all = (rs: Rec[]): Rec[] => rs.flatMap(r => [r, ...all(r.kids)]);

/** The presentation the output's current edit describes: its UserEditAtom, and its persist objects by id. */
async function presentation(bytes: Uint8Array) {
  const cfb = await CompoundFile.open(bufferSource(bytes), LIMITS);
  const stream = Buffer.from(await cfb.bytes(must(cfb.child(cfb.root, 'PowerPoint Document'), 'stream')));
  const user = Buffer.from(await cfb.bytes(must(cfb.child(cfb.root, 'Current User'), 'Current User')));
  const editAt = user.readUInt32LE(16);
  const edit = parse(stream, editAt, editAt + 8 + stream.readUInt32LE(editAt + 4))[0];
  const dirAt = edit.body.readUInt32LE(12);
  const dir = parse(stream, dirAt, dirAt + 8 + stream.readUInt32LE(dirAt + 4))[0].body;
  const objects = new Map<number, Rec>();
  for (let p = 0; p + 4 <= dir.length; ) {
    const entry = dir.readUInt32LE(p);
    p += 4;
    for (let k = 0; k < entry >>> 20; k++, p += 4) {
      const at = dir.readUInt32LE(p);
      objects.set((entry & 0xfffff) + k, parse(stream, at, at + 8 + stream.readUInt32LE(at + 4))[0]);
    }
  }
  return { token: user.readUInt32LE(12), edit, objects, records: all([...objects.values()]) };
}

describe('PowerPoint rewrite', () => {
  it('writes a full save without the storages and records the findings remove, and actions that do nothing', async () => {
    const r = await disarmOffice(fixture('ole', 'objects.ppt'));
    expect([r.status, must(r.after, 'after').status]).to.deep.equal(['defused', 'clean']);
    const p = await presentation(must(r.bytes, 'output'));
    expect(p.edit.body.readUInt32LE(8)).to.equal(0);
    // The object, the VBA project and the control go, and so do the records that name them.
    expect([...p.objects.values()].filter(o => o.type === 0x1011)).to.deep.equal([]);
    expect(p.records.filter(x => [0x03ff, 0x0fcc, 0x0fee].includes(x.type))).to.deep.equal([]);
    // Each container's length is what its children add up to.
    for (const x of p.records) if (x.kids.length) expect(x.kids.reduce((n, k) => n + 8 + k.body.length, 0)).to.equal(x.body.length);
    const actions = p.records.filter(x => x.type === 0x0ff3).map(x => x.body[8]);
    expect(actions).to.deep.equal([0, 0, 0]);
    // The hyperlink to a web page stays; the movie keeps its container and loses its path.
    expect(p.records.filter(x => x.type === 0x0fd7)).to.have.length(1);
    const movie = must(
      p.records.find(x => x.type === 0x1005),
      'movie',
    );
    expect(movie.kids.find(k => k.type === 0x0fba)?.body.length).to.equal(0);
  });

  it('writes a presentation VelvetSweatshop opens decrypted, without its encryption session', async () => {
    const r = await disarmOffice(fixture('ole', 'ppt-default.ppt'));
    expect([r.status, ...kinds(must(r.after, 'after'))]).to.deep.equal(['defused', 'LINK/SAFE']);
    const p = await presentation(must(r.bytes, 'output'));
    expect(p.token).to.equal(0xe391c05f);
    expect(p.edit.body.length).to.equal(28);
    expect([...p.objects.values()].some(o => o.type === 0x2f14)).to.equal(false);
  });

  it('rejects an encrypted presentation with pictures rather than decrypt only part of it, since its Pictures stream is encrypted too', async () => {
    const cfb = await CompoundFile.open(bufferSource(fixture('ole', 'ppt-default.ppt')), LIMITS);
    const { subtree } = await import('../src/rewrite-cfb');
    const tree = subtree(cfb, cfb.root);
    const pictures = Buffer.from('inert picture bytes');
    const sink = bufferSink();
    await writeCfb({ clsid: tree.clsid, children: [...tree.children, { kind: 'stream', name: 'Pictures', size: pictures.length, source: bufferSource(pictures) }] }, sink);
    const r = await disarmOffice(sink.result());
    expect([r.status, ...of(r.before, 'PROCESSING/VERIFICATION_FAILED').map(f => f.data?.reasons)]).to.deep.equal(['rejected', 'ENCRYPTED/DEFAULT_PASSWORD']);
  });

  it('keeps a VBA project the plugins pass without its compiled code, and the action that runs its macro', async () => {
    const r = await disarmOffice(fixture('vba', 'vba.ppt'), { scriptPlugins: [passAll] });
    expect(of(r.before, 'MACRO/LINK').map(f => f.action)).to.deep.equal(['info']);
    expect(must(r.after, 'after').status).to.equal('clean');
    const p = await presentation(must(r.bytes, 'output'));
    const storage = must(p.objects.get(4), 'VBA storage');
    expect([storage.type, storage.instance]).to.deep.equal([0x1011, 1]);
    const inner = await CompoundFile.open(bufferSource(zlib.inflateSync(storage.body.subarray(4))), LIMITS);
    const compiled = must(inner.child(must(inner.child(inner.root, 'VBA'), 'VBA'), '_VBA_PROJECT'), '_VBA_PROJECT');
    expect(Buffer.from(await inner.bytes(compiled))).to.deep.equal(Buffer.from(PURGED_VBA_PROJECT));
    // The macro's action stays; the hyperlink to a local file goes, and so does the action that opens it.
    expect(p.records.filter(x => x.type === 0x0ff3).map(x => x.body[8])).to.deep.equal([1, 0]);
    expect(p.records.filter(x => x.type === 0x0fd7)).to.deep.equal([]);
    expect(p.records.filter(x => x.type === 0x03ff)).to.have.length(1);
  });

  it('reads a PowerPoint 97 add-in as one, and keeps it one without its project', async () => {
    expect(await detectOffice(fixture('vba', 'addin.ppa'))).to.deep.include({ extension: 'ppa', mimeType: 'application/vnd.ms-powerpoint' });
    const kept = await disarmOffice(fixture('vba', 'addin.ppa'), { scriptPlugins: [passAll] });
    expect([kept.status, must(kept.after, 'after').format, ...of(kept.before, 'MACRO/VBA_PROJECT').map(f => f.action)]).to.deep.equal(['defused', 'ppa', 'info']);
    const removed = await disarmOffice(fixture('vba', 'addin.ppa'));
    expect([removed.status, must(removed.after, 'after').format, ...of(removed.before, 'MACRO/VBA_PROJECT').map(f => f.action)]).to.deep.equal(['defused', 'ppa', 'strip']);
  });
});
