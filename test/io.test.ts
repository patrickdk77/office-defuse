import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { PassThrough, Writable } from 'node:stream';
import { expect } from 'chai';
import { disarmOffice, fileSink, writableSink } from '../src';
import { SpillSink, TempDir } from '../src/io';
import { fixture, must } from './helpers/util';

describe('sinks and temporary files', () => {
  it('writes a file, and removes a partial one on abort', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-sink-'));
    try {
      const file = path.join(dir, 'out.bin');
      const sink = fileSink(file);
      await sink.write(Buffer.from('abc'));
      await sink.close();
      expect(fs.readFileSync(file, 'latin1')).to.equal('abc');
      const empty = fileSink(path.join(dir, 'empty.bin'));
      await empty.close();
      expect(fs.statSync(path.join(dir, 'empty.bin')).size).to.equal(0);
      const partial = fileSink(path.join(dir, 'partial.bin'));
      await partial.write(Buffer.from('x'));
      await must(partial.abort, 'abort')(new Error('stop'));
      expect(fs.existsSync(path.join(dir, 'partial.bin'))).to.equal(false);
      await must(fileSink(path.join(dir, 'never.bin')).abort, 'abort')(new Error('stop'));
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('writes to a stream with backpressure, and fails on the stream error', async () => {
    const pass = new PassThrough({ highWaterMark: 4 });
    const chunks: Buffer[] = [];
    pass.on('data', c => chunks.push(c));
    const sink = writableSink(pass);
    await sink.write(Buffer.from('0123456789'));
    await sink.close();
    expect(Buffer.concat(chunks).toString()).to.equal('0123456789');
    const broken = new Writable({ write: (_c, _e, done) => done(new Error('broken pipe')) });
    const failing = writableSink(broken);
    let thrown: unknown;
    try {
      await failing.write(Buffer.from('x'));
    } catch (e) {
      thrown = e;
    }
    expect((thrown as Error).message).to.equal('broken pipe');
    const destroyed = new PassThrough();
    const later = writableSink(destroyed);
    await must(later.abort, 'abort')(new Error('stop'));
    let refused: unknown;
    try {
      await later.write(Buffer.from('x'));
    } catch (e) {
      refused = e;
    }
    expect(refused).to.be.an('error');
  });

  it('spills past its threshold to a temporary file, which the run removes', async () => {
    const temp = new TempDir();
    const sink = new SpillSink(temp, 4);
    await sink.write(Buffer.from('ab'));
    await sink.write(Buffer.from('cdef'));
    await sink.close();
    expect([sink.spilled, sink.length]).to.deep.equal([true, 6]);
    const read: Buffer[] = [];
    for await (const c of sink.read(4)) read.push(Buffer.from(c));
    expect(Buffer.concat(read).toString()).to.equal('abcdef');
    await sink.dispose();
    await temp.cleanup();
  });

  it('defuses through temporary files when the memory threshold is zero, and leaves none behind', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-temp-'));
    try {
      const passed = { kind: 'file' as const, name: 'pass', accepts: () => true, process: async () => 'passed' as const };
      const r = await disarmOffice(fixture('docs', 'objects.docx'), { tempDir: dir, memoryThreshold: 0, filePlugins: [passed] });
      expect(r.status).to.equal('defused');
      expect(fs.readdirSync(dir)).to.deep.equal([]);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
