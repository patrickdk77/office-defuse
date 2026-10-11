import { spawnSync } from 'node:child_process';
import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { expect } from 'chai';
import { bufferSource } from '../src/io';
import { ZipArchive } from '../src/zip';
import { ZipWriter } from '../src/zip-writer';
import { bufferSink, failure, fixture, must } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 512 * 1024 * 1024, totalDecompressedBytes: 2 * 1024 * 1024 * 1024 };
const sha = (b: Uint8Array) => crypto.createHash('sha256').update(b).digest('hex');

// Reads a ZIP with Python's zipfile, checks every CRC, and prints each entry's name and SHA-256.
const PYTHON = `
import hashlib, sys, zipfile
with zipfile.ZipFile(sys.argv[1]) as z:
    bad = z.testzip()
    if bad is not None:
        sys.exit('bad entry ' + bad)
    for i in z.infolist():
        sys.stdout.buffer.write((i.filename + '\\t' + hashlib.sha256(z.read(i)).hexdigest() + '\\n').encode('utf-8'))
`;

const entries: Array<[string, Buffer, boolean]> = [
  ['mimetype', Buffer.from('application/vnd.oasis.opendocument.text'), false],
  ['folder/', Buffer.alloc(0), false],
  ['folder/text.xml', Buffer.from('<a>'.repeat(5000)), true],
  ['caf\xe9/na\xefve.txt', Buffer.from('non-ASCII name'), true],
  ['random.bin', crypto.randomBytes(70_000), true],
  ['empty.txt', Buffer.alloc(0), true],
];

async function write(options: { forceZip64?: boolean } = {}): Promise<Buffer> {
  const sink = bufferSink();
  const w = new ZipWriter(sink, options);
  for (const [name, data, compress] of entries) await w.add(name, data, { compress });
  await w.finish();
  // The writer leaves the sink open for its caller.
  expect(sink.closed).to.equal(false);
  return sink.result();
}

describe('ZIP writer', () => {
  let dir: string;
  before(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-writer-'));
  });
  after(() => fs.rmSync(dir, { recursive: true, force: true }));

  for (const forceZip64 of [false, true]) {
    const label = forceZip64 ? 'with ZIP64 records everywhere' : 'with plain records';
    it(`writes an archive its own reader accepts, ${label}`, async () => {
      const zip = await write({ forceZip64 });
      const z = await ZipArchive.open(bufferSource(zip), LIMITS);
      expect(z.notes).to.deep.equal([]);
      expect(z.entries.map(e => e.name)).to.deep.equal(entries.map(e => e[0]));
      for (const [name, data] of entries) {
        const chunks: Uint8Array[] = [];
        for await (const c of z.read(must(z.get(name), name))) chunks.push(c);
        expect(sha(Buffer.concat(chunks)), name).to.equal(sha(data));
      }
      expect(must(z.get('mimetype'), 'mimetype').method).to.equal(0);
      expect(must(z.get('mimetype'), 'mimetype').localExtra).to.equal(forceZip64);
    });

    it(`writes an archive Python's zipfile and unzip -t read, ${label}`, async () => {
      const file = path.join(dir, `out-${forceZip64}.zip`);
      fs.writeFileSync(file, await write({ forceZip64 }));
      const py = spawnSync('python3', ['-I', '-c', PYTHON, file], { encoding: 'utf8' });
      expect(py.status, py.stderr).to.equal(0);
      expect(py.stdout.trim().split('\n')).to.deep.equal(entries.map(([name, data]) => `${name}\t${sha(data)}`));
      const unzip = spawnSync('unzip', ['-t', file], { encoding: 'utf8' });
      expect(unzip.status, unzip.stdout + unzip.stderr).to.equal(0);
      expect(unzip.stdout).to.match(/No errors detected/);
    });
  }

  it('copies compressed entries unchanged from an archive it reads', async () => {
    const src = await ZipArchive.open(bufferSource(fixture('lo', 'clean.docx')), LIMITS);
    const sink = bufferSink();
    const w = new ZipWriter(sink);
    for (const e of src.entries) await w.copy(e, src.raw(e));
    await w.finish();
    const copy = await ZipArchive.open(bufferSource(sink.result()), LIMITS);
    expect(copy.entries.map(e => [e.name, e.crc, e.compressedSize, e.size, e.dosTime, e.dosDate])).to.deep.equal(src.entries.map(e => [e.name, e.crc, e.compressedSize, e.size, e.dosTime, e.dosDate]));
    for (const e of copy.entries) {
      for await (const _ of copy.read(e)) {
        // Reading checks the size and CRC.
      }
    }
  });

  it('refuses a copy whose bytes do not match its compressed size, and a name given twice', async () => {
    const w = new ZipWriter(bufferSink());
    const short = w.copy(
      { name: 'a', method: 0, crc: 0, compressedSize: 5, size: 5 },
      (async function* () {
        yield Buffer.from('abc');
      })(),
    );
    expect(String(await failure(short))).to.match(/3 compressed bytes, not 5/);
    await w.add('b', Buffer.from('x'));
    expect(String(await failure(w.add('b', Buffer.from('y'))))).to.match(/already in the archive/);
  });
});
