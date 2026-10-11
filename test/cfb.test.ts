import { expect } from 'chai';
import { inspectOffice } from '../src';
import { CompoundFile, entryPath, guid, showName } from '../src/cfb';
import { bufferSource } from '../src/io';
import { fixture, kinds, must, of } from './helpers/util';

const cfb = (name: string) => fixture('cfb', name);
const LIMITS = { entries: 10_000, decompressedBytes: 512 * 1024 * 1024, totalDecompressedBytes: 2 * 1024 * 1024 * 1024 };

/** The finding a refused file stops at, with its reason. */
async function refusal(bytes: Uint8Array): Promise<string> {
  const r = await inspectOffice(bytes);
  expect(r.status).to.equal('rejected');
  expect(r.risk).to.equal('UNKNOWN');
  const last = must(r.findings[r.findings.length - 1], 'finding');
  return `${last.category}/${last.detail}`;
}

describe('compound file reader', () => {
  it('reads version 3 and version 4 files, a DIFAT chain, and sizes whose high half a version 3 file leaves set', async () => {
    for (const name of ['valid.doc', 'version-4.doc', 'difat.doc', 'size-high-bits.doc', 'chain-longer.doc', 'partial-last-sector.doc']) {
      const r = await inspectOffice(cfb(name));
      expect({ name, status: r.status, format: r.format, kinds: kinds(r) }).to.deep.equal({ name, status: 'clean', format: 'doc', kinds: [] });
    }
  });

  const refused: Array<[string, string]> = [
    ['short.doc', 'CONTAINER/BAD_HEADER'],
    ['version-mismatch.doc', 'CONTAINER/BAD_HEADER'],
    ['mini-sector.doc', 'CONTAINER/BAD_HEADER'],
    ['mini-cutoff.doc', 'CONTAINER/BAD_HEADER'],
    ['byte-order.doc', 'CONTAINER/BAD_HEADER'],
    ['no-fat.doc', 'CONTAINER/BAD_HEADER'],
    ['fat-count.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['fat-past-count.doc', 'CONTAINER/BAD_HEADER'],
    ['difat-count.doc', 'CONTAINER/BAD_HEADER'],
    ['fat-reserved.doc', 'CONTAINER/BROKEN_CHAIN'],
    ['fat-outside.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['chain-loop.doc', 'CONTAINER/CHAIN_LOOP'],
    ['chain-free.doc', 'CONTAINER/BROKEN_CHAIN'],
    ['chain-short.doc', 'CONTAINER/DATA_ERROR'],
    ['chain-outside.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['chain-into-fat.doc', 'CONTAINER/OVERLAP'],
    ['shared-sector.doc', 'CONTAINER/OVERLAP'],
    ['directory-loop.doc', 'CONTAINER/CHAIN_LOOP'],
    ['no-directory.doc', 'CONTAINER/BAD_HEADER'],
    ['tree-loop.doc', 'CONTAINER/DIRECTORY_LOOP'],
    ['entry-twice.doc', 'CONTAINER/DIRECTORY_LOOP'],
    ['entry-outside.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['empty-entry.doc', 'CONTAINER/BAD_ENTRY'],
    ['second-root.doc', 'CONTAINER/BAD_ENTRY'],
    ['bad-type.doc', 'CONTAINER/BAD_ENTRY'],
    ['name-odd.doc', 'CONTAINER/BAD_ENTRY'],
    ['name-long.doc', 'CONTAINER/BAD_ENTRY'],
    ['name-nul.doc', 'CONTAINER/BAD_ENTRY'],
    ['name-unterminated.doc', 'CONTAINER/BAD_ENTRY'],
    ['name-slash.doc', 'CONTAINER/UNSAFE_NAME'],
    ['duplicate-name.doc', 'CONTAINER/DUPLICATE_NAME'],
    ['stream-children.doc', 'CONTAINER/BAD_ENTRY'],
    ['mini-loop.doc', 'CONTAINER/CHAIN_LOOP'],
    ['mini-free.doc', 'CONTAINER/BROKEN_CHAIN'],
    ['mini-short.doc', 'CONTAINER/DATA_ERROR'],
    ['mini-outside.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['mini-shared.doc', 'CONTAINER/OVERLAP'],
    ['mini-fat-count.doc', 'CONTAINER/BAD_HEADER'],
    ['size-past-limit.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['cut-last-sector.doc', 'CONTAINER/OUT_OF_RANGE'],
    ['cut-mid-stream.doc', 'CONTAINER/OUT_OF_RANGE'],
  ];
  for (const [name, kind] of refused) {
    it(`refuses ${name} as ${kind}`, async () => {
      expect(await refusal(cfb(name))).to.equal(kind);
    });
  }

  it('reports bytes no stream covers, and bytes past the sectors the FAT describes, as removable', async () => {
    const hidden = await inspectOffice(cfb('hidden-data.doc'));
    expect(kinds(hidden)).to.deep.equal(['CONTAINER/HIDDEN_DATA']);
    // A free sector, a sector of an orphan chain, the rest of a stream's last sector, a free mini sector and an
    // unreachable directory entry.
    expect(hidden.findings[0].data).to.deep.equal({ bytes: 512 + 512 + 120 + 64 + 128 });
    expect(kinds(await inspectOffice(cfb('mini-slack.doc')))).to.deep.equal(['CONTAINER/HIDDEN_DATA']);
    expect(of(await inspectOffice(cfb('header-padding.doc')), 'CONTAINER/HIDDEN_DATA')[0].data).to.deep.equal({ bytes: 3584 });
    const trailing = await inspectOffice(cfb('trailing.doc'));
    expect(trailing.findings.map(f => [f.category, f.detail, f.data])).to.deep.equal([['CONTAINER', 'TRAILING_BYTES', { bytes: 11795 }]]);
    expect(trailing.status).to.equal('strippable');
  });

  /** difat.doc with one 32-bit value changed: in the header, or in its DIFAT sector, which is sector 110. */
  const patched = (at: number, value: number) => {
    const b = Buffer.from(cfb('difat.doc'));
    b.writeUInt32LE(value >>> 0, at);
    return b;
  };
  const DIFAT = 512 * 111;

  it('refuses a DIFAT chain that loops, leaves the file, ends early, lists too many FAT sectors or does not end', async () => {
    expect(await refusal(patched(DIFAT + 508, 110))).to.equal('CONTAINER/BROKEN_CHAIN');
    // The header lists FAT sector 2 twice.
    expect(await refusal(patched(0x4c, 2))).to.equal('CONTAINER/CHAIN_LOOP');
    expect(await refusal(patched(0x44, 9000))).to.equal('CONTAINER/OUT_OF_RANGE');
    expect(await refusal(patched(0x44, 0xfffffffe))).to.equal('CONTAINER/BROKEN_CHAIN');
    expect(await refusal(patched(DIFAT + 4, 120))).to.equal('CONTAINER/BAD_HEADER');
    expect(await refusal(patched(0x48, 2))).to.equal('CONTAINER/BAD_HEADER');
  });

  it('refuses a DIFAT chain that names its own sector twice', async () => {
    // A FAT that needs two DIFAT sectors, the second of which is the first again, in a file long enough to hold them.
    const loop = Buffer.concat([cfb('difat.doc'), Buffer.alloc(120 * 512)]);
    loop.writeUInt32LE(110 + 127, 0x2c);
    loop.writeUInt32LE(2, 0x48);
    loop.writeUInt32LE(110, DIFAT + 508);
    expect(await refusal(loop)).to.equal('CONTAINER/CHAIN_LOOP');
  });

  it('applies the entry, stream size and total size limits', async () => {
    const limited = async (limits: Record<string, number>) => {
      const r = await inspectOffice(cfb('valid.doc'), { limits });
      return `${r.findings[0].category}/${r.findings[0].detail}`;
    };
    expect(await limited({ entries: 2 })).to.equal('LIMIT/ENTRY_COUNT');
    expect(await limited({ decompressedBytes: 1000 })).to.equal('LIMIT/DECOMPRESSED_SIZE');
    const mini = await inspectOffice(cfb('valid.doc'), { limits: { decompressedBytes: 100 } });
    expect(mini.findings[0].data?.reason).to.match(/^The mini stream is \d+ bytes$/);
    expect(await limited({ totalDecompressedBytes: 9000 })).to.equal('LIMIT/TOTAL_DECOMPRESSED_SIZE');
    expect(await limited({ timeMs: -1 })).to.equal('LIMIT/TIME');
  });

  it('reads streams through their chains, in the mini stream or in sectors, and names entries for locations', async () => {
    const c = await CompoundFile.open(bufferSource(cfb('valid.doc')), LIMITS);
    const names = c.entries.map(e => showName(e.name));
    expect(names).to.include.members(['', 'WordDocument', '1Table', '\\x01CompObj', 'Data']);
    const data = must(c.child(c.root, 'data'), 'Data');
    const bytes = await c.bytes(data);
    expect(bytes.length).to.equal(5000);
    expect([...bytes.subarray(250, 260)]).to.deep.equal([250, 251, 252, 253, 254, 255, 0, 1, 2, 3]);
    expect([...(await c.source(data).read(4990, 100))]).to.deep.equal([...bytes.subarray(4990)]);
    const comp = must(c.child(c.root, '\x01COMPOBJ'), 'CompObj');
    expect((await c.bytes(comp)).length).to.equal(comp.size);
    expect(entryPath(comp)).to.equal('\\x01CompObj');
    expect(entryPath(c.root)).to.equal('');
    expect(await c.source(comp).size()).to.equal(comp.size);
    expect((await c.source(comp).read(comp.size, 10)).length).to.equal(0);
    expect(guid(Uint8Array.from([0x02, 0xce, 0x02, 0x00, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46]))).to.equal('0002CE02-0000-0000-C000-000000000046');
  });

  it('refuses a version 4 file shorter than its 4,096-byte header sector', async () => {
    expect(await refusal(cfb('version-4.doc').subarray(0, 3000))).to.equal('CONTAINER/BAD_HEADER');
  });

  it('refuses a file that is not a compound file when opened as one', async () => {
    const e = await CompoundFile.open(bufferSource(Buffer.alloc(600)), LIMITS).catch((x: unknown) => x);
    expect(e).to.have.property('detail', 'BAD_HEADER');
  });
});
