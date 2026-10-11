import * as zlib from 'node:zlib';
import { expect } from 'chai';
import { inspectOffice } from '../src';
import { StopError } from '../src/errors';
import { bufferSource, Reader } from '../src/io';
import { ZipArchive } from '../src/zip';
import { contentTypesXml, descriptorOf, extraRecord, failure, fixture, kinds, le, must, of, type RawEntry, rawZip, relsXml, T, WORD_MAIN, wordXml, zipOf } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 512 * 1024 * 1024, totalDecompressedBytes: 2 * 1024 * 1024 * 1024 };

/** The three parts of a minimal Word document, as raw entries. */
const docx = (): RawEntry[] => [
  { name: '[Content_Types].xml', data: contentTypesXml({ 'word/document.xml': WORD_MAIN }) },
  { name: '_rels/.rels', data: relsXml(['rId1', `${T}officeDocument`, 'word/document.xml']) },
  { name: 'word/document.xml', data: wordXml('<w:p/>') },
];
/** The minimal document with changes to word/document.xml's entry. */
const docxWith = (changes: Partial<RawEntry>): RawEntry[] => {
  const es = docx();
  es[2] = { ...es[2], ...changes };
  return es;
};
const DOC = wordXml('<w:p/>');
const DOC_DEFLATED = zlib.deflateRawSync(Buffer.from(DOC));

async function contents(z: ZipArchive, name: string): Promise<Buffer> {
  const chunks: Uint8Array[] = [];
  for await (const c of z.read(must(z.get(name), name))) chunks.push(c);
  return Buffer.concat(chunks);
}

/** The finding a refused archive stops the run with, as category/detail. */
async function refusal(bytes: Uint8Array): Promise<string> {
  const r = await inspectOffice(bytes);
  expect([r.status, r.score, r.risk]).to.deep.equal(['rejected', null, 'UNKNOWN']);
  return kinds(r).slice(-1)[0];
}

/** The detail a StopError carries, or what else the promise did. */
async function stopOf(p: Promise<unknown>): Promise<string> {
  const e = await failure(p);
  return e instanceof StopError ? `${e.category}/${e.detail}` : e === undefined ? 'no error' : String(e);
}

/** Reads every entry, which inflates it and checks it. */
async function readAll(bytes: Uint8Array): Promise<void> {
  const z = await ZipArchive.open(bufferSource(bytes), LIMITS);
  for (const e of z.entries) for await (const _ of z.read(e));
}

/** The archive with a 16-bit field at `at` in the header that starts with `signature` and names `name` set to `value`. */
function patch16(zip: Buffer, signature: number, name: string, at: number, value: number): Buffer {
  const nameAt = signature === 0x04034b50 ? 30 : 46;
  for (let p = zip.indexOf(le.u32(signature)); p >= 0; p = zip.indexOf(le.u32(signature), p + 1)) {
    if (zip.subarray(p + nameAt, p + nameAt + name.length).toString('latin1') !== name) continue;
    const out = Buffer.from(zip);
    out.writeUInt16LE(value, p + at);
    return out;
  }
  throw new Error(`no header for ${name}`);
}

/** The archive with its end record's directory size and offset changed by `size` and `offset` bytes. */
function shiftDirectory(zip: Buffer, size: number, offset: number): Buffer {
  const out = Buffer.from(zip);
  const end = out.lastIndexOf(le.u32(0x06054b50));
  out.writeUInt32LE(out.readUInt32LE(end + 12) + size, end + 12);
  out.writeUInt32LE(out.readUInt32LE(end + 16) + offset, end + 16);
  return out;
}

/** An archive whose end record is preceded by a ZIP64 end record and locator, with changes to each. */
function zip64End(
  o: { end64?: Partial<{ size: number; count: number; cdSize: number; cdOffset: number; signature: number }>; locator?: Partial<{ disk: number; at: number; disks: number }>; end?: object } = {},
): Buffer {
  const base = rawZip(docx(), { end: o.end });
  const endAt = base.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const cdSize = base.readUInt32LE(endAt + 12);
  const cdOffset = base.readUInt32LE(endAt + 16);
  const e = { size: 44, count: 3, cdSize, cdOffset, signature: 0x06064b50, ...o.end64 };
  const end64 = Buffer.concat([le.u32(e.signature), le.u64(e.size), le.u16(45), le.u16(45), le.u32(0), le.u32(0), le.u64(e.count), le.u64(e.count), le.u64(e.cdSize), le.u64(e.cdOffset)]);
  const l = { disk: 0, at: endAt, disks: 1, ...o.locator };
  const locator = Buffer.concat([le.u32(0x07064b50), le.u32(l.disk), le.u64(l.at), le.u32(l.disks)]);
  return Buffer.concat([base.subarray(0, endAt), end64, locator, base.subarray(endAt)]);
}

describe('ZIP reader', () => {
  const refusedFixtures: Array<[string, string]> = [
    ['multi-disk.docx', 'CONTAINER/MULTI_DISK'],
    ['count-mismatch.docx', 'CONTAINER/DIRECTORY_MISMATCH'],
    ['prepended.docx', 'CONTAINER/DIRECTORY_MISMATCH'],
    ['overlap.docx', 'CONTAINER/OVERLAP'],
    ['out-of-range.docx', 'CONTAINER/OUT_OF_RANGE'],
    ['local-name.docx', 'CONTAINER/HEADER_MISMATCH'],
    ['local-size.docx', 'CONTAINER/HEADER_MISMATCH'],
    ['local-method.docx', 'CONTAINER/HEADER_MISMATCH'],
    ['descriptor-mismatch.docx', 'CONTAINER/HEADER_MISMATCH'],
    ['traversal.docx', 'CONTAINER/UNSAFE_NAME'],
    ['drive.docx', 'CONTAINER/UNSAFE_NAME'],
    ['directory-content.docx', 'CONTAINER/UNSAFE_NAME'],
    ['name-encoding.docx', 'CONTAINER/NAME_ENCODING'],
    ['bzip2.docx', 'CONTAINER/UNSUPPORTED_METHOD'],
    ['zip64-missing.docx', 'CONTAINER/ZIP64_MISMATCH'],
    ['zip64-extra-length.docx', 'CONTAINER/ZIP64_MISMATCH'],
    ['zip64-count.docx', 'CONTAINER/ZIP64_MISMATCH'],
    ['bad-extra.docx', 'CONTAINER/BAD_EXTRA_FIELD'],
    ['repeated-extra.docx', 'CONTAINER/BAD_EXTRA_FIELD'],
    ['bad-size.docx', 'CONTAINER/DATA_ERROR'],
    ['deflate-leftover.docx', 'CONTAINER/DATA_ERROR'],
    ['stored-sizes.docx', 'CONTAINER/DATA_ERROR'],
  ];
  for (const [name, expected] of refusedFixtures) {
    it(`refuses ${name} with ${expected} and no score`, async () => {
      expect(await refusal(fixture('zip', name))).to.equal(expected);
    });
  }

  const ones = 2n ** 64n - 1n;
  const zip64Sizes = (usize: number, csize: number) => extraRecord(1, Buffer.concat([le.u64(usize), le.u64(csize)]));
  const refusedBuilt: Array<[string, () => Buffer, string]> = [
    ['a file shorter than an end record', () => Buffer.alloc(10), 'CONTAINER/NOT_ZIP'],
    ['an end record whose comment runs past the end of the file', () => rawZip(docx(), { end: { commentLength: 50 } }), 'CONTAINER/NOT_ZIP'],
    ['a disk number other than 0', () => rawZip(docx(), { end: { disk: 1 } }), 'CONTAINER/MULTI_DISK'],
    ['entry counts that differ between this disk and the archive', () => rawZip(docx(), { end: { countDisk: 2 } }), 'CONTAINER/DIRECTORY_MISMATCH'],
    ['a directory offset past the end of the file', () => rawZip(docx(), { cdOffsetDelta: 100_000 }), 'CONTAINER/OUT_OF_RANGE'],
    ['a directory size larger than the file', () => rawZip(docx(), { end: { cdSize: 100_000 } }), 'CONTAINER/OUT_OF_RANGE'],
    ['a directory that ends before the end record', () => rawZip(docx(), { end: { cdSize: 46 } }), 'CONTAINER/DIRECTORY_MISMATCH'],
    ['a central header without its signature', () => shiftDirectory(rawZip(docx()), 4, -4), 'CONTAINER/DIRECTORY_MISMATCH'],
    ['a central header whose extra field runs past the directory', () => patch16(rawZip(docx()), 0x02014b50, 'word/document.xml', 30, 1000), 'CONTAINER/DIRECTORY_MISMATCH'],
    ['an entry name over 1,024 bytes', () => rawZip([...docx(), { name: `word/${'a'.repeat(1100)}.xml`, data: 'x' }]), 'CONTAINER/UNSAFE_NAME'],
    ['an empty entry name', () => rawZip([...docx(), { name: '', data: 'x' }]), 'CONTAINER/UNSAFE_NAME'],
    ['an entry name with a control character', () => rawZip([...docx(), { name: 'word/a\u0001.xml', data: 'x' }]), 'CONTAINER/UNSAFE_NAME'],
    ['an empty path segment', () => rawZip([...docx(), { name: 'word//a.xml', data: 'x' }]), 'CONTAINER/UNSAFE_NAME'],
    ['a name flagged UTF-8 that is not valid UTF-8', () => rawZip([...docx(), { name: Buffer.from([0x61, 0xff]), data: 'x', flags: 0x0800 }]), 'CONTAINER/NAME_ENCODING'],
    [
      'an Info-ZIP Unicode path field of an unknown version',
      () => rawZip(docxWith({ centralExtra: extraRecord(0x7075, Buffer.concat([Buffer.from([2]), le.u32(0), Buffer.from('word/document.xml')])) })),
      'CONTAINER/NAME_ENCODING',
    ],
    ['an Info-ZIP Unicode path field that is not UTF-8', () => rawZip(docxWith({ centralExtra: extraRecord(0x7075, Buffer.from([1, 0, 0, 0, 0, 0xff])) })), 'CONTAINER/NAME_ENCODING'],
    ['an extra field block that ends inside a record header', () => rawZip(docxWith({ centralExtra: Buffer.from([1, 2]) })), 'CONTAINER/BAD_EXTRA_FIELD'],
    ['strong encryption', () => rawZip(docxWith({ flags: 0x0040 })), 'CONTAINER/ENCRYPTED_ENTRY'],
    ['an encrypted central directory', () => rawZip(docxWith({ flags: 0x2000 })), 'CONTAINER/ENCRYPTED_ENTRY'],
    ['a ZIP64 field whose disk number is not 0', () => rawZip(docxWith({ central: { disk: 0xffff }, centralExtra: extraRecord(1, le.u32(1)) })), 'CONTAINER/MULTI_DISK'],
    ['a ZIP64 offset field beyond a safe integer', () => rawZip(docxWith({ central: { offset: 0xffffffff }, centralExtra: extraRecord(1, le.u64(2n ** 60n)) })), 'CONTAINER/OUT_OF_RANGE'],
    ['a declared size past the limit', () => rawZip(docxWith({ central: { usize: 0xffffffff, csize: 0xffffffff }, centralExtra: zip64Sizes(2 ** 40, 2 ** 40) })), 'LIMIT/DECOMPRESSED_SIZE'],
    ['a directory that holds more than the end record counts', () => rawZip(docx(), { end: { count: 2, countDisk: 2 } }), 'CONTAINER/DIRECTORY_MISMATCH'],
    ['two entries whose names match once percent-decoded', () => rawZip([...docx(), { name: 'word/a%20b.xml', data: 'x' }, { name: 'word/a b.xml', data: 'y' }]), 'CONTAINER/DUPLICATE_NAME'],
    [
      'two entries whose names match in Unicode NFC',
      () => rawZip([...docx(), { name: `word/caf${String.fromCodePoint(0xe9)}.xml`, data: 'x', flags: 0x800 }, { name: `word/cafe${String.fromCodePoint(0x301)}.xml`, data: 'y', flags: 0x800 }]),
      'CONTAINER/DUPLICATE_NAME',
    ],
    ['a file and a directory of the same name', () => rawZip([...docx(), { name: 'word/media', data: 'x' }, { name: 'word/media/', data: '', method: 0 }]), 'CONTAINER/DUPLICATE_NAME'],
    ['a local header past the start of the directory', () => rawZip(docxWith({ central: { offset: 99999 } })), 'CONTAINER/OUT_OF_RANGE'],
    ['a local header without its signature', () => patch16(rawZip(docx()), 0x04034b50, 'word/document.xml', 0, 0), 'CONTAINER/HEADER_MISMATCH'],
    ['a local extra field that runs into the directory', () => patch16(rawZip(docx()), 0x04034b50, 'word/document.xml', 28, 0xffff), 'CONTAINER/OUT_OF_RANGE'],
    ['a local header with other flags', () => rawZip(docxWith({ local: { flags: 0x0800 } })), 'CONTAINER/HEADER_MISMATCH'],
    ['a local ZIP64 field of an odd length', () => rawZip(docxWith({ localExtra: extraRecord(1, Buffer.alloc(8)) })), 'CONTAINER/ZIP64_MISMATCH'],
    ['a local ZIP64 field that disagrees with the local size', () => rawZip(docxWith({ localExtra: zip64Sizes(1, DOC_DEFLATED.length) })), 'CONTAINER/ZIP64_MISMATCH'],
    ['a local ZIP64 field that disagrees with the local compressed size', () => rawZip(docxWith({ localExtra: zip64Sizes(DOC.length, 1) })), 'CONTAINER/ZIP64_MISMATCH'],
    ['a local size at its maximum with no ZIP64 field', () => rawZip(docxWith({ local: { usize: 0xffffffff } })), 'CONTAINER/ZIP64_MISMATCH'],
    [
      'a local CRC that disagrees with a data descriptor entry',
      () => rawZip(docxWith({ flags: 8, local: { crc: 1, csize: 0, usize: 0 }, after: descriptorOf(DOC, DOC_DEFLATED) })),
      'CONTAINER/HEADER_MISMATCH',
    ],
    [
      'a local size that disagrees with a data descriptor entry',
      () => rawZip(docxWith({ flags: 8, local: { crc: 0, csize: 0, usize: 1 }, after: descriptorOf(DOC, DOC_DEFLATED) })),
      'CONTAINER/HEADER_MISMATCH',
    ],
    [
      'a local compressed size that disagrees with a data descriptor entry',
      () => rawZip(docxWith({ flags: 8, local: { crc: 0, csize: 1, usize: 0 }, after: descriptorOf(DOC, DOC_DEFLATED) })),
      'CONTAINER/HEADER_MISMATCH',
    ],
    [
      'a data descriptor with other sizes in either width',
      () => rawZip(docxWith({ flags: 8, local: { crc: 0, csize: 0, usize: 0 }, after: descriptorOf(DOC, DOC_DEFLATED).fill(0, 8) })),
      'CONTAINER/HEADER_MISMATCH',
    ],
    ['entry data that runs into the directory', () => rawZip(docxWith({ central: { csize: DOC_DEFLATED.length + 10 }, local: { csize: DOC_DEFLATED.length + 10 } })), 'CONTAINER/OVERLAP'],
    ['an end record at its maximum with no ZIP64 record', () => rawZip(docx(), { end: { cdOffset: 0xffffffff } }), 'CONTAINER/ZIP64_MISMATCH'],
    ['a ZIP64 locator on another disk', () => zip64End({ locator: { disks: 2 } }), 'CONTAINER/MULTI_DISK'],
    ['a ZIP64 locator that points past itself', () => zip64End({ locator: { at: 1_000_000 } }), 'CONTAINER/ZIP64_MISMATCH'],
    ['a ZIP64 locator that points at something else', () => zip64End({ end64: { signature: 0x12345678 } }), 'CONTAINER/ZIP64_MISMATCH'],
    ['a ZIP64 end record that does not end at its locator', () => zip64End({ end64: { size: 48 } }), 'CONTAINER/ZIP64_MISMATCH'],
    ['a ZIP64 end record whose directory size disagrees', () => zip64End({ end64: { cdSize: 1 } }), 'CONTAINER/ZIP64_MISMATCH'],
    ['a deflate stream that is corrupt', () => rawZip(docxWith({ payload: Buffer.from([0xff, 0xff, 0xff, 0xff]) })), 'CONTAINER/DATA_ERROR'],
    ['a deflate stream cut short', () => rawZip(docxWith({ payload: DOC_DEFLATED.subarray(0, 20) })), 'CONTAINER/DATA_ERROR'],
  ];
  for (const [what, make, expected] of refusedBuilt) {
    it(`refuses ${what} with ${expected}`, async () => {
      expect(await refusal(make())).to.equal(expected);
    });
  }

  it('names the entry a refusal is about, in the location and the description', async () => {
    const r = await inspectOffice(fixture('cases', 'wrong-crc.docx'));
    const f = must(of(r, 'CONTAINER/DATA_ERROR')[0], 'finding');
    expect(f.location).to.equal('word/document.xml');
    expect(f.description).to.equal("An entry's data does not match its declared size or checksum: word/document.xml");
  });

  const accepted = ['minimal.docx', 'descriptor.docx', 'descriptor-nosig.docx', 'streamed.docx', 'zip64.docx'];
  for (const name of accepted) {
    it(`reads ${name} as a clean docx`, async () => {
      const r = await inspectOffice(fixture('zip', name));
      expect({ status: r.status, format: r.format, findings: kinds(r) }).to.deep.equal({ status: 'clean', format: 'docx', findings: [] });
    });
  }

  it('reads a data descriptor after an empty deflated entry, as LibreOffice writes Configurations2/accelerator/current.xml', async () => {
    // The eight bytes after a narrow descriptor are the next local header's signature: read as a 64-bit size they
    // overflow a number, which once refused the archive as OUT_OF_RANGE.
    // LibreOffice sets the UTF-8 flag on every entry, which puts a large value in those eight bytes.
    const empty = zlib.deflateRawSync(Buffer.alloc(0));
    const entries = docxWith({ flags: 0x0800 });
    entries.splice(2, 0, { name: 'word/empty.xml', data: '', flags: 0x0808, local: { crc: 0, csize: 0, usize: 0 }, payload: empty, after: descriptorOf('', empty) });
    const z = await ZipArchive.open(bufferSource(rawZip(entries)), LIMITS);
    expect(z.notes).to.deep.equal([]);
    expect((await contents(z, 'word/empty.xml')).length).to.equal(0);
  });

  it('reads streamed ZIP64 entries whose local sizes are ones, with the sizes in 24-byte data descriptors', async () => {
    const data = Buffer.from(DOC);
    const entries = docxWith({
      flags: 8,
      local: { crc: 0, csize: 0xffffffff, usize: 0xffffffff },
      localExtra: extraRecord(1, Buffer.concat([le.u64(ones), le.u64(ones)])),
      central: { csize: 0xffffffff, usize: 0xffffffff },
      centralExtra: zip64Sizes(data.length, DOC_DEFLATED.length),
      after: descriptorOf(data, DOC_DEFLATED, true),
    });
    const z = await ZipArchive.open(bufferSource(rawZip(entries)), LIMITS);
    expect(z.notes).to.deep.equal([]);
    expect((await contents(z, 'word/document.xml')).equals(data)).to.equal(true);
    // Without the data descriptor flag the ones are a size, and too large a one.
    const plain = docxWith({ local: { csize: 0xffffffff, usize: 0xffffffff }, localExtra: extraRecord(1, Buffer.concat([le.u64(ones), le.u64(ones)])) });
    expect(await refusal(rawZip(plain))).to.equal('CONTAINER/OUT_OF_RANGE');
  });

  it('accepts a data descriptor of the other width when its values agree, and counts it as part of the entry', async () => {
    const narrowWithZip64 = docxWith({
      flags: 8,
      local: { crc: 0, csize: 0, usize: 0 },
      localExtra: zip64Sizes(DOC.length, DOC_DEFLATED.length),
      after: descriptorOf(DOC, DOC_DEFLATED),
    });
    const wideWithout = docxWith({ flags: 8, local: { crc: 0, csize: 0, usize: 0 }, after: descriptorOf(DOC, DOC_DEFLATED, true) });
    for (const entries of [narrowWithZip64, wideWithout]) expect((await ZipArchive.open(bufferSource(rawZip(entries)), LIMITS)).notes).to.deep.equal([]);
  });

  it('accepts a ZIP64 end record that agrees with the end record, and one the end record defers to', async () => {
    expect(await stopOf(ZipArchive.open(bufferSource(zip64End()), LIMITS))).to.equal('no error');
    expect(await stopOf(ZipArchive.open(bufferSource(zip64End({ end: { count: 0xffff, countDisk: 0xffff } })), LIMITS))).to.equal('no error');
  });

  it('accepts an Info-ZIP Unicode path field that names the entry itself', async () => {
    const field = extraRecord(0x7075, Buffer.concat([Buffer.from([1]), le.u32(0), Buffer.from('word/document.xml')]));
    expect(await stopOf(readAll(rawZip(docxWith({ centralExtra: field, localExtra: field }))))).to.equal('no error');
  });

  const noted: Array<[string, string, number]> = [
    ['leading-bytes.docx', 'CONTAINER/LEADING_BYTES', 64],
    ['hidden-data.docx', 'CONTAINER/HIDDEN_DATA', 46],
    ['trailing-bytes.docx', 'CONTAINER/TRAILING_BYTES', 40],
  ];
  for (const [name, k, bytes] of noted) {
    it(`reports ${k} in ${name}, a strip finding, and reads on`, async () => {
      const r = await inspectOffice(fixture('zip', name));
      expect({ status: r.status, format: r.format, kinds: kinds(r), data: r.findings[0].data }).to.deep.equal({ status: 'strippable', format: 'docx', kinds: [k], data: { bytes } });
    });
  }

  it('reports bytes before the directory of an archive with no entries as leading bytes', async () => {
    const z = await ZipArchive.open(bufferSource(rawZip([], { prefix: Buffer.from('stub') })), LIMITS);
    expect(z.notes).to.deep.equal([{ detail: 'LEADING_BYTES', bytes: 4 }]);
  });

  it('reads the same content Python wrote, with and without data descriptors and ZIP64', async () => {
    const expected = await contents(await ZipArchive.open(bufferSource(fixture('zip', 'minimal.docx')), LIMITS), 'word/document.xml');
    for (const name of ['descriptor.docx', 'descriptor-nosig.docx', 'streamed.docx', 'zip64.docx']) {
      const z = await ZipArchive.open(bufferSource(fixture('zip', name)), LIMITS);
      expect(z.entries.map(e => e.name)).to.deep.equal(['[Content_Types].xml', '_rels/.rels', 'word/document.xml']);
      expect((await contents(z, 'word/document.xml')).equals(expected), name).to.equal(true);
    }
  });

  it('looks entries up ignoring case and percent-encoding', async () => {
    const z = await ZipArchive.open(bufferSource(rawZip([...docx(), { name: 'word/a%20b.xml', data: 'x' }])), LIMITS);
    expect(must(z.get('[content_types].XML'), 'entry').name).to.equal('[Content_Types].xml');
    expect(must(z.get('WORD/A B.XML'), 'entry').name).to.equal('word/a%20b.xml');
    // A name that is not valid percent-encoding is compared as written.
    const odd = await ZipArchive.open(bufferSource(rawZip([...docx(), { name: 'word/100%.xml', data: 'x' }])), LIMITS);
    expect(must(odd.get('word/100%.xml'), 'entry').name).to.equal('word/100%.xml');
  });

  it('refuses entry counts and declared sizes past the limits before reading any entry', async () => {
    const bytes = fixture('zip', 'minimal.docx');
    let reads = 0;
    const counting = {
      size: async () => bytes.length,
      read: async (o: number, n: number) => {
        reads++;
        return bytes.subarray(o, o + n);
      },
    };
    const stop = (limits: Partial<typeof LIMITS>) => stopOf(ZipArchive.open(counting, { ...LIMITS, ...limits }));
    expect(await stop({ entries: 2 })).to.equal('LIMIT/ENTRY_COUNT');
    expect(await stop({ decompressedBytes: 100 })).to.equal('LIMIT/DECOMPRESSED_SIZE');
    expect(await stop({ totalDecompressedBytes: 600 })).to.equal('LIMIT/TOTAL_DECOMPRESSED_SIZE');
    expect(await stop({ entries: 3, decompressedBytes: 1000, totalDecompressedBytes: 2000 })).to.equal('no error');
    // Each open reads the one block that holds the whole small archive, and no entry data.
    expect(reads).to.equal(4);
  });

  it('stops inflating once an entry passes its declared size', async () => {
    const big = Buffer.alloc(4 * 1024 * 1024);
    const zip = rawZip([{ name: 'a.bin', data: big, local: { usize: 1000 }, central: { usize: 1000 } }]);
    const z = await ZipArchive.open(bufferSource(zip), LIMITS);
    let seen = 0;
    const e = await failure(
      (async () => {
        for await (const c of z.read(must(z.get('a.bin'), 'entry'))) seen += c.length;
      })(),
    );
    expect((e as StopError).detail).to.equal('DATA_ERROR');
    expect(seen).to.be.lessThan(big.length);
  });

  it('refuses deflate data followed by more than a chunk of leftover bytes', async () => {
    const data = Buffer.from('hello hello hello');
    const zip = rawZip([{ name: 'x.txt', data, payload: Buffer.concat([zlib.deflateRawSync(data), Buffer.alloc(600 * 1024, 7)]) }]);
    expect(await stopOf(readAll(zip))).to.equal('CONTAINER/DATA_ERROR');
  });

  it('stops a read that passes the time limit', async () => {
    const zip = await zipOf([['a.bin', Buffer.alloc(1024)]]);
    const z = await ZipArchive.open(bufferSource(zip), LIMITS, Date.now() + 20);
    await new Promise(resolve => setTimeout(resolve, 30));
    const e = await failure(
      (async () => {
        for await (const _ of z.read(z.entries[0]));
      })(),
    );
    expect((e as { code?: string }).code).to.equal('DEFUSE_LIMIT');
  });
});

describe('block reader', () => {
  it('reads ranges across blocks and keeps only the newest blocks', async () => {
    const bytes = Buffer.from(Array.from({ length: 4096 }, (_, i) => i & 0xff));
    const reads: number[] = [];
    const source = {
      size: async () => bytes.length,
      read: async (o: number, n: number) => {
        reads.push(o);
        return bytes.subarray(o, o + n);
      },
    };
    const r = new Reader(source, bytes.length, 256, 2);
    expect(Buffer.from(await r.read(250, 20)).equals(bytes.subarray(250, 270))).to.equal(true);
    expect(Buffer.from(await r.read(4090, 100)).equals(bytes.subarray(4090))).to.equal(true);
    expect((await r.read(5000, 10)).length).to.equal(0);
    // Blocks 0 and 1 were read first, then block 15 pushed block 0 out.
    await r.read(0, 1);
    expect(reads).to.deep.equal([0, 256, 3840, 0]);
  });
});
