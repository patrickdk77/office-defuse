import { expect } from 'chai';
import { XlsbLinks } from '../src/xlsb';

/** One record: the type and the size in seven-bit groups, then the body. */
function record(type: number, body: Uint8Array): Buffer {
  const head: number[] = [];
  for (const [value, most] of [
    [type, 2],
    [body.length, 4],
  ]) {
    let v = value;
    for (let i = 0; i < most; i++) {
      head.push((v & 0x7f) | (v > 0x7f ? 0x80 : 0));
      v >>>= 7;
      if (v === 0) break;
    }
  }
  return Buffer.concat([Buffer.from(head), body]);
}
const wide = (s: string) => Buffer.concat([Buffer.from(Uint32Array.of(s.length).buffer), Buffer.from(s, 'utf16le')]);
const supBook = (type: number, ...names: string[]) => record(360, Buffer.concat([Buffer.from([type, 0]), ...names.map(wide)]));

function read(bytes: Uint8Array, piece = bytes.length || 1): XlsbLinks {
  const links = new XlsbLinks();
  for (let i = 0; i < bytes.length; i += piece) links.write(bytes.subarray(i, i + piece));
  return links;
}

describe('xlsb external link records', () => {
  it('collects the application and topic of each DDE link record, however the bytes arrive', () => {
    const bytes = Buffer.concat([record(359, Buffer.alloc(0)), supBook(1, 'cmd', '/c calc'), record(361, Buffer.alloc(0)), supBook(1, 'app', 'topic')]);
    for (const piece of [1, 2, 5, bytes.length])
      expect(read(bytes, piece).dde, String(piece)).to.deep.equal([
        { application: 'cmd', topic: '/c calc' },
        { application: 'app', topic: 'topic' },
      ]);
  });

  it('skips workbook and OLE link records and every other record, including one over 64 KiB', () => {
    const big = record(0x1234, Buffer.alloc(70_000));
    expect(read(Buffer.concat([supBook(0, 'a.xlsx', 'Sheet1'), supBook(2, 'Excel.Sheet', 'x'), big])).dde).to.deep.equal([]);
    // Only the first 64 KiB of a DDE link record are kept, which is plenty for its names.
    expect(read(Buffer.concat([record(360, Buffer.concat([Buffer.from([1, 0]), wide('app'), wide('topic'), Buffer.alloc(70_000)]))]), 1000).dde).to.deep.equal([
      { application: 'app', topic: 'topic' },
    ]);
  });

  it('reads a DDE link record cut short as far as it goes', () => {
    expect(read(record(360, Buffer.from([1, 0]))).dde).to.deep.equal([{ application: '', topic: '' }]);
    expect(read(record(360, Buffer.concat([Buffer.from([1, 0]), Buffer.from(Uint32Array.of(50).buffer), Buffer.from('ab', 'utf16le')]))).dde).to.deep.equal([{ application: 'ab', topic: '' }]);
    expect(read(record(360, Buffer.from([1]))).dde).to.deep.equal([]);
  });

  it('stops at a record header longer than the format allows', () => {
    expect(read(Buffer.concat([Buffer.from([0x80, 0x80, 0x01]), supBook(1, 'a', 'b')])).dde).to.deep.equal([]);
    expect(read(Buffer.concat([Buffer.from([0x01, 0x80, 0x80, 0x80, 0x80, 0x01]), supBook(1, 'a', 'b')])).dde).to.deep.equal([]);
  });
});
