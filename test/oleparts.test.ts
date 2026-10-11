import { expect } from 'chai';
import { autoName, parseBoundSheet, parseLbl, xlString } from '../src/biff';
import { ByteReader, bounded, parseCompObj, parseOle1, parseOleStream, parsePackage, readMoniker, replaceSized, Truncated } from '../src/oleobj';
import { hyperlink, virtualPath } from '../src/xls';
import { must } from './helpers/util';

const u32 = (n: number) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n >>> 0);
  return b;
};
const lp = (s: string) => Buffer.concat([u32(s.length + 1), Buffer.from(`${s}\0`, 'latin1')]);
const guid = (s: string) => {
  const h = Buffer.from(s.replace(/-/g, ''), 'hex');
  return Buffer.concat([Buffer.from(h.subarray(0, 4)).reverse(), Buffer.from(h.subarray(4, 6)).reverse(), Buffer.from(h.subarray(6, 8)).reverse(), h.subarray(8)]);
};
const URL = guid('79EAC9E0-BAF9-11CE-8C82-00AA004BA90B');
const FILE = guid('00000303-0000-0000-C000-000000000046');
const COMPOSITE = guid('00000309-0000-0000-C000-000000000046');
const ITEM = guid('00000304-0000-0000-C000-000000000046');
const urlMoniker = (url: string) => {
  const s = Buffer.from(`${url}\0`, 'utf16le');
  return Buffer.concat([URL, u32(s.length), s]);
};

describe('OLE structures', () => {
  it('reads strings bounded by their data, and says when a structure runs past it', () => {
    const r = new ByteReader(Buffer.from('ab\0cd'));
    expect(r.ansiz()).to.equal('ab');
    expect(() => r.ansiz()).to.throw(Truncated);
    expect(() => new ByteReader(Buffer.from('abc\0')).ansiz(2)).to.throw(Truncated);
    expect(new ByteReader(u32(0)).lengthPrefixedAnsi()).to.equal('');
    expect(new ByteReader(Buffer.concat([u32(3), Buffer.from('abc')])).lengthPrefixedAnsi()).to.equal('abc');
    expect(() => new ByteReader(u32(70000)).lengthPrefixedAnsi()).to.throw(Truncated);
    expect(new ByteReader(Buffer.concat([u32(3), Buffer.from('ab\0\0\0\0', 'utf16le')])).lengthPrefixedUnicode()).to.equal('ab');
    expect(() => new ByteReader(u32(40000)).lengthPrefixedUnicode()).to.throw(Truncated);
    expect(() => new ByteReader(Buffer.alloc(2)).skip(-1)).to.throw(Truncated);
    expect(
      bounded(
        () => new ByteReader(Buffer.alloc(1)).u32(),
        () => -1,
      ),
    ).to.equal(-1);
    expect(() =>
      bounded(
        () => {
          throw new RangeError('not a structure');
        },
        () => 0,
      ),
    ).to.throw(RangeError);
  });

  it('reads the user type and ProgID of a CompObj stream, whatever clipboard format it gives', () => {
    const head = Buffer.alloc(28);
    expect(parseCompObj(Buffer.concat([head, lp('Inert'), u32(0xffffffff), u32(3), lp('Inert.1')]))).to.deep.equal({ userType: 'Inert', progId: 'Inert.1' });
    expect(parseCompObj(Buffer.concat([head, lp('Inert'), u32(4), Buffer.from('fmt\0'), lp('Inert.2')]))).to.deep.equal({ userType: 'Inert', progId: 'Inert.2' });
    expect(parseCompObj(Buffer.concat([head, lp('Inert')]))).to.deep.equal({ userType: 'Inert' });
  });

  it('reads URL, file and composite monikers, and gives nothing for others', () => {
    expect(readMoniker(new ByteReader(urlMoniker('https://example.invalid/a')))).to.deep.equal({ target: 'https://example.invalid/a', url: true });
    const path = Buffer.from('C:\\inert\\wide.txt', 'utf16le');
    const file = Buffer.concat([
      FILE,
      Buffer.from([2, 0]),
      lp('short.txt'),
      Buffer.from([0xff, 0xff, 0xad, 0xde]),
      Buffer.alloc(16),
      u32(path.length + 6),
      u32(path.length),
      Buffer.from([3, 0]),
      path,
    ]);
    expect(readMoniker(new ByteReader(file))).to.deep.equal({ target: '..\\..\\C:\\inert\\wide.txt', url: false });
    const item = Buffer.concat([ITEM, Buffer.alloc(8)]);
    expect(readMoniker(new ByteReader(Buffer.concat([COMPOSITE, u32(2), urlMoniker('https://example.invalid/b'), item])))).to.deep.equal({ target: 'https://example.invalid/b', url: true });
    expect(readMoniker(new ByteReader(Buffer.concat([COMPOSITE, u32(1), item])))).to.deep.equal({ target: '', url: false });
  });

  it('reads an OLE stream: embedded, linked by a relative moniker only, or cut short', () => {
    expect(parseOleStream(Buffer.concat([u32(0x02000001), u32(0), u32(0), u32(0), u32(0)]))).to.deep.equal({ linked: false, target: '', url: false });
    const moniker = urlMoniker('https://example.invalid/rel');
    const relative = Buffer.concat([u32(0x02000001), u32(1), u32(1), u32(0), u32(0), u32(moniker.length), moniker, u32(0)]);
    expect(parseOleStream(relative)).to.deep.equal({ linked: true, target: 'https://example.invalid/rel', url: true });
    expect(parseOleStream(Buffer.concat([u32(0x02000001), u32(1), u32(1)]))).to.deep.equal({ linked: true, target: '', url: false });
  });

  it('reads Packager data with its file, as a link, or not at all', () => {
    const head = Buffer.concat([Buffer.from([2, 0]), Buffer.from('a.txt\0C:\\a.txt\0'), Buffer.alloc(8), Buffer.from('C:\\t\\a.txt\0')]);
    const whole = Buffer.concat([head, u32(5), Buffer.from('inert'), Buffer.from('tail')]);
    const pkg = must(parsePackage(whole), 'package');
    expect(pkg).to.deep.equal({ fileName: 'a.txt', sourcePath: 'C:\\a.txt', tempPath: 'C:\\t\\a.txt', embedded: true, size: 5, dataAt: 40 });
    // A file put in its place keeps what came before and after it, with the size set to match.
    expect(Buffer.from(replaceSized(whole, pkg.dataAt, pkg.size, Buffer.from('xy'))).equals(Buffer.concat([head, u32(2), Buffer.from('xytail')]))).to.equal(true);
    expect(parsePackage(head)?.embedded).to.equal(false);
    expect(parsePackage(Buffer.from([2, 0, 0x61]))).to.equal(undefined);
  });

  it('reads OLE 1.0 objects: embedded, linked, static, and cut short', () => {
    const embedded = Buffer.concat([u32(0x0501), u32(2), lp('Package'), lp(''), lp(''), u32(3), Buffer.from('abc')]);
    expect(parseOle1(embedded)).to.deep.include({ formatId: 2, className: 'Package' });
    expect(parseOle1(Buffer.concat([u32(0x0501), u32(1), lp('Excel.Sheet.8'), lp('C:\\book.xls'), lp('A1')]))).to.deep.equal({
      formatId: 1,
      className: 'Excel.Sheet.8',
      topicName: 'C:\\book.xls',
      itemName: 'A1',
    });
    expect(parseOle1(Buffer.concat([u32(0x0501), u32(5)]))).to.deep.equal({ formatId: 5, className: '', topicName: '', itemName: '' });
    expect(parseOle1(Buffer.concat([u32(0x0501), u32(2), lp('Package')]))).to.equal(undefined);
  });
});

describe('BIFF records', () => {
  it('reads 8-bit and 16-bit strings, and an empty one at the end of a record', () => {
    expect(xlString(Buffer.from([0, 0x61, 0x62]), 0, 2)).to.deep.equal({ text: 'ab', next: 3 });
    expect(xlString(Buffer.from([1, 0x61, 0, 0x62, 0]), 0, 2)).to.deep.equal({ text: 'ab', next: 5 });
    expect(xlString(Buffer.from([1]), 1, 2)).to.deep.equal({ text: '', next: 1 });
  });

  it('reads BoundSheet and Lbl records in BIFF5 and BIFF8, short or not, and names built-in names', () => {
    expect(parseBoundSheet(Buffer.from([16, 0, 0, 0, 1, 1, 3, 0x4d, 0x31, 0x31]), false)).to.deep.equal({ position: 16, state: 1, type: 1, name: 'M11' });
    expect(parseBoundSheet(Buffer.from([16, 0, 0, 0, 0, 0]), true)).to.deep.equal({ position: 16, state: 0, type: 0, name: '' });
    const biff5 = Buffer.concat([Buffer.from([0, 0, 0, 9]), Buffer.alloc(10), Buffer.from('Auto_Open')]);
    expect(parseLbl(biff5, false)).to.deep.equal({ name: 'Auto_Open', builtin: false, hidden: false });
    expect(parseLbl(Buffer.concat([Buffer.from([0x21, 0, 0, 1]), Buffer.alloc(10), Buffer.from([0, 0x0b])]), true)).to.deep.equal({ name: 'Auto_Deactivate', builtin: true, hidden: true });
    expect(parseLbl(Buffer.concat([Buffer.from([0x20, 0, 0, 1]), Buffer.alloc(10), Buffer.from([0, 0x06])]), true).name).to.equal('built-in name 6');
    expect(parseLbl(Buffer.from([0x20, 0]), true)).to.deep.equal({ name: '', builtin: true, hidden: false });
    expect(autoName('Auto_Activate')).to.deep.equal({ event: 'Auto_Activate', onOpen: true });
    expect(autoName('_xlnm.Auto_Close2')).to.deep.equal({ event: 'Auto_Close', onOpen: false });
    expect(autoName('Print_Area')).to.equal(undefined);
  });

  it('writes out the volumes, folders and steps of an encoded VirtualPath', () => {
    expect(virtualPath('inert-book.xls')).to.equal('inert-book.xls');
    expect(virtualPath('\x01\x01@host\x03share\x03a.xls')).to.equal('\\\\host\\share\\a.xls');
    expect(virtualPath('\x01\x01D\x03a.xls')).to.equal('D:\\a.xls');
    expect(virtualPath('\x01\x02dir\x03\x04up.xls')).to.equal('\\dir\\..\\up.xls');
    expect(virtualPath('\x01\x06a.xls\x07\x08')).to.equal('XLSTART\\a.xlsXLSTART\\XLSTART\\');
    expect(virtualPath('\x01\x01')).to.equal(':');
  });

  it("reads an HLink record's target from its moniker, a string, or nothing", () => {
    const head = Buffer.concat([Buffer.alloc(8), Buffer.alloc(16), u32(2)]);
    const text = (s: string) => {
      const b = Buffer.from(`${s}\0`, 'utf16le');
      return Buffer.concat([u32(b.length / 2), b]);
    };
    expect(hyperlink(Buffer.concat([head, u32(0x01 | 0x10 | 0x80 | 0x100), text('Shown'), text('_blank'), text('https://example.invalid/s')]))).to.equal('https://example.invalid/s');
    expect(hyperlink(Buffer.concat([head, u32(0x08), text('Sheet2!A1')]))).to.equal('');
    expect(hyperlink(Buffer.concat([head, u32(0x01)]))).to.equal('');
  });
});
