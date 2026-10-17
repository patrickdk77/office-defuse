import { expect } from 'chai';
import { disarmOffice, type ScriptPlugin } from '../src';
import { type BiffRecord, R, RecordReader } from '../src/biff';
import { CompoundFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-writer';
import { bufferSink, bufferSource } from '../src/io';
import { fixture, kinds, must, of } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const T = { index: 0x020b, dbcell: 0x00d7, extsst: 0x00ff, sst: 0x00fc, defColWidth: 0x0055, row: 0x0208, labelSst: 0x00fd, supBook: 0x01ae };

/** The records of the output's workbook stream, with their bodies. */
async function workbookRecords(bytes: Uint8Array): Promise<{ records: BiffRecord[]; stream: Uint8Array }> {
  const cfb = await CompoundFile.open(bufferSource(bytes), LIMITS);
  const entry = must(
    cfb.entries.find(e => e.name === 'Workbook'),
    'Workbook',
  );
  const reader = new RecordReader(cfb, entry, 0);
  const records: BiffRecord[] = [];
  for (let r = await reader.next(); r; r = await reader.next()) records.push({ ...r, body: Buffer.from(r.body) });
  return { records, stream: await cfb.bytes(entry) };
}

const u32 = (b: Uint8Array, offset: number) => Buffer.from(b).readUInt32LE(offset);
const at = (records: BiffRecord[], offset: number) => records.find(r => r.offset === offset);

const passAll: ScriptPlugin = { kind: 'script', name: 'test', accepts: () => true, process: async () => ({ result: 'passed' }) };

describe('BIFF8 workbook rewrite', () => {
  it('turns a macro sheet into an empty worksheet and renames the names that run it', async () => {
    const r = await disarmOffice(fixture('ole', 'xlm.xls'));
    expect([r.status, must(r.after, 'after').status]).to.deep.equal(['defused', 'clean']);
    const { records } = await workbookRecords(must(r.bytes, 'output'));
    const sheets = records.filter(x => x.type === R.boundSheet);
    const macro = must(
      sheets.find(x => Buffer.from(x.body.subarray(8)).toString('latin1') === 'Macro1'),
      'Macro1',
    );
    expect(macro.body[5]).to.equal(0);
    // The sheet's BOF says worksheet, and nothing but an empty sheet's records follow it.
    const start = records.findIndex(x => x.offset === u32(macro.body, 0));
    const substream = records.slice(start, start + 5).map(x => x.type);
    expect(substream).to.deep.equal([R.bof, T.defColWidth, 0x0200, 0x023e, R.eof]);
    expect(Buffer.from(records[start].body).readUInt16LE(2)).to.equal(0x10);
    const names = records.filter(x => x.type === R.lbl).map(x => [x.body[0] & 0x20, Buffer.from(x.body.subarray(15, 15 + x.body[3])).toString('latin1')]);
    expect(names).to.deep.equal([
      [0, 'Removed_Auto_Open_1'],
      [0, 'Removed_Auto_Close_1'],
    ]);
  });

  it('writes a workbook VelvetSweatshop opens decrypted, without its FILEPASS record', async () => {
    const r = await disarmOffice(fixture('ole', 'rc4.xls'));
    expect([r.status, ...kinds(r.before)]).to.deep.equal(['defused', 'ENCRYPTED/DEFAULT_PASSWORD', 'ENCRYPTED/RC4_40', 'MACRO/EVENT_ON_OPEN', 'MACRO/XLM_MACROSHEET']);
    const { records } = await workbookRecords(must(r.bytes, 'output'));
    expect(records.some(x => x.type === R.filePass)).to.equal(false);
    expect(must(r.after, 'after').findings).to.deep.equal([]);
  });

  it('points links at the workbook itself, and drops hyperlinks and button macros that strip', async () => {
    const r = await disarmOffice(fixture('ole', 'links.xls'));
    expect([r.status, ...kinds(must(r.after, 'after'))]).to.deep.equal(['defused', 'LINK/SAFE']);
    const { records } = await workbookRecords(must(r.bytes, 'output'));
    const paths = records
      .filter(x => x.type === T.supBook && Buffer.from(x.body).readUInt16LE(2) < 0x401)
      .map(x => Buffer.from(x.body.subarray(5, 5 + Buffer.from(x.body).readUInt16LE(2))).toString('latin1'));
    expect(paths).to.deep.equal(['\x02', '\x02', '\x02', '\x02', '\x02', ' ']);
    expect(records.filter(x => x.type === R.hlink)).to.have.length(1);
  });

  it("keeps a button that runs the kept project's own macro, and drops one that runs another workbook's", async () => {
    const r = await disarmOffice(fixture('vba', 'button.xls'), { scriptPlugins: [passAll] });
    expect(of(r.before, 'MACRO/LINK').map(f => [f.action, f.data?.macro])).to.deep.equal([
      ['info', 'SumColumn'],
      ['strip', 'Remote'],
    ]);
    expect(must(r.after, 'after').status).to.equal('clean');
    const { records } = await workbookRecords(must(r.bytes, 'output'));
    const macros = records.filter(x => x.type === R.obj).map(x => Buffer.from(x.body).includes(Buffer.from([0x04, 0x00, 0x0e, 0x00])));
    expect(macros).to.deep.equal([true, false]);
  });

  it('rewrites an Excel 5.0 workbook: its macro sheet empty, and its Auto_Open name renamed without a flag byte', async () => {
    const rec = (type: number, body: number[]) => Buffer.from([type & 0xff, type >> 8, body.length & 0xff, body.length >> 8, ...body]);
    const bof = (dt: number) => rec(R.bof, [0x00, 0x05, dt, 0x00, 0xbb, 0x0d, 0xcc, 0x07]);
    const autoOpen = rec(R.lbl, [0x20, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]);
    const bound = (pos: number) => rec(R.boundSheet, [pos & 0xff, (pos >> 8) & 0xff, 0, 0, 0, 1, 6, ...Buffer.from('Macro1')]);
    const globals = Buffer.concat([bof(5), autoOpen, bound(0), rec(R.eof, [])]);
    const sheet = Buffer.concat([bof(0x40), rec(0x0006, [0, 0, 0, 0, 0, 0, ...new Array(16).fill(0)]), rec(R.eof, [])]);
    const stream = Buffer.concat([bof(5), autoOpen, bound(globals.length), rec(R.eof, []), sheet]);
    const sink = bufferSink();
    await writeCfb({ clsid: '00020810-0000-0000-C000-000000000046', children: [{ kind: 'stream', name: 'Book', size: stream.length, source: bufferSource(stream) }] }, sink);
    const r = await disarmOffice(sink.result());
    expect([r.status, must(r.after, 'after').status]).to.deep.equal(['defused', 'clean']);
    const cfb = await CompoundFile.open(bufferSource(must(r.bytes, 'output')), LIMITS);
    const book = must(cfb.child(cfb.root, 'Book'), 'Book');
    const reader = new RecordReader(cfb, book, 0);
    const records: BiffRecord[] = [];
    for (let x = await reader.next(); x; x = await reader.next()) records.push({ ...x, body: Buffer.from(x.body) });
    const lbl = must(
      records.find(x => x.type === R.lbl),
      'Lbl',
    );
    expect(Buffer.from(lbl.body.subarray(14, 14 + lbl.body[3])).toString('latin1')).to.equal('Removed_Auto_Open_1');
    const sheetRecord = must(
      records.find(x => x.type === R.boundSheet),
      'BoundSheet',
    );
    expect(sheetRecord.body[5]).to.equal(0);
    expect(records.filter(x => x.type === 0x0006)).to.deep.equal([]);
  });

  it("counts a button's macro as the kept project's own only when its formula is one PtgNameX of the workbook itself", async () => {
    // The fixture's first button, rebuilt with a PtgAttr in its place and with a token after its PtgNameX.
    const cfb = await CompoundFile.open(bufferSource(fixture('vba', 'button.xls')), LIMITS);
    const workbook = must(cfb.child(cfb.root, 'Workbook'), 'Workbook');
    const stream = Buffer.from(await cfb.bytes(workbook));
    const own = Buffer.from('0700000000003900000100000000', 'hex');
    for (const [label, swap] of [
      ['a PtgAttr', Buffer.from('0700000000001900000100000000', 'hex')],
      ['a PtgNameX and more', Buffer.from('0800000000003900000100000016', 'hex')],
    ] as const) {
      const where = stream.indexOf(own);
      expect(where, label).to.be.greaterThan(0);
      const changed = Buffer.from(stream);
      swap.copy(changed, where);
      const children = await Promise.all(
        cfb.root.children.map(async c =>
          c.name === 'Workbook'
            ? { kind: 'stream' as const, name: c.name, size: changed.length, source: bufferSource(changed) }
            : (await import('../src/rewrite-cfb')).subtree(cfb, cfb.root).children.find(n => n.name === c.name),
        ),
      );
      const sink = bufferSink();
      await writeCfb({ clsid: '00020820-0000-0000-C000-000000000046', children: children.filter(n => n !== undefined) }, sink);
      const r = await disarmOffice(sink.result(), { scriptPlugins: [passAll] });
      expect(
        of(r.before, 'MACRO/LINK').map(f => f.action),
        label,
      ).to.deep.equal(['strip', 'strip']);
    }
  });

  it('reports a SupBook path that starts with U+0004 as the external workbook it is', async () => {
    const rec = (type: number, body: Buffer) => Buffer.concat([Buffer.from([type & 0xff, type >> 8, body.length & 0xff, body.length >> 8]), body]);
    const bof = (dt: number) => rec(R.bof, Buffer.from([0x00, 0x06, dt, 0x00, 0xbb, 0x0d, 0xcc, 0x07, 0, 0, 0, 0, 0, 0, 0, 0]));
    const path = Buffer.from('\x04inert\x03book.xls', 'latin1');
    const supbook = rec(T.supBook, Buffer.concat([Buffer.from([1, 0, path.length, 0, 0]), path, Buffer.from([6, 0, 0]), Buffer.from('Sheet1')]));
    const head = Buffer.concat([bof(5), supbook]);
    const bound = (pos: number) => rec(R.boundSheet, Buffer.concat([Buffer.from([pos & 0xff, (pos >> 8) & 0xff, 0, 0, 0, 0, 6, 0]), Buffer.from('Sheet1')]));
    const sheetAt = head.length + bound(0).length + 4;
    const stream = Buffer.concat([head, bound(sheetAt), rec(R.eof, Buffer.alloc(0)), bof(0x10), rec(R.eof, Buffer.alloc(0))]);
    const sink = bufferSink();
    await writeCfb({ clsid: '00020820-0000-0000-C000-000000000046', children: [{ kind: 'stream', name: 'Workbook', size: stream.length, source: bufferSource(stream) }] }, sink);
    const r = await disarmOffice(sink.result());
    expect(of(r.before, 'EXTERNAL/LOCAL_FILE').map(f => f.data?.target)).to.deep.equal(['\x04inert\x03book.xls']);
    expect(must(r.after, 'after').status).to.equal('clean');
  });

  it('writes the stream positions of sheets, the DefColWidth and DBCell records, and shared strings anew', async () => {
    const rec = (type: number, body: Buffer) => Buffer.concat([Buffer.from([type & 0xff, type >> 8, body.length & 0xff, body.length >> 8]), body]);
    const le32 = (n: number) => {
      const b = Buffer.alloc(4);
      b.writeUInt32LE(n);
      return b;
    };
    const bof = (dt: number) => rec(R.bof, Buffer.from([0x00, 0x06, dt, 0x00, 0xbb, 0x0d, 0xcc, 0x07, 0, 0, 0, 0, 0, 0, 0, 0]));
    // Globals: a built-in Auto_Open name, which the rewrite makes longer, then the sheet, the strings and their index.
    const autoOpen = rec(R.lbl, Buffer.concat([Buffer.from([0x20, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), Buffer.from([0, 1])]));
    const sst = rec(T.sst, Buffer.concat([le32(2), le32(2), Buffer.from([5, 0, 0]), Buffer.from('inert'), Buffer.from([6, 0, 0]), Buffer.from('second')]));
    const head = Buffer.concat([bof(5), autoOpen]);
    const boundAt = head.length;
    const bound = (pos: number) => rec(R.boundSheet, Buffer.concat([le32(pos), Buffer.from([0, 0, 6, 0]), Buffer.from('Sheet1')]));
    const sstAt = boundAt + bound(0).length;
    const extsst = rec(T.extsst, Buffer.concat([Buffer.from([8, 0]), le32(sstAt + 4 + 8), Buffer.from([8, 0, 0, 0])]));
    const globals = Buffer.concat([head, bound(0), sst, extsst, rec(R.eof, Buffer.alloc(0))]);
    // The sheet: Index pointing at DefColWidth and the block's DBCell, which points back at its first Row.
    const sheetAt = globals.length;
    const indexLength = 4 + 20;
    const defColAt = sheetAt + bof(0x10).length + indexLength;
    const defCol = rec(T.defColWidth, Buffer.from([8, 0]));
    const rowAt = defColAt + defCol.length;
    const row = rec(T.row, Buffer.from([0, 0, 0, 0, 1, 0, 0xff, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
    const cell = rec(T.labelSst, Buffer.concat([Buffer.from([0, 0, 0, 0, 0, 0]), le32(1)]));
    const dbcellAt = rowAt + row.length + cell.length;
    const dbcell = rec(T.dbcell, Buffer.concat([le32(dbcellAt - rowAt), Buffer.from([0, 0])]));
    const index = rec(T.index, Buffer.concat([le32(0), le32(0), le32(1), le32(defColAt), le32(dbcellAt)]));
    const sheet = Buffer.concat([bof(0x10), index, defCol, row, cell, dbcell, rec(0x023e, Buffer.alloc(18)), rec(R.eof, Buffer.alloc(0))]);
    const stream = Buffer.concat([globals.subarray(0, boundAt), bound(sheetAt), globals.subarray(boundAt + bound(0).length), sheet]);
    const sink = bufferSink();
    await writeCfb({ clsid: '00020820-0000-0000-C000-000000000046', children: [{ kind: 'stream', name: 'Workbook', size: stream.length, source: bufferSource(stream) }] }, sink);
    const r = await disarmOffice(sink.result());
    expect([r.status, must(r.after, 'after').status]).to.deep.equal(['defused', 'clean']);
    const { records, stream: out } = await workbookRecords(must(r.bytes, 'output'));
    const lbl = must(
      records.find(x => x.type === R.lbl),
      'Lbl',
    );
    expect(lbl.body.length).to.be.greaterThan(16);
    const sheetBof = must(
      at(
        records,
        u32(
          must(
            records.find(x => x.type === R.boundSheet),
            'BoundSheet',
          ).body,
          0,
        ),
      ),
      'sheet BOF',
    );
    expect(sheetBof.type).to.equal(R.bof);
    const idx = must(
      records.find(x => x.type === T.index),
      'Index',
    );
    expect(at(records, u32(idx.body, 12))?.type).to.equal(T.defColWidth);
    const db = must(at(records, u32(idx.body, 16)), 'DBCell');
    expect(db.type).to.equal(T.dbcell);
    expect(at(records, db.offset - u32(db.body, 0))?.type).to.equal(T.row);
    const ext = must(
      records.find(x => x.type === T.extsst),
      'ExtSST',
    );
    expect(Buffer.from(out.subarray(u32(ext.body, 2), u32(ext.body, 2) + 8)).toString('latin1')).to.equal('\x05\x00\x00inert');
  });
});
