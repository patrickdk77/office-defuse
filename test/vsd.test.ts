import { expect } from 'chai';
import { disarmOffice, inspectOffice } from '../src';
import { StopError } from '../src/errors';
import { inflate } from '../src/vsd';
import { fixture, kinds, of } from './helpers/util';

const ole = (name: string) => fixture('ole', name);
const cells = (r: { findings: Array<{ category: string; detail: string; data?: Record<string, string | number> }> }) =>
  r.findings.filter(f => f.category === 'MACRO').map(f => [`${f.category}/${f.detail}`, f.data]);

describe('Visio .vsd ShapeSheet cells', () => {
  it('reports the cells of a Visio 2003 drawing that call code, as a .vsdx gets them, and passes the cells that do not', async () => {
    const r = await inspectOffice(ole('vsd-cells.vsd'));
    expect(r.format).to.equal('vsd');
    expect(cells(r)).to.deep.equal([
      ['MACRO/EVENT', { event: 'EventDblClick', formula: 'RUNADDON("inert.exe")' }],
      ['MACRO/EVENT', { event: 'EventDrop', formula: 'RUNADDONWARGS("inert","/x")' }],
      ['MACRO/LINK', { section: 'Actions', cell: 'Action', formula: 'CALLTHIS("Module1.Inert")' }],
      // The call sits inside an IF, which the formula shown does not write out.
      ['MACRO/LINK', { section: 'User-defined Cells', cell: 'Value', formula: 'RUNMACRO(...)' }],
      ['MACRO/LINK', { section: 'Scratch', cell: 'cell 0', formula: `OPENFILE("${'x'.repeat(187)}...` }],
      ['MACRO/LINK', { section: 'Shape Data', cell: 'cell 0', formula: 'DOOLEVERB("Open")' }],
      // Smart Tags is a chunk type whose cells office-defuse cannot place, so the call is found by its token alone.
      ['MACRO/LINK', { section: 'Smart Tags', cell: 'unknown', formula: 'RUNADDON(...)' }],
    ]);
    expect(of(r, 'MACRO/EVENT')[0].description).to.equal('A macro or script runs on the EventDblClick event');
    expect(new Set(r.findings.map(f => f.location))).to.deep.equal(new Set(['VisioDocument']));
  });

  it('reads a Visio 2000 drawing, whose strings are 8-bit and whose streams are not compressed', async () => {
    expect(cells(await inspectOffice(ole('vsd2000-cells.vsd')))).to.deep.equal([
      ['MACRO/EVENT', { event: 'EventXFMod', formula: 'RUNADDONWARGS("inert","/y")' }],
      ['MACRO/LINK', { section: 'Actions', cell: 'Action', formula: 'RUNADDON("inert")' }],
    ]);
  });

  it('finds a call that lies outside every chunk the walk reads', async () => {
    expect(cells(await inspectOffice(ole('vsd-between-chunks.vsd')))).to.deep.equal([['MACRO/LINK', { section: 'between chunks', cell: 'unknown', formula: 'RUNADDON(...)' }]]);
  });

  it('refuses a drawing whose structure does not read', async () => {
    const reasons: Array<[string, string]> = [
      ['vsd-block-past-chunk.vsd', 'a cell block in a chunk of type 0x84 runs past its chunk'],
      ['vsd-unknown-token.vsd', 'a formula in a chunk of type 0x84 holds the unknown token 0xee'],
      ['vsd-chunk-past-stream.vsd', 'a chunk of type 0x84 runs past its stream'],
      ['vsd-pointer-past-end.vsd', 'a pointer names bytes 1048576 to 1048592, past the end of the stream'],
      ['vsd-pointer-count.vsd', 'a pointer list counts 1000 pointers, more than its stream holds'],
    ];
    for (const [name, reason] of reasons) {
      const r = await inspectOffice(ole(name));
      expect([name, r.status, r.score, kinds(r), r.findings[0]?.data]).to.deep.equal([name, 'rejected', null, ['PACKAGE/UNREADABLE'], { stream: 'VisioDocument', reason }]);
    }
  });

  it('bounds what the streams decompress to by the limits', async () => {
    expect((await inspectOffice(ole('vsd-large.vsd'))).status).to.equal('clean');
    expect(kinds(await inspectOffice(ole('vsd-large.vsd'), { limits: { decompressedBytes: 20_000 } }))).to.deep.equal(['LIMIT/DECOMPRESSED_SIZE']);
    expect(kinds(await inspectOffice(ole('vsd-large.vsd'), { limits: { totalDecompressedBytes: 30_000 } }))).to.deep.equal(['LIMIT/TOTAL_DECOMPRESSED_SIZE']);
  });

  it('leaves a VisioDocument stream of another version unread', async () => {
    expect((await inspectOffice(ole('visio.vsd'))).status).to.equal('clean');
  });

  it('rejects the drawing on disarm, since the rewrite copies the stream as it is', async () => {
    const r = await disarmOffice(ole('vsd-cells.vsd'));
    expect([r.status, r.bytes, of(r.before, 'PROCESSING/VERIFICATION_FAILED').length]).to.deep.equal(['rejected', undefined, 1]);
  });

  it('decompresses literals and references, which may overlap what they write, and stops past the limit', () => {
    // A literal 'a', then a reference to it for 18 bytes: window position 0 is stored as 4078.
    const src = Uint8Array.of(0b01, 0x61, 4078 & 0xff, ((4078 >> 8) << 4) | 15);
    expect(Buffer.from(inflate(src, 100)).toString()).to.equal('a'.repeat(19));
    expect(() => inflate(src, 10)).to.throw(StopError);
    // A reference cut short at the end ends the output, as libvisio reads it.
    expect(Buffer.from(inflate(Uint8Array.of(0b01, 0x61, 0x00), 100)).toString()).to.equal('a');
  });
});
