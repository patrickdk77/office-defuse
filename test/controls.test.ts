import { expect } from 'chai';
import { disarmOffice, inspectOffice } from '../src';
import { CompoundFile, entryPath } from '../src/cfb';
import { bufferSource } from '../src/io';
import { entryNames } from './helpers/structure';
import { fixture, kinds, must, of } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const controls = (name: string) => fixture('controls', name);

describe('Forms 2.0 controls', () => {
  it('keeps Forms 2.0 controls persisted as a stream, a storage and a property bag, with their binary parts', async () => {
    const r = await disarmOffice(controls('forms.xlsm'));
    expect(of(r.before, 'CONTROL/FORMS_CONTROL').map(f => [f.location, f.data?.clsid])).to.deep.equal([
      ['xl/activeX/activeX1.xml', '8BD21D40-EC42-11CE-9E0D-00AA006002F3'],
      ['xl/activeX/activeX2.xml', 'D7053240-CE69-11CD-A777-00DD01143C57'],
      ['xl/activeX/activeX3.xml', '8BD21D50-EC42-11CE-9E0D-00AA006002F3'],
    ]);
    expect(of(r.before, 'CONTROL/ACTIVEX').map(f => [f.location, f.data?.reason])).to.deep.equal([['xl/activeX/activeX4.xml', 'its class is not a Forms 2.0 control office-defuse reads']]);
    expect(r.status).to.equal('defused');
    const names = (await entryNames(must(r.bytes, 'output'))).filter(n => n.startsWith('xl/activeX/'));
    expect(names.sort()).to.deep.equal([
      'xl/activeX/_rels/activeX1.xml.rels',
      'xl/activeX/_rels/activeX2.xml.rels',
      'xl/activeX/activeX1.bin',
      'xl/activeX/activeX1.xml',
      'xl/activeX/activeX2.bin',
      'xl/activeX/activeX2.xml',
      'xl/activeX/activeX3.xml',
    ]);
    expect([must(r.after, 'after').status, ...kinds(must(r.after, 'after'))]).to.deep.equal(['clean', 'CONTROL/FORMS_CONTROL', 'CONTROL/FORMS_CONTROL', 'CONTROL/FORMS_CONTROL']);
  });

  it('removes a Forms 2.0 control with a picture, bytes after its properties, a binary of another class, or a picture in its bag', async () => {
    const r = await inspectOffice(controls('forms-refused.xlsx'));
    expect(of(r, 'CONTROL/ACTIVEX').map(f => f.data?.reason)).to.deep.equal([
      'a control that holds a picture',
      "5 bytes after the control's properties",
      'its binary part names another class',
      'its Picture property loads a picture',
      'it holds elements besides plain ax:ocxPr values, such as a font or a picture',
    ]);
    expect(of(r, 'CONTROL/FORMS_CONTROL')).to.deep.equal([]);
  });

  it("keeps a Word document's Forms 2.0 control without its picture for printing, and removes another class", async () => {
    const r = await disarmOffice(controls('forms.doc'));
    expect(kinds(r.before)).to.deep.equal(['CONTROL/CONTROL_PICTURE', 'CONTROL/FORMS_CONTROL', 'CONTROL/ACTIVEX']);
    const out = await CompoundFile.open(bufferSource(must(r.bytes, 'output')), LIMITS);
    expect(out.entries.map(entryPath).filter(p => p.startsWith('ObjectPool/'))).to.deep.equal([
      'ObjectPool/_1000',
      'ObjectPool/_1000/\\x01CompObj',
      'ObjectPool/_1000/\\x03ObjInfo',
      'ObjectPool/_1000/\\x03OCXNAME',
      'ObjectPool/_1000/contents',
    ]);
    expect(must(r.after, 'after').status).to.equal('clean');
  });

  it("keeps a workbook's Ctls stream when its Obj records place Forms 2.0 controls over all of it", async () => {
    const kept = await disarmOffice(controls('forms.xls'));
    expect([kept.status, ...of(kept.before, 'CONTROL/FORMS_CONTROL').map(f => f.data?.progId)]).to.deep.equal(['clean', 'Forms.CheckBox.1', 'Forms.CommandButton.1']);
    const out = await CompoundFile.open(bufferSource(must(kept.bytes, 'output')), LIMITS);
    expect(out.entries.map(entryPath)).to.include('Ctls');
    const gap = await disarmOffice(controls('forms-gap.xls'));
    expect(of(gap.before, 'CONTROL/ACTIVEX').map(f => f.data?.reason)).to.deep.equal(["bytes at offset 0 that no control's place covers"]);
    const dropped = await CompoundFile.open(bufferSource(must(gap.bytes, 'output')), LIMITS);
    expect(dropped.entries.map(entryPath)).to.not.include('Ctls');
  });

  it('removes Forms 2.0 controls too when CONTROL/FORMS_CONTROL is overridden to strip', async () => {
    const strip = { actionOverrides: [{ category: 'CONTROL' as never, detail: 'FORMS_CONTROL' as never, action: 'strip' as const }] };
    const r = await disarmOffice(controls('forms.xlsm'), strip);
    expect((await entryNames(must(r.bytes, 'output'))).filter(n => n.startsWith('xl/activeX/'))).to.deep.equal([]);
    const xls = await disarmOffice(controls('forms.xls'), strip);
    const out = await CompoundFile.open(bufferSource(must(xls.bytes, 'output')), LIMITS);
    expect(out.entries.map(entryPath)).to.not.include('Ctls');
  });
});
