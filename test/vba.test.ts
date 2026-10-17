import { createHash } from 'node:crypto';
import { expect } from 'chai';
import { type ContainedScript, disarmOffice, inspectOffice, type OfficeDisarmResult, type ScriptPlugin } from '../src';
import { type CfbEntry, CompoundFile, entryPath } from '../src/cfb';
import { bufferSource } from '../src/io';
import { compress, decompress, PURGED_VBA_PROJECT, readProject, VbaError } from '../src/vba';
import { entryNames, entryText } from './helpers/structure';
import { fixture, kinds, must, of } from './helpers/util';

const LIMITS = { entries: 10_000, decompressedBytes: 1 << 30, totalDecompressedBytes: 2 ** 31 };
const hex = (s: string) => Buffer.from(s.replace(/ /g, ''), 'hex');
const vba = (name: string) => fixture('vba', name);

/** A script plugin that passes every module unless `refuse` matches its text, and keeps what it was shown. */
function scripts(refuse?: RegExp, options: { throws?: boolean; accepts?: boolean } = {}): ScriptPlugin & { seen: ContainedScript[] } {
  const seen: ContainedScript[] = [];
  return {
    kind: 'script',
    name: 'test',
    seen,
    accepts: () => options.accepts ?? true,
    async process(script) {
      seen.push(script);
      if (options.throws) throw new Error('inert plugin failure');
      const bad = refuse?.test(script.text) ?? false;
      return { result: bad ? 'removed' : 'passed', findings: bad ? [{ category: 'TEST', detail: 'REFUSED', description: 'refused', action: 'strip', location: script.location }] : [] };
    },
  };
}

/** The compound file in a package entry, or the file itself. */
async function compoundOf(bytes: Uint8Array, part?: string): Promise<CompoundFile> {
  let data = bytes;
  if (part) {
    const { ZipArchive } = await import('../src/zip');
    const zip = await ZipArchive.open(bufferSource(bytes), LIMITS);
    const entry = must(
      zip.entries.find(e => e.name === part),
      part,
    );
    const chunks: Buffer[] = [];
    for await (const c of zip.read(entry)) chunks.push(Buffer.from(c));
    data = Buffer.concat(chunks);
  }
  return CompoundFile.open(bufferSource(data), LIMITS);
}

const storage = (cfb: CompoundFile, path: string) =>
  must(
    cfb.entries.find(e => entryPath(e) === path && e.kind === 'storage'),
    path,
  );
const streamsOf = (e: CfbEntry) => e.children.filter(c => c.kind === 'stream').map(c => c.name);

async function disarmed(bytes: Uint8Array, plugin: ScriptPlugin): Promise<OfficeDisarmResult & { bytes: Uint8Array }> {
  const r = await disarmOffice(bytes, { scriptPlugins: [plugin] });
  return { ...r, bytes: must(r.bytes, `output (${r.status}: ${kinds(r.before).join(', ')})`) };
}

/** Each module of the project in `vbaPath` is its source compressed at offset 0, and nothing compiled is left. */
async function expectPurged(cfb: CompoundFile, vbaPath: string): Promise<void> {
  const v = storage(cfb, vbaPath);
  const project = await readProject(cfb, v);
  for (const m of project.modules) {
    expect(m.offset, m.name).to.equal(0);
    expect(Buffer.from(await cfb.bytes(m.stream)), m.name).to.deep.equal(Buffer.from(compress(m.bytes)));
  }
  const compiled = must(
    v.children.find(c => c.name === '_VBA_PROJECT'),
    '_VBA_PROJECT',
  );
  expect(Buffer.from(await cfb.bytes(compiled))).to.deep.equal(Buffer.from(PURGED_VBA_PROJECT));
  expect(streamsOf(v).filter(n => n.startsWith('__SRP_'))).to.deep.equal([]);
}

describe('MS-OVBA compression', () => {
  it('decompresses the examples of MS-OVBA 3.2', () => {
    expect(Buffer.from(decompress(hex('01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E'))).toString('latin1')).to.equal('abcdefghijklmnopqrstuv.');
    const second = '#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa';
    expect(
      Buffer.from(decompress(hex('01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 30 6D 6E 6F 70 06 71 02 70 04 10 72 73 74 75 76 10 77 78 79 7A 00 3C'))).toString(
        'latin1',
      ),
    ).to.equal(second);
    expect(Buffer.from(decompress(hex('01 03 B0 02 61 45 00'))).toString('latin1')).to.equal('a'.repeat(73));
  });

  it('compresses the first and third examples as MS-OVBA does, and everything back to what it was', () => {
    expect(Buffer.from(compress(Buffer.from('abcdefghijklmnopqrstuv.', 'latin1'))).toString('hex')).to.equal('0119b000616263646566676800696a6b6c6d6e6f70007172737475762e');
    expect(Buffer.from(compress(Buffer.from('a'.repeat(73), 'latin1'))).toString('hex')).to.equal('0103b002614500');
    // Text, a long run, bytes that do not repeat (written as raw chunks), and sizes around the 4,096-byte chunk.
    const noise = Buffer.concat(Array.from({ length: 300 }, (_, i) => createHash('sha256').update(`inert ${i}`).digest()));
    const text = Buffer.from('Sub Fill()\r\n    Cells(1, 1).Value = 1\r\nEnd Sub\r\n'.repeat(400), 'latin1');
    for (const data of [text, Buffer.alloc(10000, 0x61), noise.subarray(0, 8192), Buffer.alloc(0), text.subarray(0, 4096), text.subarray(0, 4097)]) {
      expect(Buffer.from(decompress(compress(data)))).to.deep.equal(data);
    }
    // A last chunk that only a padded raw chunk could hold.
    expect(() => compress(noise.subarray(0, 4096 + 4000))).to.throw(VbaError);
  });

  it('refuses a copy token that reaches before its chunk, and reads chunk headers without their signature bits', () => {
    expect(() => decompress(hex('01 10 b0 01 41 ff ff'))).to.throw(VbaError, /before its chunk/);
    expect(() => decompress(hex('02 00 b0 00'))).to.throw(VbaError, /signature byte/);
    // The first example with the signature bits of its header cleared.
    expect(Buffer.from(decompress(hex('01 19 80 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E 00'))).toString('latin1')).to.equal('abcdefghijklmnopqrstuv.');
  });
});

describe('VBA project reader', () => {
  it('reads the modules, their kinds and sources, and the references from the dir and PROJECT streams', async () => {
    const cfb = await compoundOf(vba('reference.xlsm'), 'xl/vbaProject.bin');
    const project = await readProject(cfb, storage(cfb, 'VBA'));
    expect(project.codePage).to.equal(1252);
    expect(project.modules.map(m => [m.name, m.kind, m.offset])).to.deep.equal([['Sums', 'standard', 16]]);
    expect(project.modules[0].source).to.match(/^Attribute VB_Name = "Sums"\r\nPublic Sub SumColumn\(\)/);
    expect(project.references.map(r => [r.name, r.kind, r.libid.slice(0, 45)])).to.deep.equal([
      ['Office', 'registered', '*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0'],
      ['MSXML2', 'registered', '*\\G{F5078F18-C551-11D3-89B9-0000F81FE221}#6.0'],
    ]);
  });

  it('refuses a source that does not decompress', async () => {
    const cfb = await compoundOf(vba('unreadable.docm'), 'word/vbaProject.bin');
    const e = await readProject(cfb, storage(cfb, 'VBA')).then(
      () => undefined,
      (x: unknown) => x,
    );
    expect(e).to.be.instanceOf(VbaError);
  });
});

describe('script plugins', () => {
  it('shows each module to the plugins, with the references on the first, and keeps the project they pass without its p-code', async () => {
    const plugin = scripts();
    const r = await disarmed(vba('safe.xlsm'), plugin);
    expect(plugin.seen.map(s => [s.fieldName, s.trigger, s.location, s.language, s.form, s.references?.length])).to.deep.equal([
      ['ThisWorkbook', 'vba-document', 'xl/vbaProject.bin > VBA/ThisWorkbook', 'vba', 'office', 1],
      ['Sums', 'vba-standard', 'xl/vbaProject.bin > VBA/Sums', 'vba', 'office', undefined],
    ]);
    expect(r.status).to.equal('defused');
    expect(r.removed.map(f => `${f.category}/${f.detail}`)).to.deep.equal(['MACRO/COMPILED_CODE']);
    // The plugins' answers are final, so office-defuse reports nothing of its own about the project.
    expect(of(r.before, 'MACRO/VBA_PROJECT')).to.deep.equal([]);
    expect(of(r.before, 'MACRO/PLUGIN_PASSED').map(f => f.location)).to.deep.equal(['xl/vbaProject.bin > VBA/ThisWorkbook', 'xl/vbaProject.bin > VBA/Sums']);
    // The button runs a macro of the project, so it stays.
    expect(of(r.before, 'MACRO/LINK').map(f => f.action)).to.deep.equal(['info']);
    expect(await entryText(r.bytes, 'xl/drawings/vmlDrawing1.vml')).to.include('<x:FmlaMacro>[0]!SumColumn</x:FmlaMacro>');
    // The output check found the project as the plugins passed it, without asking them again.
    expect([must(r.after, 'after').status, plugin.seen.length]).to.deep.equal(['clean', 2]);
    await expectPurged(await compoundOf(r.bytes, 'xl/vbaProject.bin'), 'VBA');
  });

  it('removes the project, and the button that runs another file, when a plugin removes a module', async () => {
    const r = await disarmed(vba('shell.xlsm'), scripts(/\bShell\b/));
    expect(r.status).to.equal('defused');
    expect(kinds(r.before)).to.include.members(['MACRO/PLUGIN_PASSED', 'MACRO/PLUGIN_REMOVED', 'TEST/REFUSED', 'MACRO/LINK']);
    // The plugin's answer removed the project, so office-defuse reports nothing of its own about it.
    expect(of(r.before, 'MACRO/VBA_PROJECT')).to.deep.equal([]);
    expect(of(r.before, 'MACRO/LINK').map(f => f.action)).to.deep.equal(['strip']);
    expect((await entryNames(r.bytes)).filter(n => /vba/i.test(n))).to.deep.equal([]);
    expect(await entryText(r.bytes, 'xl/drawings/vmlDrawing1.vml')).to.not.include('FmlaMacro');
  });

  it('removes the project when a plugin fails, or no plugin takes a module', async () => {
    const failed = await disarmed(vba('safe.xlsm'), scripts(undefined, { throws: true }));
    expect(of(failed.before, 'MACRO/PLUGIN_FAILED').length).to.equal(2);
    expect(of(failed.before, 'MACRO/VBA_PROJECT').map(f => f.action)).to.deep.equal(['strip']);
    const untaken = await disarmed(vba('safe.xlsm'), scripts(undefined, { accepts: false }));
    expect(kinds(untaken.before).filter(k => k.startsWith('MACRO/PLUGIN'))).to.deep.equal([]);
    expect(of(untaken.before, 'MACRO/VBA_PROJECT').map(f => f.action)).to.deep.equal(['strip']);
    expect((await entryNames(untaken.bytes)).filter(n => /vbaProject/.test(n))).to.deep.equal([]);
  });

  it('writes a stomped project with the source it shows and without the p-code that differs from it', async () => {
    const before = await compoundOf(vba('stomped.docm'), 'word/vbaProject.bin');
    const module = must(
      storage(before, 'VBA').children.find(c => c.name === 'ThisDocument'),
      'ThisDocument',
    );
    expect(Buffer.from(await before.bytes(module)).toString('latin1')).to.include('inert-program.exe');
    const r = await disarmed(vba('stomped.docm'), scripts());
    const after = await compoundOf(r.bytes, 'word/vbaProject.bin');
    await expectPurged(after, 'VBA');
    const kept = must(
      storage(after, 'VBA').children.find(c => c.name === 'ThisDocument'),
      'ThisDocument',
    );
    expect(Buffer.from(await after.bytes(kept)).toString('latin1')).to.not.include('inert-program.exe');
  });

  it('removes a project whose source does not read, saying why', async () => {
    const plugin = scripts();
    const r = await disarmed(vba('unreadable.docm'), plugin);
    expect(plugin.seen).to.deep.equal([]);
    expect(of(r.before, 'MACRO/VBA_PROJECT').map(f => [f.action, f.data?.reason])).to.deep.equal([['strip', 'a copy token that reaches before its chunk']]);
  });

  it("keeps a user form's designer with its project, and removes the project when a form may hold an outside control", async () => {
    const r = await disarmed(vba('form.xlsm'), scripts());
    expect([r.status, ...r.removed.map(f => `${f.category}/${f.detail}`)]).to.deep.equal(['defused', 'MACRO/COMPILED_CODE']);
    const out = await compoundOf(r.bytes, 'xl/vbaProject.bin');
    expect(streamsOf(storage(out, 'PickForm')).sort()).to.deep.equal(['\x01CompObj', '\x03VBFrame', 'f', 'o']);
    const activex = await disarmed(vba('form-activex.xlsm'), scripts());
    expect(of(activex.before, 'CONTROL/ACTIVEX').map(f => [f.location, f.data?.reason])).to.deep.equal([
      ['xl/vbaProject.bin > PickForm', 'a form names controls in its class table, which may be any ActiveX control'],
    ]);
    expect(of(activex.before, 'MACRO/VBA_PROJECT').map(f => f.action)).to.deep.equal(['strip']);
  });

  it('keeps and purges a project in a Word document and in a workbook', async () => {
    for (const [name, path] of [
      ['safe.doc', 'Macros/VBA'],
      ['safe.xls', '_VBA_PROJECT_CUR/VBA'],
    ]) {
      const r = await disarmed(vba(name), scripts());
      expect([r.status, must(r.after, 'after').status, ...r.removed.map(f => `${f.category}/${f.detail}`)], name).to.deep.equal(['defused', 'clean', 'MACRO/COMPILED_CODE']);
      await expectPurged(await compoundOf(r.bytes), path);
    }
  });

  it('scores a project the plugins answered for by their findings alone, and its own finding only when they did not', async () => {
    const weighs = (weight: number): ScriptPlugin => ({
      kind: 'script',
      name: 'scorer',
      accepts: () => true,
      process: async s => ({
        result: weight ? 'removed' : 'passed',
        findings: [{ category: 'TEST', detail: weight ? 'REFUSED' : 'FORM', description: 'test', action: weight ? 'strip' : 'info', location: s.location, weight }],
      }),
    });
    const kept = await disarmOffice(vba('safe.xlsm'), { scriptPlugins: [weighs(0)] });
    const removed = await disarmOffice(vba('safe.xlsm'), { scriptPlugins: [weighs(35)] });
    const failed = await disarmOffice(vba('safe.xlsm'), { scriptPlugins: [scripts(undefined, { throws: true })] });
    const weightOf = (r: OfficeDisarmResult, kind: string) => r.before.findings.filter(f => `${f.category}/${f.detail}` === kind).length;
    expect({
      // The button that runs a kept macro weighs nothing either; the compiled code office-defuse removes itself does.
      kept: [kept.before.risk, kept.before.score, weightOf(kept, 'MACRO/VBA_PROJECT'), must(kept.after, 'after').score],
      removed: [removed.before.score, weightOf(removed, 'MACRO/VBA_PROJECT')],
      failed: [failed.before.score, weightOf(failed, 'MACRO/VBA_PROJECT')],
    }).to.deep.equal({ kept: ['LOW', 10, 0, 0], removed: [35, 0], failed: [83, 1] });
  });

  it('decides in an inspection too, and leaves a project to the file plugins without script plugins', async () => {
    const inspected = await inspectOffice(vba('safe.xlsm'), { scriptPlugins: [scripts()] });
    expect([inspected.status, of(inspected, 'MACRO/VBA_PROJECT').length, of(inspected, 'MACRO/PLUGIN_PASSED').length]).to.deep.equal(['strippable', 0, 2]);
    const without = await inspectOffice(vba('safe.xlsm'));
    expect(of(without, 'MACRO/VBA_PROJECT').map(f => [f.location, f.action])).to.deep.equal([['xl/vbaProject.bin', 'strip']]);
  });
});
