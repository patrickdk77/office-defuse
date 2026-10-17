import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { expect } from 'chai';
import { disarmOffice } from '../src';
import { fixture, fixtures, must, runCli, wordPackage } from './helpers/util';

/** The fields of `--json` output the tests read. */
interface InspectJson {
  status: string;
  findings: Array<{ category: string; detail: string; action: string }>;
}

const at = (...parts: string[]) => path.join(fixtures, ...parts);
const json = (args: string[], opts: { input?: Buffer; env?: NodeJS.ProcessEnv } = {}) => {
  const r = runCli([...args, '--json'], opts);
  return { status: r.status, report: JSON.parse(r.stdout.toString()) as InspectJson };
};
const details = (r: InspectJson, category: string) => r.findings.filter(f => f.category === category).map(f => `${f.detail}:${f.action}`);

describe('command-line tool', () => {
  let dir: string;
  before(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-defuse-cli-'));
    fs.writeFileSync(path.join(dir, 'keep.js'), "module.exports = { plugins: [{ kind: 'file', name: 'keepall', accepts: () => true, process: async () => 'passed' }] };\n");
    fs.writeFileSync(path.join(dir, 'config.js'), "module.exports = { actionOverrides: [{ category: 'DDE', detail: 'FIELD', action: 'reject' }], limits: { entries: 1 } };\n");
    fs.writeFileSync(path.join(dir, 'script.js'), "module.exports = { plugins: [{ kind: 'script', name: 'x', accepts: () => true, process: async () => ({ result: 'passed' }) }] };\n");
  });
  after(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('inspects with exit codes 0, 1 and 2', () => {
    expect(runCli(['inspect', at('zip', 'minimal.docx')]).status).to.equal(0);
    const dde = runCli(['inspect', at('cases', 'dde-field.docx')]);
    const lines = dde.stdout.toString().split('\n');
    expect([dde.status, lines[0].startsWith('inspection: strippable, risk '), lines[0].endsWith(', DOCX'), lines[1].startsWith('  strip  DDE/FIELD [')]).to.deep.equal([1, true, true, true]);
    expect(runCli(['inspect', at('cases', 'billion-laughs.xlsx')]).status).to.equal(2);
  });

  it('prints JSON', () => {
    const r = json(['inspect', at('cases', 'dde-field.docx')]);
    expect([r.status, r.report.status, r.report.findings.map(f => f.detail)]).to.deep.equal([1, 'strippable', ['FIELD']]);
  });

  it('defuses to a file and to stdout, reading stdin', async () => {
    const expected = must((await disarmOffice(fixture('cases', 'dde-field.docx'))).bytes, 'the defused copy');
    const out = path.join(dir, 'out.docx');
    const r = runCli(['defuse', at('cases', 'dde-field.docx'), out]);
    expect([r.status, r.stdout.toString().split('\n')[0], Buffer.compare(fs.readFileSync(out), expected)]).to.deep.equal([1, 'result: defused', 0]);
    const piped = runCli(['defuse', '-', '-'], { input: fixture('cases', 'dde-field.docx') });
    expect([piped.status, Buffer.compare(piped.stdout, expected)]).to.deep.equal([1, 0]);
    // The report goes to stderr when the file goes to stdout.
    expect(piped.stderr.toString().split('\n')[0]).to.equal('result: defused');
    const clean = runCli(['defuse', at('zip', 'minimal.docx'), path.join(dir, 'minimal.docx')]);
    const rewritten = must((await disarmOffice(fixture('zip', 'minimal.docx'))).bytes, 'the clean copy');
    expect([clean.status, Buffer.compare(fs.readFileSync(path.join(dir, 'minimal.docx')), rewritten)]).to.deep.equal([0, 0]);
  });

  it('exits 2 for a rejected defuse and writes no output file', () => {
    const out = path.join(dir, 'rejected.xlsx');
    expect(runCli(['defuse', at('cases', 'billion-laughs.xlsx'), out]).status).to.equal(2);
    expect(fs.existsSync(out)).to.equal(false);
  });

  it('loads file plugins, and keeps embedded Office files with --defuse-embedded-office', () => {
    const embedded = (args: string[]) => {
      const r = json(['inspect', at('cases', 'chart-workbook.docx'), ...args]);
      return [r.status, details(r.report, 'EMBEDDED')];
    };
    expect(embedded([])).to.deep.equal([1, ['PACKAGE:strip']]);
    expect(embedded(['--defuse-embedded-office'])).to.deep.equal([0, ['PLUGIN_PASSED:info']]);
    expect(embedded(['--plugin', path.join(dir, 'keep.js')])).to.deep.equal([0, ['PLUGIN_PASSED:info']]);
    const script = runCli(['inspect', at('zip', 'minimal.docx'), '--plugin', path.join(dir, 'script.js')]);
    expect([script.status, script.stderr.toString().split('\n')[0]]).to.deep.equal([3, 'A plugin module must export plugins with kind "file"']);
  });

  it('reads passwords from a file, one per line, or one from OFFICE_DEFUSE_PASSWORD', () => {
    const file = at('ole', 'password.docx');
    const encrypted = (args: string[], env: NodeJS.ProcessEnv = {}) => {
      const r = json(['inspect', file, ...args], { env });
      return [r.status, details(r.report, 'ENCRYPTED')];
    };
    const opened = [1, ['USER_PASSWORD:strip', 'AES_256:info']];
    expect(encrypted([], { OFFICE_DEFUSE_PASSWORD: '' })[0]).to.equal(2);
    const list = path.join(dir, 'passwords.txt');
    // CRLF line ends, a blank line, and the right password last with no line end after it.
    fs.writeFileSync(list, 'wrong\r\n\r\nalso wrong\ninert-password');
    expect(encrypted(['--password-file', list])).to.deep.equal(opened);
    expect(encrypted([], { OFFICE_DEFUSE_PASSWORD: 'inert-password' })).to.deep.equal(opened);
    // The file replaces the environment variable.
    const wrong = path.join(dir, 'wrong.txt');
    fs.writeFileSync(wrong, 'wrong\n');
    expect(encrypted(['--password-file', wrong], { OFFICE_DEFUSE_PASSWORD: 'inert-password' })[0]).to.equal(2);
    // A password on the command line is refused, and the message does not repeat it.
    const argv = runCli(['inspect', file, '--password=inert-password']);
    expect([argv.status, argv.stderr.toString().split('\n')[0], argv.stderr.toString().includes('inert-password')]).to.deep.equal([
      3,
      'Pass the password with --password-file or OFFICE_DEFUSE_PASSWORD, not on the command line',
      false,
    ]);
  });

  it('loads a config module and parses limits, the flags applying on top', () => {
    const rejects = (args: string[]) => {
      const r = json(['inspect', at('cases', 'dde-field.docx'), '--config', path.join(dir, 'config.js'), ...args]);
      return { status: r.status, rejects: r.report.findings.filter(f => f.action === 'reject').map(f => `${f.category}/${f.detail}`) };
    };
    expect(rejects([])).to.deep.equal({ status: 2, rejects: ['LIMIT/ENTRY_COUNT'] });
    expect(rejects(['--limit', 'entries=100'])).to.deep.equal({ status: 2, rejects: ['DDE/FIELD'] });
    expect(runCli(['inspect', at('cases', 'dde-field.docx'), '--limit', 'fileSize=1mb']).status).to.equal(1);
    expect(runCli(['inspect', at('cases', 'dde-field.docx'), '--limit', 'fileSize=10']).status).to.equal(2);
    const big = json(['inspect', '-', '--limit', 'fileSize=10'], { input: fixture('zip', 'minimal.docx') });
    expect(big.report.findings.map(f => `${f.category}/${f.detail}`)).to.deep.equal(['LIMIT/FILE_SIZE']);
  });

  it('exits 3 on usage errors and I/O errors', () => {
    expect(runCli([]).status).to.equal(3);
    expect(runCli(['inspect']).status).to.equal(3);
    // A clean input, so an ignored option would exit 0, and the message tells it apart from a stray positional.
    const bogus = runCli(['inspect', at('zip', 'minimal.docx'), '--bogus=secret']);
    expect([bogus.status, bogus.stderr.toString().split('\n')[0]]).to.deep.equal([3, 'Unknown option --bogus=...']);
    expect(runCli(['inspect', at('zip', 'minimal.docx'), '--limit', 'entries=many']).status).to.equal(3);
    expect(runCli(['inspect', at('zip', 'minimal.docx'), '--limit', 'pixels=4']).status).to.equal(3);
    expect(runCli(['inspect', at('zip', 'minimal.docx'), '--password-file', path.join(dir, 'missing.txt')]).status).to.equal(3);
    expect(runCli(['defuse', path.join(dir, 'missing.docx'), path.join(dir, 'x.docx')]).status).to.equal(3);
    const same = path.join(dir, 'same.docx');
    fs.writeFileSync(same, fixture('zip', 'minimal.docx'));
    const r = runCli(['defuse', same, same]);
    expect([r.status, r.stderr.toString().split('\n')[0], fs.readFileSync(same).length]).to.deep.equal([3, 'The input and the output are the same file', fixture('zip', 'minimal.docx').length]);
  });

  it('shows control characters from the file as escapes, in text and in JSON', async () => {
    const file = path.join(dir, 'escape.docx');
    // A part name holding ESC and the C1 control CSI, which a terminal can run as commands.
    fs.writeFileSync(file, await wordPackage('<w:p/>', { parts: { 'word/x\u001b[31m\u009b2J.bin': 'x' } }));
    const text = runCli(['inspect', file]).stdout.toString();
    expect([text.includes('\u001b'), text.includes('\u009b'), text.includes('\\u001b[31m\\u009b2J')]).to.deep.equal([false, false, true]);
    const out = runCli(['inspect', file, '--json']).stdout.toString();
    expect([out.includes('\\u001b[31m'), out.includes('\\u009b2J'), out.includes('\u009b')]).to.deep.equal([true, true, false]);
  });
});
