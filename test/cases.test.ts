import * as fs from 'node:fs';
import * as path from 'node:path';
import { expect } from 'chai';
import { inspectOffice } from '../src';
import { fixtures } from './helpers/util';

interface Result {
  status: string;
  risk: string;
  format: string;
  kinds: string[];
}

interface Case {
  file: string;
  issue: string;
  expected: Result;
  before: Result;
}

const folder = path.join(fixtures, 'cases');
const manifest = JSON.parse(fs.readFileSync(path.join(folder, 'manifest.json'), 'utf8')) as Case[];

describe('corpus cases', () => {
  it('lists every file in test/fixtures/cases in the manifest, and nothing else', () => {
    const files = fs.readdirSync(folder).filter(f => f !== 'manifest.json');
    expect(manifest.map(c => c.file).sort()).to.deep.equal(files.sort());
  });

  for (const c of manifest) {
    it(`${c.file}: ${c.expected.status}${c.expected.kinds.length ? ` with ${c.expected.kinds.join(', ')}` : ''}`, async () => {
      const r = await inspectOffice(fs.readFileSync(path.join(folder, c.file)));
      const kinds = [...new Set(r.findings.map(f => `${f.category}/${f.detail}`))].sort();
      expect({ status: r.status, risk: r.risk, format: r.format, kinds }).to.deep.equal(c.expected);
    });
  }
});
