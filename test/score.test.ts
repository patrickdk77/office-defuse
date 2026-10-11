import * as fs from 'node:fs';
import * as path from 'node:path';
import { expect } from 'chai';
import { allFindingSpecs, OfficeCategory as C, COMBINATIONS, OfficeDetail as D, DEFAULT_BANDS, type DefuseFinding, inspectOffice, OfficeRisk, scoreFindings } from '../src';
import { FindingFactory } from '../src/findings';
import { riskFor } from '../src/score';
import { fixture, kinds, root } from './helpers/util';

const f = new FindingFactory();

describe('findings and scoring', () => {
  it('has a description, default action and weight for every pair it makes', () => {
    const specs = allFindingSpecs();
    expect(specs.length).to.be.greaterThan(70);
    for (const s of specs) {
      expect(Object.values(C)).to.include(s.category);
      expect(Object.values(D)).to.include(s.detail);
      expect(s.description).to.match(/^[A-Z<]/);
      expect(s.description).to.match(/^[\x20-\x7e]+$/);
      expect(['reject', 'strip', 'info']).to.include(s.action);
      expect(s.weight).to.be.within(0, 100);
    }
    // Every detail the source names must have a spec, and so must every category and detail named on one line.
    const known = new Set(specs.map(s => `${s.category}/${s.detail}`));
    const details = new Set<string>(specs.map(s => s.detail));
    const category = new Map<string, string>(Object.entries(C));
    const detail = new Map<string, string>(Object.entries(D));
    const missing: string[] = [];
    let pairs = 0;
    const src = path.join(root, 'src');
    for (const name of fs.readdirSync(src).filter(n => n.endsWith('.ts') && n !== 'types.ts')) {
      for (const line of fs.readFileSync(path.join(src, name), 'utf8').split('\n')) {
        for (const part of line.split(/(?=\bC\.\w)/)) {
          const c = /^C\.(\w+)/.exec(part);
          for (const m of part.matchAll(/\bD\.(\w+)/g)) {
            const d = detail.get(m[1]) ?? m[1];
            if (!details.has(d)) missing.push(`${name}: D.${m[1]}`);
            if (!c) continue;
            pairs++;
            if (!known.has(`${category.get(c[1])}/${d}`)) missing.push(`${name}: C.${c[1]} with D.${m[1]}`);
          }
        }
      }
    }
    expect({ missing, scanned: pairs > 100 }).to.deep.equal({ missing: [], scanned: true });
  });

  it('every detail in the enum is made by some spec', () => {
    const used = new Set<string>(allFindingSpecs().map(s => s.detail));
    expect(Object.values(D).filter(d => !used.has(d))).to.deep.equal([]);
  });

  it('ends every choice in a description with an alternative that needs no value', () => {
    for (const s of allFindingSpecs()) for (const [, body] of s.description.matchAll(/<([^<>]*)>/g)) expect(body.split('|').slice(-1)[0], s.description).to.not.match(/\{/);
  });

  it('refuses to make a finding of a kind it has no spec for', () => {
    expect(() => f.make('NOPE' as C, 'NOPE' as D)).to.throw('No finding spec for NOPE/NOPE');
  });

  it('fills placeholders and choices from data, and shows no control characters from the file', () => {
    expect(f.make(C.External, D.Template, undefined, { target: 'https://t.example/x.dotm' }).description).to.equal('The document loads its template from https://t.example/x.dotm');
    expect(f.make(C.External, D.Template).description).to.equal('The document loads its template from outside the file');
    expect(f.make(C.Package, D.UntypedPart).description).to.equal('A part has no content type');
    expect(f.make(C.Ui, D.Callback, undefined, { callback: 'onLoad=\u001b[2Jx' }).description).to.equal('The custom ribbon calls a macro: onLoad=?[2Jx');
    expect(f.make(C.External, D.Image, undefined, { target: `https://example.com/${'a'.repeat(100)}` }).description).to.have.length.lessThan(120);
  });

  it('applies overrides by category and by category and detail', () => {
    const o = new FindingFactory([
      { category: C.Link, action: 'reject' },
      { category: C.Link, detail: D.Safe, action: 'strip' },
    ]);
    expect(o.make(C.Link, D.Safe).action).to.equal('strip');
    expect(o.make(C.Link, D.IpHost).action).to.equal('reject');
    expect(o.make(C.Macro, D.VbaProject).action).to.equal('strip');
  });

  it('counts the highest weight in full and a quarter of each other kind, once per kind', () => {
    const findings = [f.make(C.Macro, D.VbaProject), f.make(C.Link, D.IpHost), f.make(C.Link, D.IpHost), f.make(C.Link, D.Safe)];
    expect(scoreFindings(findings, {})).to.deep.equal({ score: 80, risk: OfficeRisk.Critical });
    expect(scoreFindings([f.make(C.Link, D.IpHost)], {})).to.deep.equal({ score: 40, risk: OfficeRisk.Medium });
    expect(scoreFindings([], {})).to.deep.equal({ score: 0, risk: OfficeRisk.None });
  });

  it('adds the combinations', () => {
    expect(COMBINATIONS.map(c => c.weight)).to.deep.equal([20, 15, 10]);
    const score = (...list: DefuseFinding[]) => scoreFindings(list, {}).score;
    expect(score(f.make(C.Macro, D.OdfDialogs), f.make(C.External, D.Template))).to.equal(Math.round(60 + 30 / 4 + 20));
    expect(score(f.make(C.Dde, D.Field), f.make(C.External, D.Template))).to.equal(100);
    expect(score(f.make(C.Field, D.Macrobutton), f.make(C.External, D.Template))).to.equal(Math.round(60 + 40 / 4));
    expect(score(f.make(C.Macro, D.OdfDialogs), f.make(C.Embedded, D.Package))).to.equal(Math.round(30 + 30 / 4 + 15));
    expect(score(f.make(C.Macro, D.OdfDialogs), f.make(C.Container, D.HiddenData))).to.equal(Math.round(30 + 20 / 4 + 10));
  });

  it('takes weight and band overrides, and scores a finding from another package at its own weight', () => {
    const vba = f.make(C.Macro, D.VbaProject);
    expect(scoreFindings([vba], { scoreWeights: [{ category: C.Macro, weight: 10 }] }).score).to.equal(10);
    expect(
      scoreFindings([vba], {
        scoreWeights: [
          { category: C.Macro, detail: D.VbaProject, weight: 5 },
          { category: C.Macro, weight: 10 },
        ],
      }).score,
    ).to.equal(5);
    expect(scoreFindings([vba], { scoreBands: { low: 1, medium: 2, high: 3, critical: 100 } }).risk).to.equal(OfficeRisk.High);
    const foreign: DefuseFinding = { category: 'JAVASCRIPT', detail: 'OPEN_ACTION', description: 'x', action: 'strip', weight: 70 };
    expect(scoreFindings([foreign], {}).score).to.equal(70);
    expect(scoreFindings([{ ...foreign, weight: undefined }], {}).score).to.equal(0);
    expect(riskFor(25, DEFAULT_BANDS)).to.equal(OfficeRisk.Medium);
  });

  it('applies action overrides to a whole inspection, and info changes nothing but the finding', async () => {
    const base = await inspectOffice(fixture('docs', 'vba.docm'));
    expect(base.status).to.equal('strippable');
    const info = await inspectOffice(fixture('docs', 'vba.docm'), { actionOverrides: [{ category: C.Macro, action: 'info' }] });
    expect([info.status, info.score, kinds(info)]).to.deep.equal(['clean', base.score, kinds(base)]);
    const reject = await inspectOffice(fixture('docs', 'vba.docm'), { actionOverrides: [{ category: C.Macro, detail: D.VbaProject, action: 'reject' }] });
    expect([reject.status, reject.score, reject.risk]).to.deep.equal(['rejected', base.score, base.risk]);
    const safeLinks = await inspectOffice(fixture('lo', 'clean.docx'), { actionOverrides: [{ category: C.Link, detail: D.Safe, action: 'strip' }] });
    expect(safeLinks.status).to.equal('strippable');
  });

  it('rejects a run that stopped whatever the overrides say about the finding that stopped it', async () => {
    const r = await inspectOffice(fixture('cases', 'wrong-crc.docx'), { actionOverrides: [{ category: C.Container, action: 'info' }] });
    expect([r.status, r.score, r.risk, r.findings[0].action]).to.deep.equal(['rejected', null, 'UNKNOWN', 'info']);
  });
});
