import { findingSpec } from './findings';
import { OfficeCategory as C, OfficeDetail as D, type DefuseFinding, type OfficeOptions, OfficeRisk, type OfficeScoreBands } from './types';

export const DEFAULT_BANDS: OfficeScoreBands = { low: 1, medium: 25, high: 50, critical: 80 };

interface Combination {
  name: string;
  weight: number;
  test(has: (c: C, d?: D) => boolean): boolean;
}

/** Code that runs, in any of its forms. */
const runsCode = (has: (c: C, d?: D) => boolean) => has(C.Macro) || has(C.Dde);

export const COMBINATIONS: Combination[] = [
  {
    name: 'code with a template, object or include from outside',
    weight: 20,
    test: has => (runsCode(has) || has(C.Action, D.RunProgram)) && (has(C.External, D.Template) || has(C.External, D.Object) || has(C.External, D.Include)),
  },
  { name: 'code with an embedded object or control', weight: 15, test: has => runsCode(has) && (has(C.Embedded) || has(C.Control, D.ActiveX)) },
  {
    name: 'code with data hidden in the archive',
    weight: 10,
    test: has => runsCode(has) && (has(C.Container, D.LeadingBytes) || has(C.Container, D.HiddenData) || has(C.Container, D.TrailingBytes)),
  },
];

/** The caller's weights first, then the built-in table. A kind from another package scores at the weight it carries. */
export function weightOf(f: DefuseFinding, options: OfficeOptions): number {
  const w = options.scoreWeights ?? [];
  const exact = w.find(x => x.category === f.category && x.detail === f.detail);
  if (exact) return exact.weight;
  const cat = w.find(x => x.category === f.category && x.detail === undefined);
  if (cat) return cat.weight;
  return findingSpec(f.category, f.detail)?.weight ?? f.weight ?? 0;
}

/** Highest weight in full, a quarter of each other kind's weight, plus combinations, capped at 100. */
export function scoreFindings(findings: DefuseFinding[], options: OfficeOptions): { score: number; risk: OfficeRisk } {
  // Each kind of finding counts once, so twenty links of one kind do not add up to a critical file.
  const byKind = new Map<string, number>();
  for (const f of findings) {
    const k = `${f.category}/${f.detail}`;
    byKind.set(k, Math.max(byKind.get(k) ?? 0, weightOf(f, options)));
  }
  const weights = [...byKind.values()].sort((a, b) => b - a);
  let score = weights.length ? weights[0] + weights.slice(1).reduce((s, w) => s + w / 4, 0) : 0;
  const has = (c: C, d?: D) => findings.some(f => f.category === c && (d === undefined || f.detail === d) && weightOf(f, options) > 0);
  for (const combo of COMBINATIONS) if (combo.test(has)) score += combo.weight;
  score = Math.min(100, Math.round(score));
  return { score, risk: riskFor(score, options.scoreBands ?? DEFAULT_BANDS) };
}

export function riskFor(score: number, bands: OfficeScoreBands): OfficeRisk {
  if (score >= bands.critical) return OfficeRisk.Critical;
  if (score >= bands.high) return OfficeRisk.High;
  if (score >= bands.medium) return OfficeRisk.Medium;
  if (score >= bands.low) return OfficeRisk.Low;
  return OfficeRisk.None;
}
