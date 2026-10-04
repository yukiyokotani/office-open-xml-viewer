import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import {
  powerPointAscentShare, powerPointShareCacheSize, SHARE_CACHE_LIMIT,
} from './powerpoint-line-metrics.js';
import type { Paragraph, TextBody } from './types.js';

// Expected values are PowerPoint 16.113.2 PDF baselines of the #1610
// controls, counted in whole 1/100 in from the text-area top (tIns 0 here).
// At this scale 1 pt = 1 canvas unit, so N units of 1/100 in are N × 0.72.
// The export quantizes each baseline to that device unit; layout keeps
// continuous positions, so each must lie within half a unit of the export.
const SCALE = 1 / 12700;
const U = 0.72;

type Spacing = Paragraph['spaceLine'];
interface RunSpec { text: string; font: string; size: number }

function context() {
  const draws: Array<{ text: string; y: number }> = [];
  let font = '';
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', strokeStyle: '', direction: 'ltr', textAlign: 'left', textBaseline: 'alphabetic',
    measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2 }),
    fillText: (text: string, _x: number, y: number) => draws.push({ text, y }),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws };
}

function paragraph(runs: RunSpec[], extra: Partial<Paragraph> = {}): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
    spaceLine: { type: 'pct', val: 100000 }, lvl: 0, bullet: { type: 'none' },
    defFontSize: null, defColor: null, defBold: null, defItalic: null, defFontFamily: null,
    tabStops: [], eaLnBrk: true,
    runs: runs.map((r): TextRunData => ({
      type: 'text', text: r.text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: r.size, color: '000000', fontFamily: r.font, fontFamilyEa: r.font,
    } as TextRunData)),
    ...extra,
  } as Paragraph;
}

function baselines(paragraphs: Paragraph[], over: Partial<TextBody> = {}, boxHeight = 800): number[] {
  const body = {
    verticalAnchor: 't', paragraphs, defaultFontSize: 18, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none', ...over,
  } as TextBody;
  const { ctx, draws } = context();
  renderTextBody(ctx, body, 0, 0, 2000, boxHeight, SCALE);
  const ys: number[] = [];
  for (const d of draws) if (!ys.some((y) => Math.abs(y - d.y) < 1e-6)) ys.push(d.y);
  return ys;
}

const lines = (font: string, size: number, n: number, spaceLine: Spacing = { type: 'pct', val: 100000 }, text = 'Hg') =>
  Array.from({ length: n }, (_, i) => paragraph([{ text: `${text}${i}`, font, size }], { spaceLine }));

/** Continuous baselines expressed in export units, snapped the way the
 * export snaps them (nearest unit), so they compare equal to the PDF. */
const units = (ys: number[]) => ys.map((y) => Math.floor(y / U + 0.5));

describe('PowerPoint text-box line metrics (#1610)', () => {
  it('splits each face by its usWin (or USE_TYPO_METRICS typo) ascent share', () => {
    // A-* controls: first and second baselines of two same-size lines.
    expect(units(baselines(lines('Arial', 72, 2)))).toEqual([97, 217]);
    expect(units(baselines(lines('Calibri', 72, 2)))).toEqual([94, 214]);
    expect(units(baselines(lines('Yu Gothic', 72, 2)))).toEqual([92, 212]);
    expect(units(baselines(lines('Gabriola', 72, 2)))).toEqual([98, 218]);
    expect(units(baselines(lines('Baskerville Old Face', 200, 2)))).toEqual([258, 591]);
    expect(units(baselines(lines('Stencil', 200, 2)))).toEqual([255, 589]);
    expect(units(baselines(lines('Times New Roman', 200, 2)))).toEqual([268, 602]);
    expect(units(baselines(lines('Georgia', 200, 2)))).toEqual([269, 602]);
    expect(units(baselines(lines('Courier New', 200, 2)))).toEqual([245, 578]);
  });

  it('resolves a face name that is one cut of its family to that cut (#1630)', () => {
    // "Calibri Light" names the weight-300 Calibri cut. Controls: a 44 pt
    // Calibri Light title's first baseline sat 41.16 pt below the text top at
    // single spacing, 0.7795 of the 52.8 pt line box (usWin 1950 / 2500).
    expect(powerPointAscentShare('Calibri Light', false, false)).toBeCloseTo(1950 / 2500, 12);
    expect(powerPointAscentShare('Calibri Light', false, true)).toBeCloseTo(1950 / 2500, 12);
    expect(powerPointAscentShare('Calibri Light', true, false)).toBeCloseTo(1950 / 2500, 12);
    // A name shared by several cuts still needs the requested weight.
    expect(powerPointAscentShare('Calibri', false, false)).toBeCloseTo(1950 / 2500, 12);
  });

  it('follows PowerPoint substitutions for faces only macOS itself provides', () => {
    expect(units(baselines(lines('Palatino', 200, 2)))).toEqual([259, 593]);
    expect(units(baselines(lines('Helvetica', 200, 2)))).toEqual([270, 603]);
    expect(powerPointAscentShare('Helvetica Neue', false, false))
      .toBe(powerPointAscentShare('Arial', false, false));
    // Avenir, Menlo and Hiragino Sans fall back to deck-dependent faces: unresolved.
    expect(powerPointAscentShare('Avenir', false, false)).toBeUndefined();
    expect(powerPointAscentShare('Hiragino Sans', false, false)).toBeUndefined();
  });

  it('rounds each baseline, not the line pitch, to a whole 1/100 inch', () => {
    // 100 pt lines are 166.67 units high: Arial steps 167, Meiryo 166.
    expect(units(baselines(lines('Arial', 100, 2)))).toEqual([135, 302]);
    expect(units(baselines(lines('Meiryo', 100, 2, null, '日')))).toEqual([118, 284]);
    // Fractional sizes keep an exact 1.2 × size line (Z-*).
    expect(units(baselines(lines('Arial', 10.5, 8)))).toEqual([14, 32, 49, 67, 84, 102, 119, 137]);
    expect(units(baselines(lines('Meiryo', 13.33, 8, null, '日')))).toEqual([16, 38, 60, 82, 105, 127, 149, 171]);
  });

  it('unions the runs of a mixed line and rescales them into 1.2 × the largest size', () => {
    const mixed = (a: string, b: string) => baselines([
      paragraph([{ text: 'H1', font: a, size: 100 }]),
      paragraph([{ text: 'H', font: a, size: 100 }, { text: '日g2', font: b, size: 100 }]),
      paragraph([{ text: 'H3', font: a, size: 100 }]),
    ]);
    expect(units(mixed('Arial', 'Meiryo'))).toEqual([135, 289, 468]);
    expect(units(baselines([
      paragraph([{ text: 'H1', font: 'Arial', size: 40 }]),
      paragraph([{ text: 'H', font: 'Arial', size: 40 }, { text: 'Hg2', font: 'Arial', size: 100 }]),
      paragraph([{ text: 'H3', font: 'Arial', size: 40 }]),
    ]))).toEqual([54, 202, 287]);
  });

  it('re-divides spaced lines with the shared DrawingML rule and rounds spcPts to whole points', () => {
    const pts = (val: number): Spacing => ({ type: 'pts', val });
    expect(units(baselines(lines('Arial', 100, 2, pts(114))))).toEqual([127, 285]);
    expect(units(baselines(lines('Arial', 100, 2, pts(126))))).toEqual([131, 306]);
    // H equal to the natural 120 pt keeps the natural split (MS Gothic).
    expect(units(baselines(lines('MS Gothic', 100, 2, pts(120), '日')))).toEqual([143, 310]);
    expect(units(baselines(lines('Meiryo', 100, 2, pts(144), '日')))).toEqual([143, 343]);
    expect(units(baselines(lines('Meiryo', 100, 2, { type: 'pct', val: 150000 }, '日')))).toEqual([180, 430]);
    // P-*: 40.5 pt spacing lays out as 41 pt, 48.33 pt as 48 pt.
    expect(units(baselines(lines('Arial', 40, 8, pts(40.5))))).toEqual([44, 101, 158, 215, 272, 329, 386, 443]);
    expect(units(baselines(lines('Arial', 40, 8, pts(48.33))))).toEqual([54, 121, 187, 254, 321, 387, 454, 521]);
  });

  it('applies edge spacing only with spcFirstLastPara and anchors the unrounded block', () => {
    const edge = [
      paragraph([{ text: 'H1', font: 'Arial', size: 100 }], { spaceBefore: 7200 }),
      paragraph([{ text: 'H2', font: 'Arial', size: 100 }]),
    ];
    expect(units(baselines(edge))).toEqual([135, 302]);
    expect(units(baselines(edge, { spcFirstLastPara: true }))).toEqual([235, 402]);
    // G-Arial-ctr: two 100 pt lines centred in a 420 pt box, tIns/bIns 3.6 pt.
    const ctr = baselines(lines('Arial', 100, 2), { verticalAnchor: 'ctr', tIns: 45720, bIns: 45720 }, 420);
    // The anchored block top is not on the unit grid; the export counts from it.
    const top = 3.6 + (420 - 7.2 - 2 * 120) / 2;
    expect(units(ctr.map((y) => y - top)).map((n) => n + Math.floor((top - 3.6) / U + 0.5)))
      .toEqual([255, 422]);
  });

  it('sizes a line by the run latin face even when it draws no glyph, never by an unused ea face', () => {
    const probe = (mid: Paragraph) => units(baselines([
      paragraph([{ text: 'Hg1', font: 'Arial', size: 100 }]), mid,
      paragraph([{ text: 'Hg3', font: 'Arial', size: 100 }]),
    ]));
    const run = (text: string, latin: string, ea: string) => {
      const p = paragraph([{ text, font: latin, size: 100 }]);
      (p.runs[0] as { fontFamilyEa: string }).fontFamilyEa = ea;
      return p;
    };
    // supplement-3 U-*: ideographs only, latin face unused but counted.
    expect(probe(run('日本語', 'Calibri', 'MS PGothic'))).toEqual([135, 299, 468]);
    expect(probe(run('日本語', 'Meiryo', 'MS PGothic'))).toEqual([135, 291, 468]);
    // supplement-2 R-*: Latin only, the unused Meiryo ea face does not count.
    expect(probe(run('Hg2', 'Arial', 'Meiryo'))).toEqual([135, 302, 468]);
  });

  it('keeps each run latin slot when adjacent runs share a drawn face (colour never moves a line)', () => {
    // Two ideograph runs drawn in Meiryo, latin slots Calibri and Gabriola.
    const pair = (secondColor: string) => {
      const p = paragraph([
        { text: '日本', font: 'Calibri', size: 60 },
        { text: '語学', font: 'Gabriola', size: 60 },
      ]);
      for (const r of p.runs) (r as { fontFamilyEa: string }).fontFamilyEa = 'Meiryo';
      (p.runs[1] as { color: string }).color = secondColor;
      return baselines([p, paragraph([{ text: 'Hg', font: 'Arial', size: 60 }])]);
    };
    const same = pair('000000');
    const recoloured = pair('FF0000');
    expect(same).toHaveLength(2);
    expect(recoloured[0]).toBeCloseTo(same[0], 9);
    expect(recoloured[1]).toBeCloseTo(same[1], 9);
    // Gabriola's slot contributes: the pair differs from the Calibri-only line.
    const calibriOnly = paragraph([{ text: '日本語学', font: 'Calibri', size: 60 }]);
    (calibriOnly.runs[0] as { fontFamilyEa: string }).fontFamilyEa = 'Meiryo';
    const single = baselines([calibriOnly, paragraph([{ text: 'Hg', font: 'Arial', size: 60 }])]);
    expect(Math.abs(single[0] - same[0])).toBeGreaterThan(0.5);
  });

  it('bounds the share cache while many distinct face names are queried', () => {
    for (let i = 0; i < SHARE_CACHE_LIMIT * 4; i++) {
      powerPointAscentShare(`Unresolved Face ${i}`, i % 2 === 0, i % 3 === 0);
      expect(powerPointShareCacheSize()).toBeLessThanOrEqual(SHARE_CACHE_LIMIT);
    }
    expect(powerPointShareCacheSize()).toBe(SHARE_CACHE_LIMIT);
    // Known faces still resolve after eviction.
    expect(powerPointAscentShare('Arial', false, false)).toBeCloseTo(1854 / 2288, 12);
  });

  it('anchors by the last line natural descent and applies normAutofit to the split', () => {
    const box = { tIns: 45720, bIns: 45720 };
    const rel = (ys: number[]) => ys.map((y) => (y - 3.6) / U);
    // Q-Arial-b-p150: two 100 pt lines at 150 %, bottom of a 420 pt box.
    const q = rel(baselines(lines('Arial', 100, 2, { type: 'pct', val: 150000 }), { ...box, verticalAnchor: 'b' }, 420));
    expect(Math.abs(q[0] - 292.23)).toBeLessThanOrEqual(0.52);
    expect(Math.abs(q[1] - 542.21)).toBeLessThanOrEqual(0.52);
    // N-Arial-90-10-t: normAutofit fontScale 90 %, lnSpcReduction 10 %.
    const n = baselines(lines('Arial', 40 * 0.9, 4), { ...box, autoFit: 'norm', lnSpcReduction: 0.1 }, 260);
    expect(units(n.map((y) => y - 3.6))).toEqual([43, 97, 151, 205]);
  });

  it('sizes a line by its known faces when one run has an unresolved face (#1689)', () => {
    // An unresolved face adds no face to its line; the line keeps the metric
    // model of its known faces (Arial's usWin share), not the 0.8 split.
    const ys = baselines([paragraph([{ text: 'H', font: 'Arial', size: 20 }, { text: 'g', font: 'Avenir', size: 20 }])]);
    expect(ys[0]).toBeCloseTo(20 * 1.2 * (1854 / (1854 + 434)), 5);
  });

  it('preserves a larger unknown run size without inventing an ascent share (#1689)', () => {
    for (const spaceLine of [null, { type: 'pct' as const, val: 150000 }]) {
      const ys = baselines([
        paragraph([{ text: 'A', font: 'Arial', size: 20 }], { spaceLine }),
        paragraph([{ text: 'B', font: 'Arial', size: 20 }, { text: 'D', font: 'Avenir', size: 72 }], { spaceLine }),
        paragraph([{ text: 'C', font: 'Arial', size: 20 }], { spaceLine }),
      ]);
      const spacing = spaceLine ? 1.5 : 1;
      const share = 1854 / 2288;
      const spacedShare = spaceLine ? 0.75 : share;
      expect(ys[0]).toBeCloseTo(24 * spacing * spacedShare, 9);
      expect(ys[1]).toBeCloseTo((24 + 86.4 * spacedShare) * spacing, 9);
      expect(ys[2]).toBeCloseTo((24 + 86.4 + 24 * spacedShare) * spacing, 9);
    }
  });

  it('keeps only a line with no known face on the ordinary model (#1689 line scope)', () => {
    // fallback.win.pdf: an unmodelled face moves only its own line. A body
    // whose middle line has no known face keeps the metric baselines of the
    // other lines exactly.
    const known = baselines([
      paragraph([{ text: 'A', font: 'Arial', size: 20 }]),
      paragraph([{ text: 'B', font: 'Arial', size: 20 }]),
      paragraph([{ text: 'C', font: 'Arial', size: 20 }]),
    ]);
    const mixed = baselines([
      paragraph([{ text: 'A', font: 'Arial', size: 20 }]),
      paragraph([{ text: 'B', font: 'Avenir', size: 20 }]),
      paragraph([{ text: 'C', font: 'Arial', size: 20 }]),
    ]);
    expect(mixed[0]).toBeCloseTo(known[0], 9);
    expect(mixed[2]).toBeCloseTo(known[2], 9);
    expect(mixed[1] - known[0] - (known[1] - known[0])).toBeCloseTo(20 * 1.2 * (0.8 - 1854 / 2288), 5);
  });
});

// #1619: PowerPoint's reference (Windows-style) PDF export of the compatLnSpc
// controls. Each pair differs only in bodyPr@compatLnSpc; values are the
// exported baselines in 1/100 in from the text-area top (tIns 0), or from
// the box top for ctr/b anchors, where the export counts from the unrounded
// anchored block and the continuous position must lie within half a unit.
describe('bodyPr compatLnSpc (#1619)', () => {
  const ea = (font: string, size: number, n: number, spaceLine: Spacing = null, text = '日g') =>
    lines(font, size, n, spaceLine, text);
  const off = { compatLnSpc: false } as Partial<TextBody>;
  const within = (ys: number[], exported: number[]) => {
    expect(ys).toHaveLength(exported.length);
    ys.forEach((y, i) => expect(Math.abs(y / U - exported[i])).toBeLessThanOrEqual(0.5));
  };

  it('retains the whole-body fallback for unknown glyph faces under compatLnSpc=0', () => {
    for (const middle of [
      [{ text: 'B', font: 'Avenir', size: 20 }],
      [{ text: 'B', font: 'Avenir', size: 20 }, { text: 'D', font: 'Arial', size: 20 }],
    ]) {
      const p = [paragraph([{ text: 'A', font: 'Arial', size: 20 }]),
        paragraph(middle), paragraph([{ text: 'C', font: 'Arial', size: 20 }])];
      const ys = baselines(p, off);
      [19.2, 43.2, 67.2].forEach((y, i) => expect(ys[i]).toBeCloseTo(y, 10));
    }
  });

  it('retains the whole-body fontAlgn fallback for an unknown-only line', () => {
    const p = [paragraph([{ text: 'A', font: 'Arial', size: 20 }]),
      paragraph([{ text: 'B', font: 'Avenir', size: 20 }], { fontAlgn: 't' }),
      paragraph([{ text: 'C', font: 'Arial', size: 20 }])];
    const ys = baselines(p);
    [19.2, 43.2, 67.2].forEach((y, i) => expect(ys[i]).toBeCloseTo(y, 10));
  });

  it('keeps the #1610 model for compatLnSpc="1" exactly as when it is omitted', () => {
    for (const [p, exported] of [
      [lines('Times New Roman', 100, 1, null), [134]],
      [ea('Meiryo', 54, 1), [64]],
      [lines('Gabriola', 100, 1, null), [136]],
    ] as const) {
      expect(units(baselines([...p], { compatLnSpc: true }))).toEqual(exported);
      expect(units(baselines([...p]))).toEqual(exported);
    }
  });

  it('switches an explicit compatLnSpc="0" to the Excel natural box per face', () => {
    // A-*: Times New Roman takes the macOS Supplemental hhea gap (1.150 em).
    expect(units(baselines(lines('Times New Roman', 100, 1, null), off))).toEqual([130]);
    expect(units(baselines(lines('Gabriola', 100, 1, null), off))).toEqual([192]);
    expect(units(baselines(ea('Meiryo', 54, 1), off))).toEqual([96]);
    expect(units(baselines(ea('MS Gothic', 100, 1), off))).toEqual([140]);
    expect(units(baselines(lines('Aptos', 60, 2, null), off))).toEqual([78, 180]);
  });

  it('spaces compatLnSpc="0" lines with the shared rule on the Excel box', () => {
    const pts = (val: number): Spacing => ({ type: 'pts', val });
    expect(units(baselines(lines('Calibri', 40, 2, { type: 'pct', val: 80000 }), off))).toEqual([41, 95]);
    expect(units(baselines(ea('Meiryo', 40, 2, { type: 'pct', val: 150000 }), off))).toEqual([112, 275]);
    expect(units(baselines(ea('Meiryo', 40, 2, pts(30)), off))).toEqual([21, 63]);
    expect(units(baselines(lines('Arial', 100, 2, pts(100)), off))).toEqual([109, 248]);
  });

  it('bases compatLnSpc="0" percentage paragraph spacing on the natural box', () => {
    const bef = (font: string, text: string, extra: Partial<Paragraph>) => [0, 1, 2].map((i) =>
      paragraph([{ text: `${text}${i}`, font, size: 40 }], { spaceLine: null, ...extra }));
    // D-mei-bef50pct: 50 % of Meiryo's 1.95 em natural line, not of 1.2 em.
    expect(units(baselines(bef('Meiryo', '日g', { spaceBeforePct: 50000 } as Partial<Paragraph>), off)))
      .toEqual([71, 234, 396]);
    expect(units(baselines(bef('Arial', 'Hg', { spaceBefore: 1800 }), { ...off, spcFirstLastPara: true })))
      .toEqual([77, 166, 255]);
  });

  it('unions mixed faces without rescaling and anchors by the natural descent', () => {
    const mixed = [1, 2].map((i) => paragraph([
      { text: `Hg${i}`, font: 'Arial', size: 24 }, { text: '日g', font: 'Meiryo', size: 60 },
    ], { spaceLine: null }));
    expect(units(baselines(mixed, off))).toEqual([107, 270]);
    within(baselines(ea('Meiryo', 40, 2), { ...off, verticalAnchor: 'b' }, 200), [132.09, 241.13]);
    within(baselines(lines('Arial', 40, 2, null), { ...off, verticalAnchor: 'b' }, 200), [202.02, 266.01]);
    within(baselines(ea('Meiryo', 40, 2, { type: 'pct', val: 80000 }), { ...off, verticalAnchor: 'ctr' }, 200),
      [104.52, 191.52]);
  });
});

// #1636: pPr@fontAlgn t / ctr / b and the a:br / endParaRPr marks, from
// PowerPoint's reference ("electronic distribution") PDF export of the
// controls. Values are the exported baselines of every run in 1/100 in from
// the text-area top, [compatLnSpc="0", compatLnSpc="1"]. The export rounds
// each continuous position to that unit, so every one must lie within half a
// unit (a few controls land exactly on a rounding tie).
describe('pPr fontAlgn and line-break marks (#1636)', () => {
  type Model = [compat0: number[], compat1: number[]];
  const runBaselines = (paragraphs: Paragraph[], compatLnSpc: boolean): number[] => {
    const body = {
      verticalAnchor: 't', paragraphs, defaultFontSize: 18, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none', compatLnSpc,
    } as TextBody;
    const { ctx, draws } = context();
    renderTextBody(ctx, body, 0, 0, 2000, 800, SCALE);
    return draws.filter((d) => d.text).map((d) => d.y / U);
  };
  const expectWithin = (ys: number[], exported: number[], tolerance = 0.5) => {
    expect(ys).toHaveLength(exported.length);
    ys.forEach((y, i) => expect(Math.abs(y - exported[i]), `run ${i}: ${y} vs ${exported[i]}`)
      .toBeLessThanOrEqual(tolerance + 1e-9));
  };
  const expectModels = (paragraphs: () => Paragraph[], [compat0, compat1]: Model, tolerance = 0.5) => {
    expectWithin(runBaselines(paragraphs(), false), compat0, tolerance);
    expectWithin(runBaselines(paragraphs(), true), compat1, tolerance);
  };
  const spacing = (tag: string): Spacing => tag.startsWith('pct')
    ? { type: 'pct', val: Number(tag.slice(3)) * 1000 }
    : { type: 'pts', val: Number(tag.slice(3)) };
  const ea = (font: string) => font === 'Meiryo' || font === 'Yu Gothic';

  // S-*: two one-line 40 pt paragraphs, around both models' natural height
  // (default 48 pt for both faces; compatLnSpc="0" 46 pt Arial, 78 pt Meiryo).
  const sweep: Array<[string, 't' | 'ctr' | 'b', string, ...Model]> = [
    ['Arial', 't', 'pct90', [47, 105], [45, 105]],
    ['Arial', 't', 'pct99', [52, 115], [49, 116]],
    ['Arial', 't', 'pct100', [52, 116], [50, 117]],
    ['Arial', 't', 'pct101', [53, 117], [51, 118]],
    ['Arial', 't', 'pct110', [58, 129], [57, 130]],
    ['Arial', 't', 'pts44', [50, 111], [46, 107]],
    ['Arial', 't', 'pts47', [53, 119], [49, 114]],
    ['Arial', 't', 'pts52', [60, 133], [56, 128]],
    ['Arial', 'ctr', 'pct90', [47, 105], [47, 107]],
    ['Arial', 'ctr', 'pct100', [52, 116], [52, 119]],
    ['Arial', 'ctr', 'pct101', [53, 117], [53, 120]],
    ['Arial', 'ctr', 'pct110', [58, 129], [59, 132]],
    ['Arial', 'ctr', 'pts46', [52, 116], [50, 114]],
    ['Arial', 'b', 'pct90', [46, 103], [48, 108]],
    ['Arial', 'b', 'pct100', [52, 116], [54, 121]],
    ['Arial', 'b', 'pct101', [42, 107], [45, 112]],
    ['Arial', 'b', 'pct110', [47, 117], [49, 122]],
    ['Arial', 'b', 'pts46', [42, 106], [52, 115]],
    ['Arial', 'b', 'pts52', [48, 120], [48, 120]],
    ['Meiryo', 't', 'pct90', [63, 160], [54, 114]],
    ['Meiryo', 't', 'pct100', [71, 179], [59, 126]],
    ['Meiryo', 't', 'pct101', [72, 182], [60, 127]],
    ['Meiryo', 't', 'pts44', [36, 97], [55, 116]],
    ['Meiryo', 't', 'pts80', [74, 185], [103, 215]],
    ['Meiryo', 'ctr', 'pct90', [63, 161], [45, 105]],
    ['Meiryo', 'ctr', 'pct100', [71, 180], [50, 117]],
    ['Meiryo', 'ctr', 'pct110', [82, 201], [57, 130]],
    ['Meiryo', 'ctr', 'pts76', [69, 175], [89, 195]],
    ['Meiryo', 'b', 'pct99', [71, 178], [44, 110]],
    ['Meiryo', 'b', 'pct100', [72, 180], [45, 111]],
    ['Meiryo', 'b', 'pct101', [64, 173], [39, 106]],
    ['Meiryo', 'b', 'pts44', [28, 89], [39, 100]],
    ['Meiryo', 'b', 'pts52', [36, 108], [42, 114]],
    ['Meiryo', 'b', 'pts80', [65, 176], [71, 182]],
  ];
  it.each(sweep)('%s 40 pt, fontAlgn %s, lnSpc %s', (font, fontAlgn, tag, compat0, compat1) => {
    const text = ea(font) ? '日' : 'H';
    expectModels(() => [1, 2].map((i) => paragraph([{ text: `${text}${i}`, font, size: 40 }],
      { spaceLine: spacing(tag), fontAlgn })), [compat0, compat1]);
  });

  it('aligns the descent midpoints of 3-4 sizes under fontAlgn b', () => {
    const line = (spec: Array<[string, number]>, spaceLine: Spacing) => [1, 2].map((i) => paragraph(
      spec.map(([font, size], k) => ({ text: (ea(font) ? '日' : 'H') + (k === 0 ? String(i) : ''), font, size })),
      { spaceLine, fontAlgn: 'b' }));
    const ari4: Array<[string, number]> = [['Arial', 20], ['Arial', 30], ['Arial', 45], ['Arial', 60]];
    const meiAri4: Array<[string, number]> = [['Meiryo', 20], ['Arial', 30], ['Meiryo', 45], ['Arial', 60]];
    const yug3: Array<[string, number]> = [['Yu Gothic', 20], ['Yu Gothic', 30], ['Yu Gothic', 45]];
    expectModels(() => line(ari4, null), [
      [84, 83, 80, 78, 180, 179, 176, 174], [87, 86, 83, 81, 187, 186, 183, 181]]);
    expectModels(() => line(ari4, spacing('pct80')), [
      [65, 64, 61, 59, 141, 140, 137, 135], [67, 66, 63, 61, 147, 146, 143, 141]]);
    expectModels(() => line(ari4, spacing('pct120')), [
      [83, 82, 79, 77, 198, 197, 194, 192], [87, 86, 83, 81, 207, 206, 203, 201]]);
    expectModels(() => line(meiAri4, null), [
      [92, 97, 80, 92, 214, 219, 202, 214], [83, 85, 75, 80, 183, 185, 175, 180]]);
    expectModels(() => line(meiAri4, spacing('pct120')), [
      [101, 106, 89, 101, 247, 252, 235, 247], [84, 86, 76, 81, 204, 206, 196, 201]]);
    expectModels(() => line(yug3, spacing('pct80')), [
      [61, 58, 53, 145, 142, 137], [47, 45, 42, 107, 105, 102]]);
  });

  // R-* / E-*: a:br and endParaRPr contribute their face at the size of the
  // run they follow; their own size (16 / 80 pt here) never matters.
  const mark = (font: string, size: number) => ({ fontSize: size, fontFamily: font });
  const withBreak = (spec: Array<[string, number]>, br: { fontSize: number; fontFamily: string }) => {
    const [first, second] = [1, 2].map((i) => paragraph(spec.map(([font, size], k) => ({
      text: (ea(font) ? '日' : 'H') + (k === 0 ? String(i) : ''), font, size,
    })), { spaceLine: null }));
    const last = spec[spec.length - 1];
    return [{
      ...first,
      runs: [...first.runs, { type: 'break', ...br }, ...second.runs],
      endRunProperties: mark(last[0], last[1]),
    } as unknown as Paragraph];
  };
  const withEnd = (spec: Array<[string, number]>, end: { fontSize: number; fontFamily: string }) =>
    [1, 2].map((i) => ({
      ...paragraph(spec.map(([font, size], k) => ({
        text: (ea(font) ? '日' : 'H') + (k === 0 ? String(i) : ''), font, size,
      })), { spaceLine: null }),
      endRunProperties: end,
    } as unknown as Paragraph));

  it('sizes an a:br face at the preceding run size', () => {
    expectModels(() => withBreak([['Meiryo', 24], ['Arial', 60]], mark('Meiryo', 16)),
      [[107, 107, 241, 241], [73, 73, 181, 181]]);
    expectModels(() => withBreak([['Meiryo', 24], ['Arial', 60]], mark('Meiryo', 80)),
      [[107, 107, 241, 241], [73, 73, 181, 181]]);
    expectModels(() => withBreak([['Arial', 24], ['Meiryo', 60]], mark('Arial', 16)),
      [[107, 107, 270, 270], [73, 73, 171, 171]]);
    expectModels(() => withBreak([['Arial', 60], ['Arial', 24]], mark('Arial', 80)),
      [[78, 78, 174, 174], [81, 81, 181, 181]]);
  });

  it('sizes an endParaRPr face at the last run size', () => {
    expectModels(() => withEnd([['Meiryo', 24], ['Arial', 60]], mark('Meiryo', 16)),
      [[107, 107, 270, 270], [73, 73, 173, 173]]);
    expectModels(() => withEnd([['Arial', 24], ['Meiryo', 60]], mark('Arial', 80)),
      [[107, 107, 270, 270], [73, 73, 173, 173]]);
    expectModels(() => withEnd([['Arial', 60], ['Arial', 24]], mark('Arial', 80)),
      [[78, 78, 174, 174], [81, 81, 181, 181]]);
  });

  // #1663: marks without a face of their own contribute the face they inherit
  // (here the Meiryo paragraph default) the same way; an omitted endParaRPr
  // contributes nothing. [compatLnSpc="0", compatLnSpc="1"] of the
  // endmark-inherited controls, Arial 60 + Arial 24 lines. With the Meiryo
  // 24 pt face the compatLnSpc="0" second baseline lands 0.004 unit past a
  // rounding tie (178.504 vs 179), exactly as the authored control and the
  // #1636 Arial 60 + Meiryo 24 lines do, hence the 0.51 tolerance.
  const TIE = 0.51;
  const inheriting = (paragraphs: Paragraph[]) => paragraphs.map((p) => ({ ...p, defFontFamily: 'Meiryo' } as Paragraph));
  const ari = (i: number) => [{ text: `H${i}`, font: 'Arial', size: 60 }, { text: 'H', font: 'Arial', size: 24 }];
  it('adds an inherited endParaRPr face at the last run size', () => {
    const withMark = (end: object | undefined) => () => inheriting([1, 2].map((i) => ({
      ...paragraph(ari(i), { spaceLine: null }), ...(end ? { endRunProperties: end } : {}),
    } as unknown as Paragraph)));
    // lang only or size only: the parser resolves the inherited face.
    for (const size of [16, 40, 80]) {
      expectModels(withMark({ fontSize: size, fontFamily: 'Meiryo' }), [[78, 78, 179, 179], [81, 81, 181, 181]], TIE);
    }
    expectModels(withMark(undefined), [[78, 78, 174, 174], [81, 81, 181, 181]]);
    expectModels(() => withMark({ fontSize: 16, fontFamily: 'Meiryo' })().map((p) => ({ ...p, fontAlgn: 'b' } as Paragraph)),
      [[78, 83, 176, 181], [81, 86, 181, 186]]);
  });

  it('adds an inherited a:br face at the preceding run size', () => {
    const withBreakMark = (br: object) => () => {
      const [first, second] = [1, 2].map((i) => paragraph(ari(i), { spaceLine: null }));
      return inheriting([{
        ...first, runs: [...first.runs, { type: 'break', ...br }, ...second.runs],
        endRunProperties: mark('Arial', 24),
      } as unknown as Paragraph]);
    };
    // a:br without rPr, and a size-only rPr at 16 / 80 pt.
    for (const br of [{}, { fontSize: 16 }, { fontSize: 80 }]) {
      expectModels(withBreakMark(br), [[78, 78, 179, 179], [81, 81, 181, 181]], TIE);
    }
  });

  it('handles lines with 10^5 metric runs without spreading them into Math.max', () => {
    // Alternating Latin / CJK characters split one run into a segment per
    // character, each adding its latin face as a second metric entry.
    const text = 'a日'.repeat(50_000);
    for (const fontAlgn of ['t', 'ctr', 'b'] as const) {
      for (const compat of [false, true]) {
        const p = paragraph([{ text, font: 'Arial', size: 20 }], { spaceLine: null, fontAlgn });
        p.runs = p.runs.map((r) => ({ ...r, fontFamilyEa: 'Meiryo' }));
        expect(() => runBaselines([p], compat)).not.toThrow();
      }
    }
    // Six layouts of a 10^5-character line. The cost is linear and the same
    // as the baseline (non-fontAlgn) path, ~1.2-1.7 s each locally (10^4: 0.3 s,
    // 2 x 10^5: 3.5 s), so the default 5 s budget is too tight on CI. The size
    // stays: ~150 000 metric entries are needed to exceed V8's spread-argument
    // limit (~125 000 here) that this test guards.
  }, 60_000);
});
