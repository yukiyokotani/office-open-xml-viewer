import { describe, expect, it, vi } from 'vitest';
import { PT_TO_PX, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { bindXlsxOfficeFontRoutes, drawShapeText } from './renderer.js';
import { xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import {
  canvasShapeFontBoxProbe, excelShapeSpacing, shapeRunLineRatios, type ShapeFontBoxProbe,
} from './shape-office-line.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun, Worksheet } from './types.js';

// Expectations below come from Excel for Mac 16.113.2 PDF exports of the
// issue #1604 controls (baselines relative to the shape top, top anchor,
// tIns 3.6 pt). Arial: ascent 0.938 em (usWinAscent + lineGap), descent
// 0.212 em. Meiryo: 1.95 em box, ascent 1.285 em.

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;

function run(text: string, fontFace: string, size: number, bold = false): TextRun {
  return { type: 'text', text, fontFace, fontFaceEa: fontFace, bold, italic: false, size };
}

function para(runs: TextRun[], extra: Partial<ShapeParagraph> = {}): ShapeParagraph {
  return { align: 'l', runs, ...extra };
}

function body(paragraphs: ShapeParagraph[], anchor: 't' | 'b' = 't'): ShapeText {
  return {
    anchor, wrap: 'none', autoFit: 'none',
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720, paragraphs,
  };
}

function context(): { ctx: CanvasRenderingContext2D; draws: Array<{ text: string; y: number; font: string; baseline: string }> } {
  let font = '11px sans-serif';
  const draws: Array<{ text: string; y: number; font: string; baseline: string }> = [];
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    measureText() { return { width: 40, actualBoundingBoxAscent: 20 }; },
    fillText(text: string, _x: number, y: number) {
      draws.push({ text, y, font, baseline: (ctx as { textBaseline: string }).textBaseline });
    },
    fillStyle: '#000', textBaseline: 'alphabetic',
  } as unknown as CanvasRenderingContext2D;
  return { ctx, draws };
}

function route(family: string, weight: 400 | 700 = 400): OfficeFontFallbackRoute {
  const alias = `__exact_${family.toLowerCase().replace(/\s+/gu, '_')}_${weight}`;
  return {
    requestedFamily: family, family: alias, source: 'local',
    resourceIdentity: `office-local:local("${family}")`,
    weight, style: 'normal', metric: { family: alias, synthesized: false },
  };
}

const ROUTES: Record<string, OfficeFontFallbackRoute> = {
  arial: route('Arial'),
  meiryo: route('Meiryo'),
  'ms gothic': route('MS Gothic'),
  'yu gothic': route('Yu Gothic'),
  'times new roman': route('Times New Roman'),
};

/** Office PDF positions carry up to 0.16 pt of export rounding in these controls. */
function expectOffice(actual: number, office: number): void {
  expect(Math.abs(actual - office)).toBeLessThanOrEqual(0.16);
}

/** Paint and return each line's baseline in pt below the shape top. */
function baselinesPt(text: ShapeText, routes: Record<string, OfficeFontFallbackRoute> | null = ROUTES): number[] {
  const { ctx, draws } = context();
  if (routes) bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, routes);
  drawShapeText(ctx, text, 400, 900, 1);
  return [...new Set(draws.map((draw) => draw.y))].map((y) => y / PT_TO_PX);
}

describe('Excel shape-text line box from font metrics (#1604)', () => {
  it('projects the Office line box from usWin metrics and the Far East class', () => {
    const sum = (r: ReturnType<typeof shapeRunLineRatios>) => (r ? r.ascentRatio + r.descentRatio : NaN);
    // Excel: Arial pitch 1.150 em (usWin box + hhea lineGap 67).
    expect(sum(shapeRunLineRatios(run('Hg', 'Arial', 24), ROUTES.arial))).toBeCloseTo(1.1499, 4);
    // Excel: Yu Gothic pitch 1.673 em = 1.3 × usWin box (the hhea box would give 1.433).
    expect(sum(shapeRunLineRatios(run('日g', 'Yu Gothic', 24), ROUTES['yu gothic']))).toBeCloseTo(1.6732, 3);
    // Earlier single-line controls: Meiryo UI Bold measured 1.646–1.649 em.
    expect(sum(shapeRunLineRatios(run('予算', 'Meiryo UI', 25, true), route('Meiryo UI', 700)))).toBeCloseTo(1.651, 3);
    // Supplement L-*: bundled Baskerville Old Face follows usWin (1.141 em),
    // bundled Gabriola its USE_TYPO_METRICS typo box (1.700 em), and the macOS
    // system Palatino and Helvetica their hhea box (1.100 and 1.000 em).
    expect(sum(shapeRunLineRatios(run('Hg', 'Baskerville Old Face', 24), route('Baskerville Old Face')))).toBeCloseTo(1.1406, 4);
    expect(sum(shapeRunLineRatios(run('Hg', 'Gabriola', 24), route('Gabriola')))).toBeCloseTo(1.7, 3);
    // This installed OS/2 v3 resource declares USE_TYPO_METRICS. The shared
    // catalogue must pass its declared typo+gap box through XLSX as well.
    expect(sum(shapeRunLineRatios(run('■', 'Cambria Math', 24), route('Cambria Math'))))
      .toBeCloseTo(2401 / 2048, 12);
    expect(sum(shapeRunLineRatios(run('Hg', 'Palatino', 24), route('Palatino')))).toBeCloseTo(1.1001, 4);
    expect(sum(shapeRunLineRatios(run('Hg', 'Helvetica', 24), route('Helvetica')))).toBeCloseTo(1.0, 4);
    // Unverified resources and distinct East Asian faces stay undefined.
    expect(shapeRunLineRatios(run('Hg', 'Arial', 24), { ...ROUTES.arial, resourceIdentity: 'injected:x' })).toBeUndefined();
    expect(shapeRunLineRatios({ ...run('Hg', 'Arial', 24), fontFaceEa: 'Meiryo' }, ROUTES.arial)).toBeUndefined();
  });

  it('resolves disagreeing catalog profiles only from the loaded copy\'s font box', () => {
    const sum = (r: ReturnType<typeof shapeRunLineRatios>) => (r ? r.ascentRatio + r.descentRatio : NaN);
    const box = (ascent: number, descent: number, upm = 2048): ShapeFontBoxProbe =>
      () => ({ ascent: ascent / upm, descent: descent / upm });
    const rockwell = run('Hg', 'Rockwell', 24);
    // Rockwell: macOS copy hhea 1391/-657/410 (system rule 1.200 em), Office
    // copy 1937/-468 (usWin rule 1.174 em). A name-only route proves neither.
    expect(shapeRunLineRatios(rockwell, route('Rockwell'))).toBeUndefined();
    // Unambiguous match: the loaded box is the macOS copy's hhea box, within
    // the 0.002 em probe tolerance, so Excel's system-copy rule applies.
    expect(sum(shapeRunLineRatios(rockwell, route('Rockwell'), box(1391 + 3, 657 - 3)))).toBeCloseTo(2458 / 2048, 4);
    // Beyond the tolerance, nothing matches.
    expect(shapeRunLineRatios(rockwell, route('Rockwell'), box(1391 + 6, 657))).toBeUndefined();
    // The Office copy is loaded, but Excel prefers a system copy the viewer may
    // also have, so the Office box is not proof: declined.
    expect(shapeRunLineRatios(rockwell, route('Rockwell'), box(1937, 468))).toBeUndefined();
    // Ambiguous: both Times New Roman copies report 1825/443 (they differ only
    // in hhea lineGap, 87 vs 0, which Canvas does not expose): declined.
    expect(shapeRunLineRatios(run('Hg', 'Times New Roman', 24), ROUTES['times new roman'], box(1825, 443)))
      .toBeUndefined();
    // Two macOS Helvetica Neue Bold faces disagree; the probe picks one.
    expect(sum(shapeRunLineRatios(run('Hg', 'Helvetica Neue', 24, true), route('Helvetica Neue', 700),
      box(961, 221, 1000)))).toBeCloseTo((961 + 28 + 221) / 1000, 4);
    // A single-profile family needs no probe, and a probe does not override it.
    const probe = vi.fn(box(0, 0));
    expect(sum(shapeRunLineRatios(run('Hg', 'Baskerville Old Face', 24), route('Baskerville Old Face'), probe)))
      .toBeCloseTo(1.1406, 4);
    expect(probe).not.toHaveBeenCalled();
  });

  it('probes the route face once and restores the context font', () => {
    let font = '10px serif';
    const measureText = vi.fn(() => ({ fontBoundingBoxAscent: 600, fontBoundingBoxDescent: 400 }));
    const ctx = { get font() { return font; }, set font(value: string) { font = value; }, measureText };
    const probe = canvasShapeFontBoxProbe(ctx as unknown as CanvasRenderingContext2D);
    const target = route('Rockwell');
    expect(probe(target)).toEqual({ ascent: 0.6, descent: 0.4 });
    expect(probe(target)).toEqual({ ascent: 0.6, descent: 0.4 });
    expect(measureText).toHaveBeenCalledTimes(1);
    expect(font).toBe('10px serif');
  });

  it('steps same-size lines by the font line box and places the first baseline at the ascent', () => {
    // Excel S-Arial-24: 26.18, 53.78, 81.38 pt. S-Meiryo-24: 34.46, 81.26, 128.09 pt.
    const arial = baselinesPt(body([para([run('A1', 'Arial', 24)]), para([run('A2', 'Arial', 24)]),
      para([run('A3', 'Arial', 24)])]));
    expect(arial.map((v) => +v.toFixed(1))).toEqual([26.1, 53.7, 81.3]);
    const meiryo = baselinesPt(body([para([run('M1', 'Meiryo', 24)]), para([run('M2', 'Meiryo', 24)])]));
    expect(meiryo[0]).toBeCloseTo(34.44, 1);
    expect(meiryo[1] - meiryo[0]).toBeCloseTo(46.8, 1);
  });

  it('unions run ascents and descents on one shared baseline', () => {
    // Excel X-Arial-Meiryo: pitches 35.91 and 38.52 pt around a mixed line.
    const { ctx, draws } = context();
    bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, ROUTES);
    drawShapeText(ctx, body([
      para([run('A1', 'Arial', 24)]),
      para([run('Hx', 'Arial', 24), run('日g', 'Meiryo', 24)]),
      para([run('A3', 'Arial', 24)]),
    ]), 400, 900, 1);
    const y = (text: string) => draws.find((draw) => draw.text === text)!.y / PT_TO_PX;
    expect(y('Hx')).toBe(y('日g'));
    expect(y('Hx') - y('A1')).toBeCloseTo(35.91, 1);
    expect(y('A3') - y('Hx')).toBeCloseTo(38.52, 1);
    // Excel M-Arial-seq: 11→40→11 pt lines pitch 39.84 then 18.86 pt.
    const seq = baselinesPt(body([11, 40, 11].map((size, i) => para([run(`S${i}`, 'Arial', size)]))));
    expectOffice(seq[1] - seq[0], 39.84);
    expectOffice(seq[2] - seq[1], 18.86);
  });

  it('re-divides the line box for spcPct and spcPts', () => {
    const three = (font: string, size: number, spaceLine: ShapeParagraph['spaceLine']) =>
      body([1, 2, 3].map((n) => para([run(`L${n}`, font, size)], { spaceLine })));
    // Excel P-Arial-40-150: first 55.36, pitch 69.0. P-Arial-40-80: first 31.93, pitch 36.84.
    const up = baselinesPt(three('Arial', 40, { type: 'pct', val: 150000 }));
    expect(up[0]).toBeCloseTo(55.35, 1);
    expect(up[1] - up[0]).toBeCloseTo(69.0, 1);
    const down = baselinesPt(three('Arial', 40, { type: 'pct', val: 80000 }));
    expect(down[0]).toBeCloseTo(31.92, 1);
    // Excel P-Meiryo-40-150 (ascent below 75 % of the box): first 84.25 pt.
    expect(baselinesPt(three('Meiryo', 40, { type: 'pct', val: 150000 }))[0]).toBeCloseTo(84.25, 1);
    // Excel T-Arial-24-18 and T-Meiryo-24-18: first 17.14 and 12.82, pitch 18.
    const arialPts = baselinesPt(three('Arial', 24, { type: 'pts', val: 18 }));
    expect(arialPts[0]).toBeCloseTo(17.1, 1);
    expect(arialPts[1] - arialPts[0]).toBeCloseTo(18, 5);
    expect(baselinesPt(three('Meiryo', 24, { type: 'pts', val: 18 }))[0]).toBeCloseTo(12.84, 1);
  });

  it('keeps the natural descent under spcPts until a quarter of the line exceeds it', () => {
    const three = (size: number, pts: number) =>
      body([1, 2, 3].map((n) => para([run(`M${n}`, 'Meiryo', size)], { spaceLine: { type: 'pts', val: pts } })));
    // Excel T-mei-mix-30 / E-mei-14-30: 14 pt Meiryo at 30 pt keeps its 9.31 pt
    // descent (first baseline 24.26–24.52 pt); at 40 pt the descent grows to
    // 0.25·H + k (first baseline 31.6 pt).
    expectOffice(baselinesPt(three(14, 30))[0], 24.3);
    expect(Math.abs(baselinesPt(three(14, 40))[0] - 31.6)).toBeLessThanOrEqual(0.6);
    // Supplement E-mei-24-56: first baseline 43.48 pt (0.24 pt export grid).
    expect(Math.abs(baselinesPt(three(24, 56))[0] - 43.48)).toBeLessThanOrEqual(0.3);
  });

  it('rounds spcPts to whole points before the descent step at 4d', () => {
    const three = (font: string, size: number, pts: number) =>
      body([1, 2, 3].map((n) => para([run(`P${n}`, font, size)], { spaceLine: { type: 'pts', val: pts } })));
    const first = (font: string, size: number, pts: number) => baselinesPt(three(font, size, pts))[0];
    // Excel boundary controls, first baseline of the spcPts line (0.24 pt export
    // grid, up to ~0.35 pt off the rule here). Unrounded, the rule would put
    // Meiryo 14 pt at 37.24 2.5 pt higher, Meiryo 24 pt at 63.5 3.5 pt lower
    // and Yu Gothic 14 pt at 27.5 0.5 pt lower.
    const near = (actual: number, office: number) => expect(Math.abs(actual - office)).toBeLessThanOrEqual(0.4);
    near(first('Meiryo', 14, 37.24), 31.6);   // → 37: keeps d (4d = 37.24)
    near(first('Meiryo', 14, 37.35), 31.48);
    near(first('Meiryo', 14, 37.5), 29.68);   // → 38: descent 0.25·H + k
    near(first('Meiryo', 24, 63.25), 50.56);  // → 63 (4d = 63.83)
    near(first('Meiryo', 24, 63.5), 47.68);   // → 64
    near(first('Yu Gothic', 14, 27.25), 23.56); // → 27 (4d = 27.74)
    near(first('Yu Gothic', 14, 27.5), 23.68);  // → 28
  });

  it('rounds spcPct line spacing to whole percent before the descent step', () => {
    const first = (size: number, pct: number) => baselinesPt(body([1, 2, 3].map((n) =>
      para([run(`P${n}`, 'Meiryo', size)], { spaceLine: { type: 'pct', val: pct * 1000 } }))))[0];
    const near = (actual: number, office: number) => expect(Math.abs(actual - office)).toBeLessThanOrEqual(0.5);
    // Excel rounding controls, first baseline. 136.4 % → 136 %: H stays below
    // 4d and keeps d. Unrounded, H passes 4d by < 0.001 pt and the line jumps
    // 2.4 pt (14 pt) or 4.3 pt (24 pt).
    near(first(14, 136.4), 31.48);
    near(first(14, 137), 29.56);  // stepped: descent 0.25·H + k
    near(first(24, 136.4), 51.52);
    near(first(24, 125), 46.6);   // L < H < 4d keeps d
  });

  it('rounds spcBef and spcAft one by one to whole points or whole percent', () => {
    // Six natural Meiryo 14 pt paragraphs with the spacing on every gap; the
    // mean pitch over five gaps carries under 0.1 pt of export noise.
    const pitch = (bef?: ShapeParagraph['spaceBefore'], aft?: ShapeParagraph['spaceAfter']) => {
      const lines = baselinesPt(body([0, 1, 2, 3, 4, 5].map((i) => para([run(`S${i}`, 'Meiryo', 14)], {
        ...(bef && i > 0 ? { spaceBefore: bef } : {}), ...(aft && i < 5 ? { spaceAfter: aft } : {}),
      }))));
      return (lines[5] - lines[0]) / 5;
    };
    const near = (actual: number, office: number) => expect(Math.abs(actual - office)).toBeLessThanOrEqual(0.15);
    near(pitch({ type: 'pts', val: 3.4 }), 30.19);           // → 3 pt (raw 3.4 → 30.7)
    near(pitch(undefined, { type: 'pts', val: 3.5 }), 31.23); // → 4 pt
    // Both 3.3 pt on each gap: 3 + 3 = 6 pt, not round(6.6) = 7.
    near(pitch({ type: 'pts', val: 3.3 }, { type: 'pts', val: 3.3 }), 33.22);
    near(pitch({ type: 'pct', val: 10500 }), 30.43);          // → 11 % of 27.3 pt (raw → 30.17)
    near(pitch(undefined, { type: 'pct', val: 13500 }), 31.2); // → 14 % (raw → 30.99)
    expect(excelShapeSpacing({ type: 'pts', val: 36.49 })).toEqual({ type: 'pts', val: 36 });
    expect(excelShapeSpacing({ type: 'pts', val: 36.5 })).toEqual({ type: 'pts', val: 37 });
    expect(excelShapeSpacing({ type: 'pct', val: 136400 })).toEqual({ type: 'pct', val: 136000 });
    expect(excelShapeSpacing({ type: 'pct', val: 12500 })).toEqual({ type: 'pct', val: 13000 });
  });

  it('takes spcPct from the natural line even when lnSpc is not 100 %', () => {
    // Supplement Q-Arial-l150-bef50: lnSpc 150 % and spcBef 50 % → 55.2 pt
    // (41.4 + 13.8 of the 27.6 pt natural line), not 41.4 + 20.7.
    const lines = baselinesPt(body([1, 2].map((n) => para([run(`Q${n}`, 'Arial', 24)], {
      spaceLine: { type: 'pct', val: 150000 },
      ...(n === 2 ? { spaceBefore: { type: 'pct', val: 50000 } } : {}),
    }))));
    expect(Math.abs(lines[1] - lines[0] - 55.2)).toBeLessThanOrEqual(0.3);
  });

  it('adds spcAft and spcBef between paragraphs but not before the first', () => {
    const lines = (p1: Partial<ShapeParagraph>, p2: Partial<ShapeParagraph>) => baselinesPt(body([
      para([run('P1', 'Arial', 24)], p1), para([run('P2', 'Arial', 24)], p2),
    ]));
    // Excel B-Arial-both: spcAft 12 + spcBef 18 pt → pitch 57.6 (sum).
    const both = lines({ spaceAfter: { type: 'pts', val: 12 } }, { spaceBefore: { type: 'pts', val: 18 } });
    expect(both[1] - both[0]).toBeCloseTo(57.6, 1);
    // Excel B-Arial-bef-pct50: 50 % of the 27.6 pt natural line → pitch 41.42.
    const pct = lines({}, { spaceBefore: { type: 'pct', val: 50000 } });
    expect(pct[1] - pct[0]).toBeCloseTo(41.4, 1);
    // Excel B-Arial-first-bef: spcBef on the first paragraph leaves 26.23 pt.
    expect(lines({ spaceBefore: { type: 'pts', val: 24 } }, {})[0]).toBeCloseTo(26.1, 1);
  });

  it('anchors the metric line box at the bottom of the text rectangle', () => {
    // Earlier Excel control: one 25 pt Meiryo UI Bold line in a 98 px box with
    // zero insets; top and bottom anchors exposed a 1.646–1.649 em line box.
    const text = { ...body([para([run('予算', 'Meiryo UI', 25, true)])], 'b'),
      lIns: 0, rIns: 0, tIns: 0, bIns: 0 };
    const { ctx, draws } = context();
    bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, { 'meiryo ui:700:normal': route('Meiryo UI', 700) });
    drawShapeText(ctx, text, 230, 98, 1);
    const ratios = shapeRunLineRatios(run('予算', 'Meiryo UI', 25, true), route('Meiryo UI', 700))!;
    expect(draws[0].y).toBeCloseTo(98 - 25 * PT_TO_PX * ratios.descentRatio, 5);
    expect(draws[0].font).toContain('__exact_meiryo_ui_700');
  });

  it('keeps the ordinary 1.2 em box when any run lacks a verified face', () => {
    const plain = baselinesPt(body([para([run('A1', 'Arial', 20)]), para([run('A2', 'Arial', 20)])]), null);
    expect(plain[1] - plain[0]).toBeCloseTo(24, 5);
    const mixed = baselinesPt(body([para([run('A1', 'Arial', 20)]), para([run('G2', 'Gabriola', 20)])]));
    expect(mixed[1] - mixed[0]).toBeCloseTo(24, 5);
    // Paragraph spacing still applies to the ordinary box.
    const spaced = baselinesPt(body([
      para([run('A1', 'Arial', 20)], { spaceAfter: { type: 'pts', val: 10 } }), para([run('A2', 'Arial', 20)]),
    ]), null);
    expect(spaced[1] - spaced[0]).toBeCloseTo(34, 5);
  });

  it('puts runs of another size on the largest run baseline in the ordinary box', () => {
    const { ctx, draws } = context();
    drawShapeText(ctx, body([para([run('ab', 'Arial', 11), run('CD', 'Arial', 40)])]), 400, 900, 1);
    expect(draws).toHaveLength(2);
    expect(draws[0].y).toBe(draws[1].y);
    expect(draws.every((draw) => draw.baseline === 'alphabetic')).toBe(true);
  });

  it('discovers every catalogued shape run face separately from cell fonts', () => {
    const ws = { rows: [{ cells: [{ value: { type: 'text', text: 'x', runs: [{
      text: 'x', font: { name: 'Calibri', bold: true, italic: false },
    }] } }] }], shapeGroups: [{ shapes: [
      { text: body([para([run('x', 'Meiryo UI', 25, true)]), para([run('y', 'Arial', 11)])]) },
      { text: body([para([{ ...run('z', 'Arial', 11), fontFaceEa: 'Meiryo' }])]) },
    ] }] } as unknown as Worksheet;
    expect(xlsxWorksheetOfficeFontRequests(ws)).toEqual([
      { family: 'Calibri', weight: 700, style: 'normal' },
      { family: 'Meiryo UI', weight: 700, style: 'normal' },
      { family: 'Arial', weight: 400, style: 'normal' },
    ]);
  });
});
