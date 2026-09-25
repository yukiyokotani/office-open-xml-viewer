/**
 * Producer-level pins for the per-line docGrid grid-count height
 * (`LayoutLine.gridCountSingle`) computed by {@link layoutLines} — the value
 * that feeds §17.6.5 East Asian cell counting in {@link lineBoxHeight}.
 *
 * The rule (see addToLine): the max over segments of each run's Word-faithful
 * single-line height, where a tabled EA run counts from its DESIGN height, an
 * untabled EA run from the Word FE 1.3em fallback, and a Latin run does NOT contribute
 * (it is not cell-rounded). This is what keeps a substituted face's over-tall
 * or under-tall box from changing the cell count (sample-9/sample-52).
 *
 * A linear stub context makes the metrics analytic: measured box = fontSize
 * (0.8 + 0.2 em); Yu Mincho's tabled EA design height = 1.43267 × em.
 */
import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import {
  layoutLines,
  lineBoxHeight,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (t: string) => ({
      width: [...t].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function integerRoundedCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (t: string) => ({
      width: [...t].length * px() * 0.5,
      // Deliberately reproduce Chrome's integer-rounded font boxes. At 20px
      // the substituted box lands exactly on the grid-pitch boundary; cell
      // allocation must come from the run's design count height instead.
      fontBoundingBoxAscent: Math.round(px() * 0.9),
      fontBoundingBoxDescent: Math.round(px() * 0.1),
      actualBoundingBoxAscent: Math.round(px() * 0.9),
      actualBoundingBoxDescent: Math.round(px() * 0.1),
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function seg(
  text: string,
  fontFamily: string,
  fontSize: number,
  eastAsianLineHeightRatio?: number,
): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily, vertAlign: null, measuredWidth: 0,
    resolvedEastAsianLineHeightRatio: eastAsianLineHeightRatio,
  };
}

function layout(segs: LayoutTextSeg[], width = 400) {
  return layoutLines(
    linearCtx(), segs.map((s) => ({ ...s })) as LayoutSeg[], width, 0, 1,
    undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, undefined, 36, width, false, false, false,
  );
}

// Yu Mincho tabled EA design single-line height per pt: 1.3 × hhea box = 1.43267 em.
const YU = (2257 * 1.3) / 2048;

describe('layoutLines — per-line gridCountSingle (§17.6.5 docGrid cell height)', () => {
  it('a resolved EA line counts from its resource DESIGN height, not the measured box', () => {
    const [line] = layout([seg('あいうえお', 'Arbitrary Resolved EA', 12, YU)]);
    expect(line.gridCountSingle).toBeCloseTo(12 * YU, 4);
    expect(line.gridCountSingle).toBeGreaterThan(12); // above the raw box
  });

  it('EXCLUDES a Latin run on a mixed CJK+Latin line (sample-52 shape)', () => {
    // sample-52's lines mix ASCII and CJK ("TABLE 1 = …（…）"). The ASCII run in
    // Yu Mincho has no design height (eaOnly ⇒ intendedSingle 0), and its
    // SUBSTITUTED Canvas box overshoots the real Latin height — Word measured
    // sample-52's exact table at x0=417 == the CJK-design (1-cell) layout, so
    // the ASCII box must NOT drive the count. This test uses a deliberately
    // TALL 20pt Latin run (box 20px > the 12pt CJK design 17.19px) to
    // DISCRIMINATE the exclusion: were it counted the result would be 20; the
    // §17.6.5 "Latin is not cell-rounded" rule yields the CJK design alone.
    // (The tall-Latin mixed line has no Word ground truth of its own — this
    // characterizes the deliberate rule, consistent with the Latin-only case.)
    const [line] = layout([
      seg('TABLE 1 = ', 'Arbitrary Resolved EA', 20, YU),
      seg('固定幅（両軸拘束）。', 'Arbitrary Resolved EA', 12, YU),
    ]);
    expect(line.gridCountSingle).toBeCloseTo(12 * YU, 4);
    expect(line.gridCountSingle).toBeLessThan(20); // the 20pt Latin box did NOT count
  });

  it('counts a genuinely TALL untabled EA run from the Word FE design fallback', () => {
    // Small 12pt Yu Mincho (design 17.19) + a large 24pt untabled CJK face.
    // Word's untabled FE fallback is 1.3em, so the latter contributes 31.2px;
    // its substituted Canvas box (24px in this stub) is irrelevant.
    const [line] = layout([
      seg('小', 'Arbitrary Resolved EA', 12, YU),
      seg('大きい', 'Unresolved EA', 24),
    ]);
    expect(line.gridCountSingle).toBeCloseTo(24 * 1.3, 4);
  });

  it('uses 1.3em for an untabled 20pt EA run whose measured box is exactly one em', () => {
    const run = seg('物理', 'ＭＳ 明朝', 20);
    const [line] = layoutLines(
      integerRoundedCtx(), [{ ...run }] as LayoutSeg[], 400, 0, 1,
      undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, undefined, 36, 400,
      false, false, false,
    );

    expect(line.ascent + line.descent).toBe(20);
    expect(line.gridCountSingle).toBeCloseTo(26, 8);
  });

  it('keeps the untabled EA grid count scale-linear at scale 2.3529', () => {
    const paintScale = 2.3529;
    const run = seg('物理', 'ＭＳ 明朝', 20);
    const ctx = integerRoundedCtx();
    const [painted] = layoutLines(
      ctx, [{ ...run }] as LayoutSeg[], 400 * paintScale, 0, paintScale,
      undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, undefined, 36, 400,
      false, false, false,
    );

    expect(painted.ascent + painted.descent).toBe(47);
    expect(painted.gridCountSingle).toBeCloseTo(26 * paintScale, 8);
  });

  it('keeps paginator and paint advances equal across integer-rounded docGrid boundaries', () => {
    const paintScale = 2.3529;
    const cases = [
      { fontSize: 20, linePitchPt: 20, expectedAdvance: 40 },
      { fontSize: 16, linePitchPt: 18, expectedAdvance: 36 },
    ];

    for (const { fontSize, linePitchPt, expectedAdvance } of cases) {
      const ctx = integerRoundedCtx();
      const [measuredLine] = layoutLines(
        ctx, [seg('物理', 'ＭＳ 明朝', fontSize)] as LayoutSeg[], 400, 0, 1,
        undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, undefined, 36, 400,
        false, false, false,
      );
      const [paintedLine] = layoutLines(
        integerRoundedCtx(),
        [seg('物理', 'ＭＳ 明朝', fontSize)] as LayoutSeg[],
        400 * paintScale,
        0,
        paintScale,
        undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, undefined, 36, 400,
        false, false, false,
      );
      const grid = { type: 'lines' as const, linePitchPt };

      const measureAdvance = lineBoxHeight(
        null, measuredLine.ascent, measuredLine.descent, 1, grid, false,
        measuredLine.intendedSingle, measuredLine.eastAsian ?? false,
        measuredLine.gridCountSingle,
      );
      const paintAdvance = lineBoxHeight(
        null, paintedLine.ascent, paintedLine.descent, paintScale, grid, false,
        paintedLine.intendedSingle, paintedLine.eastAsian ?? false,
        paintedLine.gridCountSingle,
      );

      expect(measureAdvance).toBe(expectedAdvance);
      expect(paintAdvance).toBeCloseTo(measureAdvance * paintScale, 8);
    }
  });

  it('a pure-Latin line is NOT cell-rounded — falls back to the box (unused off the EA path)', () => {
    // No EA segment contributes, so gridCountSingle falls back to the line box
    // (== old `natural`). The line is non-EA, so lineBoxHeight never consults it.
    const [line] = layout([seg('hello world', 'Times New Roman', 12)]);
    expect(line.eastAsian ?? false).toBe(false);
    expect(line.gridCountSingle).toBeCloseTo(line.ascent + line.descent, 4);
  });
});
