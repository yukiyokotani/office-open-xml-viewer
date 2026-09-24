import { describe, it, expect } from 'vitest';
import {
  docGridLineCells,
  eastAsianGridCountSinglePx,
  isGridLineRule,
  lineBoxHeight,
} from './line-layout.js';
import type { LineSpacing } from './types.js';

// Regression: some generators emit <w:spacing w:line="0" w:lineRule="exact"/> on
// table cells (e.g. private/sample-7.docx). A literal exact-0 line collapsed every
// line to 0px, so table rows fell back to the 10px minimum and cell content
// overlapped. ECMA-376 §17.3.1.33 leaves a zero line undefined; Word's native
// model ([MS-DOC] LSPD) cannot represent an exact 0 (exact = negative dyaLine)
// and resolves a non-negative dyaLine as max(dyaLine, single spacing) — so 0
// means exactly single spacing. lineBoxHeight must do the same.

const exact = (value: number): LineSpacing => ({ value, rule: 'exact', explicit: true });
const auto = (value: number): LineSpacing => ({ value, rule: 'auto', explicit: true });
const atLeast = (value: number): LineSpacing => ({
  value,
  rule: 'atLeast',
  explicit: true,
});

// Yu Mincho design line metrics per pt of font size: 1.3 × the hhea glyph box
// (asc 1802, |desc| 455, upm 2048). Sum = 1.43267 em. The line-box contract
// receives both the selected design sides and the intended single-line height.
const YU_ASC = (1802 * 1.3) / 2048;
const YU_DESC = (455 * 1.3) / 2048;
const YU = YU_ASC + YU_DESC;

describe('lineBoxHeight — degenerate zero line spacing', () => {
  // ascent 12 + descent 3 = 15px natural single-line height (no grid).
  it('treats exact line=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(exact(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('treats auto value=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(auto(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('still honors a positive exact line value', () => {
    // 12pt exact (w:line="240" twips / 20 = 12pt; LineSpacing.value is pt).
    expect(lineBoxHeight(exact(12), 12, 3, 1, undefined)).toBe(12);
  });
  it('treats negative exact/auto values as single spacing too', () => {
    expect(lineBoxHeight(exact(-5), 12, 3, 1, undefined)).toBe(15);
    expect(lineBoxHeight(auto(-1), 12, 3, 1, undefined)).toBe(15);
  });
  it('snaps a degenerate line to the grid pitch in docGrid sections', () => {
    // Same fallback as unspecified spacing: on-grid, the pitch governs.
    expect(
      lineBoxHeight(exact(0), 12, 3, 1, { type: 'lines', linePitchPt: 18 }),
    ).toBe(18);
  });
  it('still applies the auto multiplier for a positive value', () => {
    expect(lineBoxHeight(auto(2), 12, 3, 1, undefined)).toBe(30); // natural 15 × 2
  });
  it('unspecified spacing is single (natural)', () => {
    expect(lineBoxHeight(null, 12, 3, 1, undefined)).toBe(15);
  });
});

describe('lineBoxHeight — snapToChars line-grid participation', () => {
  const grid20 = { type: 'snapToChars', linePitchPt: 20 } as const;

  it('recognizes snapToChars as an active line-grid rule', () => {
    expect(isGridLineRule(grid20)).toBe(true);
  });

  it('applies the line pitch as well as the character-grid behavior', () => {
    expect(lineBoxHeight(null, 8, 2, 1, grid20, false, 0, true)).toBe(20);
  });
});

describe('lineBoxHeight — atLeast with an active line grid', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  it('retains the active grid minimum when the authored minimum is smaller', () => {
    expect(lineBoxHeight(atLeast(18), 10, 2, 1, grid20)).toBe(20);
  });

  it('retains the authored minimum when it is larger than the grid minimum', () => {
    expect(lineBoxHeight(atLeast(24), 10, 2, 1, grid20)).toBe(24);
  });

  // ECMA-376 defines the authored atLeast minimum and the active grid pitch,
  // but not this precise tall-line interaction. The unsnapped result below is
  // retained as observed Windows Word compatibility behavior.
  it('takes the maximum of content, authored minimum, and one grid pitch', () => {
    expect(lineBoxHeight(atLeast(18), 12 * YU_ASC, 12 * YU_DESC, 1, grid18, false, 12 * YU, true)).toBe(18);
  });
  it('does not round tall East Asian content up to an additional grid cell', () => {
    // Observed output keeps a 20pt design line at 28.65px on an 18px pitch,
    // rather than applying automatic grid rounding to 36px.
    expect(lineBoxHeight(atLeast(18), 20 * YU_ASC, 20 * YU_DESC, 1, grid18, false, 20 * YU, true))
      .toBeCloseTo(20 * YU, 12);
  });
  it('preserves automatic whole-cell rounding for inherited-only spacing', () => {
    const inherited = { ...atLeast(18), explicit: false };
    expect(lineBoxHeight(inherited, 20 * YU_ASC, 20 * YU_DESC, 1, grid18, false, 20 * YU, true))
      .toBe(36);
  });
  it('keeps ruby atLeast lines rounded to cells that contain the annotation', () => {
    // glyph box 41px (base + furigana reserve) → ceil(41/18) = 3 cells = 54.
    expect(lineBoxHeight(atLeast(18), 33, 8, 1, grid18, true, 0, true)).toBe(54);
  });
});

// ECMA-376 §17.6.5 / §17.3.1.32 define the line pitch as the height of ONE
// single-spaced line; how many whole grid CELLS a single-spaced East Asian line
// occupies is Word runtime behaviour. That cell count is a function of the
// line's SINGLE-LINE HEIGHT (the document font's design line height): a line
// occupies ceil(singleLineHeight / pitch) cells — the smallest number of whole
// cells that CONTAINS it. It is NOT a function of the raw em, and NOT of the
// substituted-font Canvas glyph box (which can overstate the design height).
//
// Word-PDF ground truth (pdftotext -bbox):
//   • sample-58 adjudication sweep (issue #1013; 19 sections, {10.5,12,14,16,
//     20}pt × pitch {18,24}pt × {lrTb,tbRl} × {lines,linesAndChars,none}, all
//     Yu Mincho): with Yu Mincho's design line height 1.43267 em (1.3 × hhea
//     box — see core line-metrics), EVERY measured point is ceil(design/pitch):
//     12pt→1 cell, 14pt→2 cells on an 18pt pitch; 16pt→1 cell, 20pt→2 cells on
//     a 24pt pitch. Horizontal and vertical (tbRl) sections measured IDENTICAL
//     (36.00pt column pitch = 36.00pt line pitch), and the §17.6.5 grid type
//     (lines vs linesAndChars) does not change the count.
//   • sample-35: docGrid pitch 18pt; the 12pt CJK heading (design 17.19pt) and
//     the 10.5pt body (15.04pt) are 1 cell each.
//   • sample-9 : docGrid pitch 20pt; a 20pt CJK title (design 28.65pt) is
//     2 cells (40pt).
// An earlier em-based rule (floor(em/pitch)+1) fit the sparse pre-sweep data —
// the sweep's 10.5/12/20pt rows coincide under both rules — but under-counted
// every 14–16pt line on an 18pt pitch (Word: 2 cells, em rule: 1) in BOTH
// writing directions. Latin-only lines are NOT cell-rounded (§17.6.5 leaves
// them at default spacing); the `eastAsian` flag gates the rule.
describe('docGridLineCells — natural-height East Asian grid cell count (§17.6.5)', () => {
  it('a sub-pitch line occupies a single cell (12pt Yu Mincho / 18pt pitch → 1)', () => {
    expect(docGridLineCells(12 * YU, 18)).toBe(1); // 17.19 < 18
  });
  it('a 10.5pt body line on an 18pt pitch is a single cell (sample-58 A1/B1)', () => {
    expect(docGridLineCells(10.5 * YU, 18)).toBe(1); // 15.04
  });
  it('a 14pt line on an 18pt pitch spills into TWO cells (sample-58 A3/B3)', () => {
    expect(docGridLineCells(14 * YU, 18)).toBe(2); // 20.06 > 18
  });
  it('a 16pt line on an 18pt pitch is two cells (sample-58 A4/B4)', () => {
    expect(docGridLineCells(16 * YU, 18)).toBe(2); // 22.92
  });
  it('a 16pt line on a 24pt pitch stays a single cell (sample-58 C2)', () => {
    expect(docGridLineCells(16 * YU, 24)).toBe(1); // 22.92 < 24
  });
  it('a 20pt line on a 24pt pitch is two cells (sample-58 C3)', () => {
    expect(docGridLineCells(20 * YU, 24)).toBe(2); // 28.65
  });
  it('a 20pt title on a 20pt pitch is two cells (sample-9)', () => {
    expect(docGridLineCells(20 * YU, 20)).toBe(2); // 28.65
  });
  it('a line exactly filling k pitches occupies k cells (ceil; boundary unmeasured)', () => {
    expect(docGridLineCells(18, 18)).toBe(1);
    expect(docGridLineCells(36, 18)).toBe(2);
  });
  it('a line taller than two pitches occupies three cells (extrapolated)', () => {
    expect(docGridLineCells(30 * YU, 20)).toBe(3); // 42.98 → ceil = 3
  });
  it('returns at least one cell for degenerate heights and pitches', () => {
    expect(docGridLineCells(0, 18)).toBe(1);
    expect(docGridLineCells(5, 0)).toBe(1);
  });
});

// The sample-58 adjudication matrix as measured (Word PDF, pdftotext -bbox):
// line pitch (horizontal lrTb sections) = column pitch (vertical tbRl sections)
// for every {font size × grid pitch} row. Both directions share lineBoxHeight
// (the tbRl page is laid out by the horizontal engine and rotated), so this
// table pins the shared rule against the measured matrix.
describe('lineBoxHeight — sample-58 adjudicated docGrid matrix (Yu Mincho)', () => {
  const rows: ReadonlyArray<readonly [sizePt: number, pitchPt: number, expectPt: number, sections: string]> = [
    [10.5, 18, 18, 'A1/B1'],
    [12, 18, 18, 'A2/B2/D1'],
    [14, 18, 36, 'A3/B3'],
    [16, 18, 36, 'A4/B4/D2'],
    [20, 18, 36, 'A5/B5'],
    [12, 24, 24, 'C1'],
    [16, 24, 24, 'C2'],
    [20, 24, 48, 'C3'],
  ];
  for (const [size, pitch, expected, sections] of rows) {
    it(`${size}pt on a ${pitch}pt pitch → ${expected}pt (${sections})`, () => {
      expect(lineBoxHeight(
        null,
        YU_ASC * size,
        YU_DESC * size,
        1,
        { type: 'lines', linePitchPt: pitch },
        false,
        YU * size,
        true,
      )).toBe(expected);
    });
  }
});

describe('lineBoxHeight — docGrid line-cell rounding (East Asian vs Latin)', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  // A 12pt Yu Mincho line reaches lineBoxHeight with design-corrected metrics
  // (asc 13.73 + desc 3.47 = 17.19px, the 1.43267em box). 17.19 < 18 → one cell
  // (sample-35 heading, sample-58 A2/B2).
  it('snaps a sub-pitch EA line to ONE cell of the grid pitch', () => {
    expect(lineBoxHeight(null, 12 * YU_ASC, 12 * YU_DESC, 1, grid18, false, 12 * YU, true)).toBe(18);
  });
  it('does NOT let a substitute box a hair over the pitch add a cell — grid-count height governs (sample-52)', () => {
    // Font-substitution artifact: Hiragino Mincho ProN stands in for Yu Mincho
    // and its 12pt Canvas box measures 18.0+px — a HAIR OVER the 18pt pitch —
    // while Yu Mincho's design single-line height is 17.19px, UNDER the pitch.
    // Word counts grid cells from the REAL font's design height (1 cell here),
    // never from the substituted Canvas box, so when the caller supplies the
    // line's design grid-count height (arg 9 = 12*YU) the over-tall box must not
    // inflate the count to two cells. Counting from the box mis-widths every
    // tbRl column and shifts vertical-section block tables 36pt (issue:
    // sample-52 vertical-table probe, exact table left 417→381pt).
    const boxAsc = 18.02 * 0.8; // substitute box 18.02px > the 18pt pitch
    const boxDesc = 18.02 * 0.2;
    expect(lineBoxHeight(null, boxAsc, boxDesc, 1, grid18, false, 12 * YU, true, 12 * YU)).toBe(18);
  });
  it('counts a tall UNTABLED run on a mixed line from its 1.3em FE height', () => {
    // A mixed docGrid line: a small 12pt tabled Yu Mincho run (design 17.19px)
    // plus a 24pt untabled CJK run. The latter contributes 31.2px regardless of
    // its substituted 30px Canvas box, so the line still claims two cells.
    const untabledDesign = eastAsianGridCountSinglePx(0, 24);
    expect(lineBoxHeight(null, 24, 6, 1, grid18, false, 12 * YU, true, untabledDesign)).toBe(36);
  });
  it('rounds a 20pt EA line on a 20pt pitch UP to two cells (sample-9)', () => {
    // design box 28.65px → ceil(28.65/20) = 2 cells = 40.
    expect(lineBoxHeight(null, 20 * YU_ASC, 20 * YU_DESC, 1, grid20, false, 20 * YU, true)).toBe(40);
  });
  it('uses a scale-linear 1.3em fallback for an untabled EA line', () => {
    // MS Mincho's hhea box is 1.0em and Word's FE single-line height is 1.3 ×
    // hhea. The substituted Canvas box is exactly 1.0em here, but a 20pt line
    // on a 20pt grid still occupies two cells in the sample-9 Word PDF.
    expect(lineBoxHeight(null, 18, 2, 1, grid20, false, 0, true, undefined, 20)).toBe(40);
    expect(lineBoxHeight(null, 43, 5, 2.3529, grid20, false, 0, true, undefined, 20 * 2.3529)).toBeCloseTo(40 * 2.3529, 8);
  });
  it('uses the tabled design floor when no per-line grid-count height is supplied', () => {
    // A direct caller without arg 9 still has a deterministic tabled-font input
    // in intendedSinglePx. The substituted 30px box is not consulted for the
    // cell count; Yu Mincho's 28.65px design height takes two 20px cells.
    expect(lineBoxHeight(null, 24, 6, 1, grid20, false, 20 * YU, true)).toBe(40);
  });
  it('keeps a Latin line at natural height (one-cell floor), NOT cell-rounded', () => {
    // 22px natural, Latin → max(22, 20) = 22px (Word does not cell-round it).
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20, false, 0, false)).toBe(22);
  });
  it('combines inherited spacing with useFELayout-classified Far East grid metrics', () => {
    const inherited = { value: 1.15, rule: 'auto', explicit: false } as const;
    // An 18pt hinted Latin title has a 1.3em FE design height and therefore
    // needs two 18pt cells. The allocation itself exceeds the 1.15-pitch floor.
    expect(lineBoxHeight(
      inherited, 17, 4, 1, grid18, false, 0, true,
      eastAsianGridCountSinglePx(0, 18),
    )).toBe(36);
    // A smaller hinted Latin label stays in one cell, but inherited 1.15 line
    // spacing raises the one-pitch minimum to 20.7pt.
    expect(lineBoxHeight(
      inherited, 7, 2, 1, grid18, false, 0, true,
      eastAsianGridCountSinglePx(0, 8),
    )).toBeCloseTo(20.7, 12);
    // Face hinting alone never changes an empty/Latin mark's grid class.
    expect(lineBoxHeight(inherited, 14, 2, 1, grid18, false, 0, false)).toBe(18);
  });
  it('a RUBY EA line reserves its measured furigana height, NOT the design cell count', () => {
    // sample-5: a 13.5pt ruby base + 8pt rt on an 18pt pitch reserves the base +
    // annotation glyph box (~41px = asc 33 + desc 8), which Word spreads over
    // whole cells (ceil(41/18) = 3 cells = 54px). The base design height (13.5pt
    // → 19.3px → 2 cells) would clip the furigana, so the natural-height rule
    // must NOT apply to ruby lines — hasRuby (arg 6) keeps them on the measured
    // glyph box.
    expect(lineBoxHeight(null, 33, 8, 1, grid18, true, 0, true)).toBe(54);
  });
  it('does not cell-round when no docGrid is active (East Asian, off grid)', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, undefined, false, 0, true)).toBe(22);
  });
  it('defaults to Latin (no cell rounding) when the flag is omitted', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20)).toBe(22);
  });
});
