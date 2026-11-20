import referenceData from '../fonts/reference-font-metrics-data.json';
import { describe, expect, it } from 'vitest';
import { layoutStackedText, type StackedParagraph, type StackedLayoutOptions } from './stacked.js';
import {
  stackedFaceBox, stackedVerticalGlyph, STACKED_CELL_FACTOR, STACKED_SIDEWAYS,
} from './stacked-faces.js';
import { stackedClusters, stackedVerticalGlyphText } from './stacked-glyphs.js';

// Arial 24 pt at 1 px/pt: cell = 7/6 × usWin (1.1172 em) × 24 = 31.28.
const CELL = STACKED_CELL_FACTOR * (2288 / 2048) * 24;
const RECT = { left: 0, top: 0, width: 185.6, height: 172.8 };

type Style = { tag: string };
const S: Style = { tag: 'a' };
const para = (texts: string[], extra: Partial<StackedParagraph<Style>> = {}): StackedParagraph<Style> => ({
  runs: texts.map((text) => ({ type: 'text' as const, text, style: S })),
  emptyThickness: CELL,
  ...extra,
});
const options = (extra: Partial<StackedLayoutOptions<Style>> = {}): StackedLayoutOptions<Style> => ({
  direction: 'wordArtVert',
  rect: RECT,
  wrap: true,
  pxPerPt: 1,
  glyphs: (text, style) => [...text].map((ch) => ({
    text: ch, style, kind: 'upright' as const, advance: CELL, thickness: CELL, space: ch === ' ',
  })),
  ...extra,
});
const axes = (paragraphs: StackedParagraph<Style>[], extra: Partial<StackedLayoutOptions<Style>> = {}) =>
  layoutStackedText(paragraphs, options(extra)).columns.map((c) => c.axisX);

// Expected values follow the model; the comments give the PowerPoint export
// (issue #1626 controls, text y corrected for the export's 1.0015 scale).
describe('layoutStackedText (wordArtVert / wordArtVertRtl)', () => {
  it('stacks columns from the start side, half a cell from the edge', () => {
    // FV02 "ABCDEFGHIJKLMN" in a 172.8 pt column: 5 per column, axes 15.23 / 46.90 / 77.93.
    const l = layoutStackedText([para(['ABCDEFGHIJKLMN'])], options());
    expect(l.columns.map((c) => c.axisX)).toEqual([CELL / 2, CELL * 1.5, CELL * 2.5].map((v) => expect.closeTo(v, 6)));
    expect(l.glyphs.filter((g) => g.column === 0).map((g) => g.text).join('')).toBe('ABCDE');
    expect(l.glyphs[1].cellTop).toBeCloseTo(CELL, 6);
    // FR02: right to left from the right edge (15.65 / 46.65 / 78.30 from the right).
    expect(axes([para(['ABCDEFGHIJKLMN'])], { direction: 'wordArtVertRtl' }))
      .toEqual([185.6 - CELL / 2, 185.6 - CELL * 1.5, 185.6 - CELL * 2.5].map((v) => expect.closeTo(v, 6)));
  });

  it('wraps at spaces, keeps the trailing space in its column and starts a column per paragraph', () => {
    // FV05 "Hi all foo bar": columns "Hi␠" / "all␠" / "foo␠" / "bar".
    const l = layoutStackedText([para(['Hi all foo bar'])], options());
    const cols = [0, 1, 2, 3].map((c) => l.glyphs.filter((g) => g.column === c).map((g) => g.text).join(''));
    expect(cols).toEqual(['Hi ', 'all ', 'foo ', 'bar']);
    expect(axes([para(['AB']), para(['CDE']), para(['F'])])).toHaveLength(3);
    // wrap="none" never breaks (FV06 runs past the box bottom).
    expect(layoutStackedText([para(['ABCDEFGH'])], options({ wrap: false })).columns).toHaveLength(1);
  });

  // Whole-word matching reads `span`: glyphs of one span are continuous text.
  it.each(['wordArtVert', 'wordArtVertRtl'] as const)(
    'gives the glyphs of %s their span of authored text', (direction) => {
      const spans = (paragraphs: StackedParagraph<Style>[], extra: Partial<StackedLayoutOptions<Style>> = {}) => {
        const glyphs = layoutStackedText(paragraphs, options({ direction, ...extra })).glyphs;
        // Renumber from zero: the values need not be consecutive.
        const seen: number[] = [];
        return glyphs.map((g) => {
          if (!seen.includes(g.span)) seen.push(g.span);
          return `${g.text}${seen.indexOf(g.span)}`;
        }).join(' ');
      };
      // A word that wraps into the next column inside the word is one span; a
      // paragraph and a manual break start another.
      expect(spans([para(['ABCDEFG'])])).toBe('A0 B0 C0 D0 E0 F0 G0');
      expect(spans([para(['AB']), para(['CD'])])).toBe('A0 B0 C1 D1');
      expect(spans([{ ...para([]), runs: [
        { type: 'text', text: 'AB', style: S }, { type: 'break' }, { type: 'text', text: 'CD', style: S },
      ] }])).toBe('A0 B0 C1 D1');
      // A wrap at a space that stays in the column keeps the span (the space
      // itself separates the words); a wrap that drops the space ends it.
      expect(spans([para(['Hi all'])])).toBe('H0 i0  0 a0 l0 l0');
      expect(spans([para(['ABCDE FG'])])).toBe('A0 B0 C0 D0 E0 F1 G1');
      // A tab separates although it adds no glyph.
      expect(spans([para(['AB\tCD'])])).toBe('A0 B0 C1 D1');
      // An inline object separates the text around it, drawn or zero-size.
      const withObject = [{ ...para([]), runs: [
        { type: 'text' as const, text: 'AB', style: S },
        { type: 'object' as const, width: 10, style: S },
        { type: 'text' as const, text: 'CD', style: S },
      ] }];
      const object = (advance: number) => ({ objectGlyph: () => ({
        text: '\uFFFC', style: S, kind: 'upright' as const, advance, thickness: advance, space: false,
      }) });
      expect(spans(withObject, object(10))).toBe('A0 B0 \uFFFC1 C2 D2');
      expect(spans(withObject, object(0))).toBe('A0 B0 C1 D1');
    },
  );

  it('puts lnSpc extra on the page-right side and paragraph spacing per direction', () => {
    const spaced = (extra: Partial<StackedParagraph<Style>>) => [para(['ABCD'], extra), para(['EFGH'], extra)];
    const pct150 = { lineSpacing: { type: 'pct' as const, val: 150000 }, spaceBefore: { type: 'pts' as const, val: 12 } };
    // FV07: 16.08, 62.89 — spcBef lies right of paragraph 2's column.
    expect(axes(spaced(pct150))).toEqual([CELL / 2, CELL * 1.5 + CELL / 2].map((v) => expect.closeTo(v, 6)));
    // FR07: 30.64, 89.71 from the right — spcBef between the paragraphs.
    const r = axes(spaced(pct150), { direction: 'wordArtVertRtl' }).map((x) => 185.6 - x);
    expect(r).toEqual([CELL, CELL * 2 + 12 + CELL / 2].map((v) => expect.closeTo(v, 6)));
    // SPV02 spcAft 12 on all three: 27.53, 70.67, 102.46 — the first paragraph's
    // spcAft lies left of its column, the last one's is suppressed.
    const aft = [para(['AB'], { spaceAfter: { type: 'pts', val: 12 } }), para(['CD'], { spaceAfter: { type: 'pts', val: 12 } }),
      para(['EF'], { spaceAfter: { type: 'pts', val: 12 } })];
    expect(axes(aft)).toEqual([12 + CELL / 2, 24 + CELL * 1.5, 24 + CELL * 2.5].map((v) => expect.closeTo(v, 6)));
    // spcFirstLastPara restores it (SPV06: 113.99 for the third column).
    expect(axes(aft, { spcFirstLastPara: true })[2]).toBeCloseTo(36 + CELL * 2.5, 6);
  });

  it('keeps the natural descent of the last line when spacing shortens the lines', () => {
    // LSVU80 lnSpc 80 %: pitches 25.22 / 24.48 / 25.22 / 26.64 (H 25.02; last H + (L - H) / 4).
    const H = CELL * 0.8;
    const five = [0, 1, 2, 3, 4].map(() => para(['HH'], { lineSpacing: { type: 'pct', val: 80000 } }));
    const v = axes(five, { rect: { ...RECT, width: 205.6 } });
    const pitches = v.slice(1).map((x, i) => x - v[i]);
    expect(pitches.slice(0, 3)).toEqual([H, H, H].map((p) => expect.closeTo(p, 6)));
    expect(pitches[3]).toBeCloseTo(H + (CELL - H) / 4, 6);
    // First column: the reduced descent 0.25 H + L / 4 (LSVU80 14.37).
    expect(v[0]).toBeCloseTo(0.25 * H + CELL / 4, 6);
    // wordArtVertRtl keeps every pitch at H; the last descent sits on the far side.
    const r = axes(five, { direction: 'wordArtVertRtl', rect: { ...RECT, width: 205.6 } });
    expect(r.slice(1).map((x, i) => r[i] - x)).toEqual([H, H, H, H].map((p) => expect.closeTo(p, 6)));
  });

  it('anchors the block across the columns and anchorCtr centres it along them', () => {
    const two = [para(['AB']), para(['CDE'])];
    // ANV3 anchor b: the last column's axis half a cell from the right.
    expect(axes(two, { anchor: 'b' })[1]).toBeCloseTo(185.6 - CELL / 2, 6);
    expect(axes(two, { anchor: 'ctr' })[0]).toBeCloseTo((185.6 - 2 * CELL) / 2 + CELL / 2, 6);
    // ANR3 anchor b in Rtl: the block sits at the left edge.
    expect(axes(two, { anchor: 'b', direction: 'wordArtVertRtl' })[1]).toBeCloseTo(CELL / 2, 6);
    // ANV4 anchorCtr: columns keep a common top, the longest one centred.
    const l = layoutStackedText(two, options({ anchorCtr: true }));
    expect(l.glyphs[0].cellTop).toBeCloseTo((172.8 - 3 * CELL) / 2, 6);
    expect(l.glyphs[2].cellTop).toBeCloseTo(l.glyphs[0].cellTop, 6);
  });

  it('aligns along each column; just stretches spaces, dist spreads glyphs', () => {
    const tops = (algn: string, text = 'BCD') =>
      layoutStackedText([para([text], { alignment: algn })], options({ rect: { ...RECT, height: 212.8 } }))
        .glyphs.filter((g) => !g.space).map((g) => g.cellTop);
    expect(tops('r')[2]).toBeCloseTo(212.8 - CELL, 6); // ALV3: last glyph at the bottom
    expect(tops('ctr')[0]).toBeCloseTo((212.8 - 3 * CELL) / 2, 6);
    expect(tops('dist')).toEqual([0, (212.8 - CELL) / 2, 212.8 - CELL].map((v) => expect.closeTo(v, 6)));
    expect(tops('dist', 'A')[0]).toBeCloseTo((212.8 - CELL) / 2, 6); // one glyph is centred
    // ALV6 just "AB CDE FG HIJ KL": the first column ends at the bottom through its space.
    const just = tops('just', 'AB CDE FG HIJ KL');
    expect(just[4]).toBeCloseTo(212.8 - CELL, 6);
    expect(just[1]).toBeCloseTo(CELL, 6);
  });
});

describe('layoutStackedText input size', () => {
  it('lays out a 150,000-glyph run without spreading it into an argument list', () => {
    const text = 'A'.repeat(150_000);
    const one = layoutStackedText([para([text])], options({ wrap: false }));
    expect(one.glyphs).toHaveLength(150_000);
    expect(one.columns[0].thickness).toBeCloseTo(CELL, 6);
    const centred = layoutStackedText([para([text])], options({ wrap: false, anchorCtr: true }));
    expect(centred.glyphs[0].cellTop).toBeCloseTo((172.8 - 150_000 * CELL) / 2, 3);
    // Wrapped into columns of five cells: 30,000 columns in one paragraph.
    const wrapped = layoutStackedText([para(['A '.repeat(75_000)])], options());
    expect(wrapped.columns.length).toBeGreaterThan(20_000);
  }, 60_000);
});

describe('stacked face and character tables', () => {
  it('uses usWin metrics even for formerly fitted faces and their aliases', () => {
    expect(stackedFaceBox('游ゴシック', false, false)).toEqual({ ascent: 2017 / 2048, descent: 619 / 2048 });
    expect(stackedFaceBox('SimSun', false, false)).toEqual({ ascent: 220 / 256, descent: 36 / 256 });
    expect(stackedFaceBox('Yu Gothic', true, false)).toEqual({ ascent: 2024 / 2048, descent: 612 / 2048 });
    const light = stackedFaceBox('Yu Mincho Light', false, false);
    expect((light?.ascent ?? 0) + (light?.descent ?? 0)).toBeCloseTo(2636 / 2048, 6);
    expect(stackedFaceBox('No Such Face', false, false)).toBeUndefined();
  });

  it('uses typo metrics including lineGap when USE_TYPO_METRICS is set', () => {
    // Gabriola's usWin box (7540 units) differs from its typo box (6963).
    expect(stackedFaceBox('Gabriola', false, false)).toEqual({ ascent: (2800 + 2867) / 4096, descent: 1296 / 4096 });
  });

  it('does not shrink the typo box for a negative lineGap', () => {
    // Tamil Sangam MN sets USE_TYPO_METRICS with sTypoLineGap = -210.
    expect(stackedFaceBox('Tamil Sangam MN', false, false))
      .toEqual({ ascent: 1550 / 2048, descent: 717 / 2048 });
  });

  it('uses the requested face coverage when weights differ, including named aliases', () => {
    const cp = (ch: string) => ch.codePointAt(0) as number;
    expect(stackedVerticalGlyph('Microsoft JhengHei', cp('ー'))).toBe(true);
    expect(stackedVerticalGlyph('Microsoft JhengHei', cp('ー'), true)).toBe(false);
    expect(stackedVerticalGlyph('MicrosoftJhengHeiBold', cp('ー'))).toBe(false);
    expect(stackedVerticalGlyph('Yu Gothic', cp('〈'))).toBe(true);
    expect(stackedVerticalGlyph('Yu Gothic Medium', cp('〈'))).toBe(false);
  });

  it('preserves bold coverage and metrics when the italic face is missing', () => {
    for (const ch of 'ー‥') {
      const cp = ch.codePointAt(0) as number;
      expect(stackedVerticalGlyph('Microsoft JhengHei', cp, true, true)).toBe(false);
      expect(stackedVerticalGlyph('MicrosoftJhengHeiBold', cp, true, true)).toBe(false);
    }
    // Yu Gothic's bold face box differs from Regular, unlike JhengHei's.
    expect(stackedFaceBox('Yu Gothic', true, true))
      .toEqual({ ascent: 2024 / 2048, descent: 612 / 2048 });
  });

  it('uses regular coverage and metrics when only italic is requested but missing', () => {
    for (const ch of 'ー‥') {
      expect(stackedVerticalGlyph('Microsoft JhengHei', ch.codePointAt(0) as number, false, true)).toBe(true);
    }
    expect(stackedFaceBox('Yu Gothic', false, true))
      .toEqual({ ascent: 2017 / 2048, descent: 619 / 2048 });
  });

  it('retains a named weight when the requested weight is missing', () => {
    for (const [bold, italic] of [[false, false], [true, true]] as const) {
      // Medium has no vert substitution for 〈, whereas Regular does.
      expect(stackedVerticalGlyph('Yu Gothic Medium', 0x3008, bold, italic)).toBe(false);
      expect(stackedVerticalGlyph('Yu Gothic Medium', 0xff08, bold, italic)).toBe(true);
      expect(stackedFaceBox('Yu Gothic Medium', bold, italic))
        .toEqual({ ascent: 2028 / 2048, descent: 608 / 2048 });
    }
  });

  it('measures only the measured set sideways, upright in faces with vert glyphs', () => {
    const advances = { upright: 10, verticalGlyph: 20, rotated: (ch: string) => 30 + ch.length };
    const kinds = (family: string, text: string, bold = false) =>
      stackedClusters(text, { family, bold, italic: false }, advances)
        .map((c) => (c.nominal ? 'n' : c.kind[0])).join('');
    // u = upright, v = vertical glyph (vert), n = unturned nominal glyph in a
    // 1 em cell (vrt2 only, outside Basic Latin), r = rotated.
    expect(kinds('ＭＳ ゴシック', '(（「、ー…∥，．：-!—ぁ〜A日')).toBe('rvvvvvv' + 'u'.repeat(10));
    // #1827 PowerPoint controls: ∥ is vrt2-only in Yu Gothic, Meiryo and Yu
    // Mincho and stood unturned; the vrt2-only ASCII brackets still turned.
    for (const family of ['Yu Gothic', 'Yu Gothic Medium', 'Yu Gothic UI', 'Meiryo', 'Meiryo UI', 'Yu Mincho']) {
      for (const bold of [false, true]) expect(kinds(family, '(<[{（∥', bold), family).toBe('rrrrvn');
    }
    expect(kinds('ＭＳ 明朝', '(∥')).toBe('rv');
    expect(kinds('Arial', '(（')).toBe('rr');
    expect(kinds('Yu Mincho Light', '（')).toBe('v');
    expect(kinds('Arial Unicode MS', '(')).toBe('v');
    expect(kinds('No Such Face', '（')).toBe('r');
    // One cell per grapheme cluster, classed by its base; advances by class.
    expect(stackedClusters('（́A(', { family: 'MS Gothic', bold: false, italic: false }, advances))
      .toEqual([
        { text: '（́', kind: 'verticalGlyph', advance: 20 },
        { text: 'A', kind: 'upright', advance: 10 },
        { text: '(', kind: 'rotated', advance: 31 },
      ]);
    // Paragraph markers stand upright whatever their characters.
    expect(stackedClusters('（', { family: 'MS Gothic', bold: false, italic: false }, advances, true)[0].kind)
      .toBe('upright');
  });

  it('draws a lone vert glyph as its presentation form and a nominal glyph as itself', () => {
    expect(stackedVerticalGlyphText({ text: '（' })).toBe('︵');
    expect(stackedVerticalGlyphText({ text: '、' })).toBe('︑');
    expect(stackedVerticalGlyphText({ text: '（́' })).toBeNull();
    expect(stackedVerticalGlyphText({ text: 'ー' })).toBeNull();
    expect(stackedVerticalGlyphText({ text: '∥', nominal: true })).toBe('∥');
  });
});

it.each(['wordArtVert', 'wordArtVertRtl'] as const)('%s wraps an upright inline object atomically using its column extent', (direction) => {
  const object = { type: 'object' as const, width: 48, style: S, payload: 'equation' };
  const result = layoutStackedText([para([], { runs: [
    { type: 'text', text: 'A', style: S }, object,
    { type: 'text', text: 'B', style: S },
  ] })], options({ direction, rect: { ...RECT, height: CELL + 48 },
    objectGlyph: (segment) => ({ text: String(segment.payload), style: segment.style,
      kind: 'upright', advance: segment.width, thickness: 72, space: false }),
  }));
  expect(result.glyphs.map((g) => [g.text, g.column, g.cellTop])).toEqual([
    ['A', 0, 0], ['equation', 0, CELL], ['B', 1, 0],
  ]);
  expect(result.columns[0].thickness).toBe(72);
  expect(Math.sign(result.columns[1].axisX - result.columns[0].axisX))
    .toBe(direction === 'wordArtVert' ? 1 : -1);
});

it.each(['wordArtVert', 'wordArtVertRtl'] as const)('%s preserves blank display columns and zero inline advance for unavailable objects', (direction) => {
  for (const display of [false, true]) {
    const result = layoutStackedText([para([], { runs: [
      { type: 'text', text: 'A', style: S },
      { type: 'object', width: 0, style: S, display },
      { type: 'text', text: 'B', style: S },
    ] })], options({ direction, objectGlyph: ({ style }) => ({
      text: '', style, kind: 'upright', advance: 0, thickness: 0, space: false,
    }) }));
    expect(result.columns.map((c) => c.thickness)).toEqual(display ? [CELL, CELL, CELL] : [CELL]);
    expect(result.blockWidth).toBeCloseTo((display ? 3 : 1) * CELL);
    const [a, b] = result.glyphs;
    expect(result.glyphs.map((g) => [g.text, g.column, g.cellTop])).toEqual([
      ['A', 0, 0], ['B', display ? 2 : 0, display ? 0 : CELL],
    ]);
    expect(b.axisX - a.axisX).toBeCloseTo(display ? (direction === 'wordArtVert' ? 2 : -2) * CELL : 0);
  }
});

// The prefix is atomic: it consumes the first column's wrap/alignment budget,
// never repeats on a continuation, and participates in that column's thickness.
it('includes an atomic prefix in first-column wrapping, alignment and anchoring', () => {
  const p = para(['ABC'], { firstLineIndent: -CELL / 2,
    firstLinePrefix: { advance: 2 * CELL, thickness: 2 * CELL } });
  const l = layoutStackedText([p], options({ rect: { ...RECT, height: 3 * CELL } }));
  expect(l.glyphs.map(g => [g.text, g.column, g.cellTop / CELL])).toEqual([
    ['A', 0, 1.5], ['B', 1, 0], ['C', 1, 1],
  ]);
  expect(l.columns[0].thickness).toBe(2 * CELL);
  expect(l.columns[0].prefixTop).toBeCloseTo(-CELL / 2);
  expect(l.columns[1].prefixTop).toBeUndefined();
  const centered = layoutStackedText([para(['A'], { firstLinePrefix: p.firstLinePrefix })],
    options({ rect: { ...RECT, height: 4 * CELL }, anchorCtr: true }));
  expect(centered.columns[0].prefixTop).toBeCloseTo(CELL / 2);
  expect(centered.glyphs[0].cellTop).toBeCloseTo(2.5 * CELL);
});

it('keeps the generator coverage domain identical to the stacked sideways classifier', () => {
  expect([...referenceData.stackedSideways].map((ch) => ch.codePointAt(0))).toEqual([...STACKED_SIDEWAYS]);
});
