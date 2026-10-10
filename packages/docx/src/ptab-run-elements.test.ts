import { describe, it, expect, vi } from 'vitest';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { buildTextIndex, findMatches, createCanvasFontRoute } from '@silurus/ooxml-core';
import { renderDocumentToCanvas } from './renderer.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { eastAsianUprightPaintOps } from './layout/vertical-glyph-orientation.js';
import { hitTestDocxElementContext } from './element-context.js';
import { buildPageLayers } from './layout/page-layers.js';
import { createPaintResourceRegistry } from './layout/paint-resources.js';
import {
  splitTextForLayout,
  layoutLines,
  buildSegments,
  type LayoutTextSeg,
  type LineLayoutEnvironment,
} from './line-layout.js';
import type { DocParagraph, DocxDocumentModel, SectionProps, DocRun } from './types.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';

// ECMA-376 §17.3.3 run-content elements that were previously dropped by the
// parser's `_ => {}` arm and thus never reached the renderer:
//   §17.3.3.23 <w:ptab>        — absolute-position tab
//   §17.3.3.18 <w:noBreakHyphen> — non-breaking hyphen glyph
//   §17.3.3.29 <w:softHyphen>  — optional hyphen (visible only at a selected break)
// These end-to-end tests record fillText() calls to pin the layout geometry the
// parser + line-layout now produce. Scale is 1 px/pt (canvas width == pageWidth)
// and every glyph is FS px wide in the recording canvas.

interface FillCall {
  text: string;
  x: number;
}

function makeRecordingCanvas(widthOf?: (text: string, fontSize: number) => number): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: widthOf ? widthOf(s, p) : [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number) { fills.push({ text, x }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string): Extract<DocRun, { type: 'text' }> {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as Extract<DocRun, { type: 'text' }>;
}

function para(runs: DocRun[], indent: { left?: number; right?: number } = {}): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: indent.left ?? 0, indentRight: indent.right ?? 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs,
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function ptabRun(
  alignment: 'left' | 'center' | 'right',
  relativeTo: 'margin' | 'indent',
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot' = 'none',
): DocRun {
  return { type: 'ptab', alignment, relativeTo, leader, fontSize: 10 } as unknown as DocRun;
}

const PAGE_W = 300;

function doc(paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_W, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(paras: DocParagraph[]) {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(paras), canvas, 0, { dpr: 1, width: PAGE_W });
  return fills;
}

describe('ptab (§17.3.3.23) absolute-position tab layout', () => {
  const FS = 10; // glyph width px in the recording canvas; scale = 1 px/pt

  // relativeTo="margin": the reference box is the full text margin [0, PAGE_W]
  // (no indents here). center → PAGE_W/2, right → PAGE_W. The trailing text
  // aligns to the position per the alignment (center/right).

  it('center ptab relative to margin centers the trailing text on the line', async () => {
    const fills = await render([para([ptabRun('center', 'margin'), textRun('PAGE')])]);
    const f = fills.find((c) => c.text === 'PAGE');
    expect(f, '"PAGE" must be drawn').toBeDefined();
    // 4 glyphs wide → centered on PAGE_W/2 = 150 ⇒ starts at 150 − (4·FS)/2 = 130.
    expect(f!.x).toBeCloseTo(PAGE_W / 2 - (4 * FS) / 2, 3);
  });

  it('right ptab relative to margin right-aligns the trailing text to the margin', async () => {
    const fills = await render([para([ptabRun('right', 'margin'), textRun('12')])]);
    const f = fills.find((c) => c.text === '12');
    expect(f, '"12" must be drawn').toBeDefined();
    // Right edge on the margin (PAGE_W = 300) ⇒ 2-glyph number starts at 300 − 20 = 280.
    expect(f!.x + 2 * FS).toBeCloseTo(PAGE_W, 3);
  });

  it('left ptab relative to margin left-aligns the trailing text at the margin', async () => {
    const fills = await render([para([ptabRun('left', 'margin'), textRun('X')])]);
    const f = fills.find((c) => c.text === 'X');
    expect(f, '"X" must be drawn').toBeDefined();
    // A left ptab at the margin, with the pen already there, is a no-op advance.
    expect(f!.x).toBeCloseTo(0, 3);
  });

  // relativeTo="indent": the reference box is the paragraph content box, i.e.
  // between the left/right indents. With a 40 pt left indent and 20 pt right
  // indent, the content box is [0, PAGE_W − 40 − 20] = [0, 240] in paraX-relative
  // coordinates. A right ptab lands the text at the content-box right edge, which
  // sits at absolute X = 40 (left indent) + 240 = 280.

  it('right ptab relative to indent aligns to the content box, not the margin', async () => {
    const fills = await render([
      para([ptabRun('right', 'indent'), textRun('99')], { left: 40, right: 20 }),
    ]);
    const f = fills.find((c) => c.text === '99');
    expect(f, '"99" must be drawn').toBeDefined();
    // content-box right edge (absolute) = leftIndent(40) + (PAGE_W−40−20)=240 → 280.
    // 2-glyph number ends there ⇒ starts at 280 − 20 = 260.
    const contentRightAbs = 40 + (PAGE_W - 40 - 20);
    expect(f!.x + 2 * FS).toBeCloseTo(contentRightAbs, 3);
  });

  it('contains a margin ptab whose target is past the paragraph right indent', async () => {
    const fills = await render([
      para([ptabRun('right', 'margin'), textRun('99')], { left: 40, right: 20 }),
    ]);
    const f = fills.find((c) => c.text === '99');
    expect(f, '"99" must be drawn').toBeDefined();
    // The margin target is 300 pt, outside this paragraph's 280 pt band.
    // Library containment policy discards an unreachable empty-line gap.
    expect(f!.x).toBeGreaterThanOrEqual(40);
    expect(f!.x + 2 * FS).toBeLessThanOrEqual(PAGE_W - 20);
  });
});

describe('noBreakHyphen (§17.3.3.18) and softHyphen (§17.3.3.29)', () => {
  it('noBreakHyphen draws a U+002D hyphen glyph inline', async () => {
    // Runs: "999" | <noBreakHyphen "-"> | "99" — the parser injects "-" so the
    // renderer draws a hyphen between the numbers.
    const fills = await render([para([textRun('999'), textRun('-'), textRun('99')])]);
    const drawn = fills.map((c) => c.text).join('');
    expect(drawn).toContain('-');
    expect(drawn).toContain('999');
    expect(drawn).toContain('99');
  });

  // §17.3.3.18: "without that hyphen being a line breaking position". The
  // parser injects a real '-' (U+002D) into the run's text rather than a
  // dedicated break-suppressing token, so the WHOLE non-breaking guarantee
  // rests on `splitTextForLayout` (line-layout.ts) never treating '-' as a
  // token boundary — it must open break opportunities at spaces ONLY. Pin
  // that contract directly: if `splitTextForLayout` is ever changed to also
  // split on '-' (e.g. to support ordinary-hyphen wrapping), this fails.
  it('splitTextForLayout does not open a break opportunity at a hyphen', () => {
    expect(splitTextForLayout('999-99-9999')).toEqual(['999-99-9999']);
    // Trailing spaces travel WITH the preceding token (splitTextForLayout's
    // documented behaviour), so "co-operative " keeps its space; the point
    // here is that '-' inside the token never itself starts a new token.
    expect(splitTextForLayout('co-operative society')).toEqual(['co-operative ', 'society']);
  });

  // End-to-end companion: run the exact single-token text a same-formatting
  // noBreakHyphen merge produces (`text_runs_mergeable`/parser.rs — see the
  // spec's own §17.3.3.18 example, "999-99-9999" split into three <w:r> at
  // the hyphen positions and merged back into one DocRun::Text) through the
  // REAL tokenizer (`buildSegments`, which calls `splitTextForLayout`) and
  // then `layoutLines`, placed after a leading word so the line's REMAINING
  // width is too small for the token but the token itself easily fits the
  // full line width. This isolates the wrap decision from the (correct,
  // separate) over-long-word char-break path — see line-layout.ts's
  // "over-long-word" comment — which force-splits a token WIDER THAN THE
  // WHOLE LINE and would otherwise be indistinguishable from an incorrect
  // hyphen-triggered split. Assert the token moves to the next line WHOLE.
  it('a merged noBreakHyphen token wraps to the next line whole, never splitting at the hyphen', () => {
    const merged = textRun('lead 999-99') as DocRun & {
      noBreakRanges?: readonly Readonly<{ start: number; end: number }>[];
    };
    // The production parser preserves authored noBreakHyphen ownership even
    // after a same-format merge. Ordinary U+002D carries no such protection.
    merged.noBreakRanges = [{ start: 8, end: 9 }];
    const segs = buildSegments([merged], {} as LineLayoutEnvironment);
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    // Line width 100px. "lead " (5 glyphs * 10px = 50px) leaves 50px
    // remaining — not enough for "999-99" (6 glyphs * 10px = 60px), but
    // "999-99" alone is well under the full 100px line width, so this must
    // hit the "does not fit the CURRENT line" wrap path, not the
    // over-long-word char-break path.
    const lines = layoutLines(ctx, segs, 100, 0, 1);
    const allTexts = lines.map((l) => l.segments.map((s) => (s as LayoutTextSeg).text));
    expect(allTexts).toEqual([['lead '], ['999-99']]);
  });

  it('keeps a comment-boundary noBreakHyphen run joined without coalescing model runs', () => {
    const hyphenRun = textRun('-cd') as DocRun & { noBreakBefore?: boolean };
    hyphenRun.noBreakBefore = true;
    const segs = buildSegments(
      [textRun('lead '), textRun('ab'), hyphenRun],
      {} as LineLayoutEnvironment,
    );
    expect(segs.filter((seg): seg is LayoutTextSeg => 'text' in seg).map((seg) => [
      seg.text,
      seg.joinPrev,
    ])).toEqual([
      ['lead ', undefined],
      ['ab', undefined],
      ['-cd', true],
    ]);

    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 70, 0, 1);
    expect(lines.map((line) => line.segments.map((seg) => (seg as LayoutTextSeg).text)))
      .toEqual([['lead '], ['ab', '-cd']]);
  });

  it('keeps the following CT_R joined when noBreakHyphen ends its own run', () => {
    const mergedLeft = textRun('ab-') as DocRun & { noBreakAfter?: boolean };
    mergedLeft.noBreakAfter = true;
    const separateHyphen = textRun('-') as DocRun & {
      noBreakBefore?: boolean;
      noBreakAfter?: boolean;
    };
    separateHyphen.noBreakBefore = true;
    separateHyphen.noBreakAfter = true;
    const formattedFollower = { ...textRun('cd'), bold: true } as DocRun;

    expect(buildSegments(
      [mergedLeft, textRun('cd')],
      {} as LineLayoutEnvironment,
    ).filter((seg): seg is LayoutTextSeg => 'text' in seg).map((seg) => [
      seg.text,
      seg.joinPrev,
    ])).toEqual([
      ['ab-', undefined],
      ['cd', true],
    ]);

    const segs = buildSegments(
      [textRun('lead '), textRun('ab'), separateHyphen, formattedFollower],
      {} as LineLayoutEnvironment,
    );
    expect(segs.filter((seg): seg is LayoutTextSeg => 'text' in seg).map((seg) => [
      seg.text,
      seg.joinPrev,
    ])).toEqual([
      ['lead ', undefined],
      ['ab', undefined],
      ['-', true],
      ['cd', true],
    ]);

    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 70, 0, 1);
    expect(lines.map((line) => line.segments.map((seg) => (seg as LayoutTextSeg).text)))
      .toEqual([['lead '], ['ab', '-', 'cd']]);
  });

  it.each([
    ['CJK', '漢-字語'],
    ['mixed CJK and SEA', '漢-字ไทย'],
  ])('protects both edges of noBreakHyphen through %s split paths', (_name, text) => {
    const run = textRun(text) as DocRun & {
      noBreakRanges?: readonly Readonly<{ start: number; end: number }>[];
    };
    run.noBreakRanges = [{ start: 1, end: 2 }];
    const segs = buildSegments([run], {} as LineLayoutEnvironment);
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 20, 0, 1);
    const lineTexts = lines.map((line) => line.segments
      .map((segment) => (segment as LayoutTextSeg).text)
      .join(''));

    expect(lineTexts.join('')).toBe(text);
    expect(lineTexts[0]).toBe('漢-字');
    expect(lineTexts.every((line) => !line.endsWith('漢') && !line.startsWith('-'))).toBe(true);
    expect(lineTexts.every((line) => !line.endsWith('-') && !line.startsWith('字'))).toBe(true);
  });

  it.each([
    ['CJK', '字語'],
    ['mixed CJK and SEA', '字ไทย'],
  ])('moves a cross-formatting noBreakHyphen %s group to a fresh line', (_name, followerText) => {
    const hyphen = textRun('-') as DocRun & {
      noBreakBefore?: boolean;
      noBreakAfter?: boolean;
      noBreakRanges?: readonly Readonly<{ start: number; end: number }>[];
    };
    hyphen.noBreakBefore = true;
    hyphen.noBreakAfter = true;
    hyphen.noBreakRanges = [{ start: 0, end: 1 }];
    const follower = { ...textRun(followerText), bold: true } as DocRun;
    const segs = buildSegments(
      [textRun('lead '), textRun('ab'), hyphen, follower],
      {} as LineLayoutEnvironment,
    );
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 70, 0, 1);
    const texts = lines.map((line) => line.segments
      .map((segment) => (segment as LayoutTextSeg).text)
      .join(''));

    expect(texts[0]).toBe('lead ');
    expect(texts.slice(1).join('')).toBe(`ab-${followerText}`);
    expect(texts.some((line) => line.endsWith('ab') || line.startsWith('-'))).toBe(false);
    expect(texts.some((line) => line.endsWith('-') || line.startsWith('字'))).toBe(false);
  });

  it.each([
    ['CJK', 'A漢', undefined],
    ['SEA', 'Aไทยไทย', [1, 4, 'Aไทยไทย'.length]],
  ] as const)(
    'keeps custom-kinsoku %s leaders with a cross-run noBreakHyphen group',
    (_name, followingText, seaBreaks) => {
      const segment = (
        text: string,
        extra: Partial<LayoutTextSeg> = {},
      ): LayoutTextSeg => ({
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'Times New Roman',
        vertAlign: null,
        measuredWidth: 0,
        ...extra,
      });
      const customKinsoku: KinsokuRules = {
        enabled: true,
        lineStartForbidden: new Set(['A'.codePointAt(0)!]),
        lineEndForbidden: new Set(),
      };
      const segs = [
        segment('lead'),
        segment('ab'),
        segment('-', {
          joinPrev: true,
          hardJoinPrev: true,
          noBreakRanges: [{ start: 0, end: 1 }],
        }),
        segment('字', { joinPrev: true, hardJoinPrev: true }),
        segment(followingText, { seaBreaks }),
      ];
      const { canvas } = makeRecordingCanvas();
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
      const lines = layoutLines(
        ctx,
        segs,
        40,
        0,
        1,
        [],
        undefined,
        {},
        0,
        customKinsoku,
      );
      const texts = lines.map((line) => line.segments
        .map((item) => (item as LayoutTextSeg).text)
        .join(''));

      expect(texts.join('')).toBe(`leadab-字${followingText}`);
      expect(texts).toContain('ab-字A');
      expect(texts.some((line) => line.endsWith('ab') || line.startsWith('-'))).toBe(false);
      expect(texts.some((line) => line.endsWith('-') || line.startsWith('字'))).toBe(false);
      expect(texts.some((line) => line.startsWith('A'))).toBe(false);
      if (_name === 'SEA') {
        // Removing the custom-kinsoku leader must rebase the dictionary offsets
        // from [1, 4, 7] to [3, 6]. Pin the actual reprocessed tail partitions,
        // not only concatenated text, so a stale/intra-grapheme offset is visible.
        expect(texts).toEqual(['lead', 'ab-字A', 'ไทย', 'ไทย']);
      }
    },
  );

  it('clears a consumed hard seam when pagination resumes inside its segment', () => {
    const segment: LayoutTextSeg = {
      text: '字語',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'Times New Roman',
      vertAlign: null,
      measuredWidth: 0,
      joinPrev: true,
      hardJoinPrev: true,
      src: { segIndex: 0, charOffset: 0 },
    };
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(
      ctx, [segment], 20, 0, 1,
      undefined, undefined, undefined, undefined, undefined, undefined,
      undefined, undefined, undefined, undefined, undefined,
      { segIndex: 0, charOffset: 1 },
    );
    const resumed = lines.flatMap((line) => line.segments)
      .find((item): item is LayoutTextSeg => 'text' in item);

    expect(resumed?.text).toBe('語');
    expect(resumed?.joinPrev).toBeUndefined();
    expect(resumed?.hardJoinPrev).toBeUndefined();
  });

  it('an unselected optional hyphen contributes no glyph or gap through parser acquisition', async () => {
    const marker = { ...textRun(''), __optionalHyphen: true } as DocRun;
    const fills = await render([para([textRun('br'), marker, textRun('eaking')])]);
    const drawn = fills.map((c) => c.text).join('');
    expect(drawn).not.toContain('-');
    expect(drawn.replace(/[^a-z]/g, '')).toBe('breaking');
    const whole = await render([para([textRun('breaking')])]);
    expect(fills.at(-1)!.x + fills.at(-1)!.text.length * 10)
      .toBe(whole.at(-1)!.x + whole.at(-1)!.text.length * 10);
  });

  it('selects an authored optional break with its own styled and measured hyphen', () => {
    const marker = { ...textRun(''), optionalHyphen: true, color: 'ff0000', fontSize: 14 } as DocRun;
    const segs = buildSegments([textRun('br'), marker, textRun('eaking')], {} as LineLayoutEnvironment);
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 70, 0, 1);
    expect(lines.map(line => line.segments.filter(segment => 'text' in segment && segment.text)
      .map(segment => (segment as LayoutTextSeg).text))).toEqual([['br', '-'], ['eaking']]);
    const hyphen = lines[0].segments.find(segment => 'text' in segment && segment.text === '-') as LayoutTextSeg;
    expect([hyphen.fontSize, hyphen.color, hyphen.measuredWidth, hyphen.sourceRunIndex])
      .toEqual([14, 'ff0000', 14, 1]);
  });

  it('an unselected styled marker does not interrupt contextual shaping of the word', () => {
    const marker = { ...textRun(''), optionalHyphen: true, color: 'ff0000' } as DocRun;
    const segs = buildSegments([textRun('A'), marker, textRun('V')], {} as LineLayoutEnvironment);
    // A font can kern the pair to fourteen units while separate glyph probes
    // return ten each. Zero-width source formatting must not change that
    // uninterrupted shape or introduce a false overflow at width fifteen.
    const { canvas } = makeRecordingCanvas((text, size) => text === 'AV' ? 14 : [...text].length * size);
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 15, 0, 1);
    expect(lines[0].segments.reduce((sum, segment) => sum + segment.measuredWidth, 0)).toBe(14);
    expect(lines.map(line => line.segments.filter(segment => 'text' in segment && segment.text)
      .map(segment => (segment as LayoutTextSeg).text).join(''))).toEqual(['AV']);
  });

  it('does not select an optional glyph that exceeds the remaining line width', () => {
    const marker = { ...textRun(''), optionalHyphen: true, fontSize: 40 } as DocRun;
    const segs = buildSegments([textRun('lead br'), marker, textRun('eaking')], {} as LineLayoutEnvironment);
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 80, 0, 1);
    const text = lines.map(line => line.segments.filter(segment => 'text' in segment && segment.text)
      .map(segment => (segment as LayoutTextSeg).text));
    expect(text.map(parts => parts.join(''))).toEqual(['lead ', 'breaking']);
  });

  it('selects the last fitting optional owner and resumes after that marker', () => {
    const first = { ...textRun(''), optionalHyphen: true, color: 'ff0000', fontSize: 14 } as DocRun;
    const last = { ...textRun(''), optionalHyphen: true, color: '0000ff' } as DocRun;
    const segs = buildSegments([textRun('ab'), first, textRun('cd'), last, textRun('efgh')], {} as LineLayoutEnvironment);
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, segs, 65, 0, 1);
    expect(lines.map(line => line.segments.filter(segment => 'text' in segment && segment.text)
      .map(segment => (segment as LayoutTextSeg).text).join(''))).toEqual(['abcd-', 'efgh']);
    const glyph = lines[0].segments.find(segment => 'text' in segment && segment.text === '-') as LayoutTextSeg;
    expect([glyph.color, glyph.measuredWidth, glyph.sourceRunIndex]).toEqual(['0000ff', 10, 3]);
  });

  it('bounds dense authored-marker prefix shaping with the production pass quota', () => {
    // The unbroken word fits. Its 5,800 authored opportunities still request
    // over 16 Mi UTF-16 units of prefix work, so the shared pass quota must
    // abort rather than allow an uncharged quadratic scan or partial layout.
    // Supply the compact segment representation directly so this regression
    // isolates the production line-breaking quota from run acquisition work.
    const segs = buildSegments([textRun('x'.repeat(5801))], {} as LineLayoutEnvironment);
    const owner = buildSegments([textRun('x'), { ...textRun(''), optionalHyphen: true } as DocRun,
      textRun('x')], {} as LineLayoutEnvironment)[0] as LayoutTextSeg;
    const word = segs[0] as LayoutTextSeg;
    word.optionalHyphenWord = true;
    word.optionalHyphenBreaks = Array.from({ length: 5800 }, (_, index) => ({
      offset: index + 1, glyph: owner.optionalHyphenBreaks![0].glyph,
    }));
    const { canvas } = makeRecordingCanvas(text => text.length);
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    expect(() => layoutLines(ctx, segs, 6000, 0, 1))
      .toThrow('DOCX line-break prefix search exceeded its UTF-16 work quota');
  });

  function intrinsicWidths(runs: DocRun[], kinsoku: KinsokuRules) {
    const { canvas } = makeRecordingCanvas();
    const context = {
      lineGrid: { active: false, pitchPt: null },
      characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
      physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
      lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
      baseRtl: false, isJustified: false, stretchLastLine: false,
      tabStops: [], hasRuby: false, hasEastAsianText: false, kinsoku, defaultTabPt: 36,
    } as ParagraphLayoutContext;
    return measureParagraphIntrinsicWidths(para(runs), context, 100000,
      { context: canvas.getContext('2d') as unknown as CanvasRenderingContext2D, fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false });
  }

  const optionalRules: KinsokuRules = { enabled: true, lineStartForbidden: new Set([41]), lineEndForbidden: new Set() };

  it('intrinsic optional breaks respect the following authored hard seam', () => {
    const runs = [textRun('ab'), { ...textRun(''), optionalHyphen: true } as DocRun,
      { ...textRun('-cd'), noBreakBefore: true, noBreakRanges: [{ start: 0, end: 1 }] } as DocRun];
    expect(intrinsicWidths(runs, optionalRules)).toEqual({ minWidthPt: 50, maxWidthPt: 50 });
  });

  it('intrinsic optional breaks respect a forbidden continuation line start', () => {
    const runs = [textRun('abcd'), { ...textRun(''), optionalHyphen: true } as DocRun, textRun(')efgh')];
    expect(intrinsicWidths(runs, optionalRules)).toEqual({ minWidthPt: 90, maxWidthPt: 90 });
  });

  it('intrinsic token traversal does not rescan all authored marker offsets', () => {
    const runs: DocRun[] = [];
    for (let index = 0; index < 200; index += 1)
      runs.push(textRun('ab'), { ...textRun(''), optionalHyphen: true } as DocRun, textRun('cd '));
    // Count actual full-map visits rather than using an elapsed-time heuristic.
    // This word/space model executes the public intrinsic entry point; 200
    // tokens must not enumerate their complete 200-marker map 200 times.
    const keys = Map.prototype.keys;
    let markerVisits = 0;
    let mergedMarkerCopies = 0;
    const arrayIterator = Array.prototype[Symbol.iterator];
    // A plain override avoids spy bookkeeping recursively iterating arrays.
    Array.prototype[Symbol.iterator] = function (this: unknown[]) {
      const iterator = arrayIterator.call(this);
      return (function* () {
        for (const value of iterator) {
          if (value && typeof value === 'object' && 'offset' in value && 'glyph' in value) mergedMarkerCopies += 1;
          yield value;
        }
      })() as ArrayIterator<unknown>;
    };
    const spy = vi.spyOn(Map.prototype, 'keys').mockImplementation(function (this: Map<unknown, unknown>) {
      const iterator = keys.call(this);
      return (function* () {
        for (const key of iterator) {
          if (typeof key === 'number' && key > 0) markerVisits += 1;
          yield key;
        }
      })() as MapIterator<unknown>;
    });
    try {
      expect(intrinsicWidths(runs, optionalRules).minWidthPt).toBe(30);
      expect(markerVisits).toBeLessThanOrEqual(400);
      expect(mergedMarkerCopies).toBeLessThanOrEqual(1200);
    } finally { Array.prototype[Symbol.iterator] = arrayIterator; spy.mockRestore(); }
  });

  it('selects the authored marker at a real font and script seam', () => {
    const runs = [{ ...textRun('ab'), fontSize: 20 } as DocRun,
      { ...textRun(''), optionalHyphen: true, color: 'ff0000' } as DocRun,
      { ...textRun('日本'), fontFamily: 'Other Face', fontSize: 20 } as DocRun];
    const { canvas } = makeRecordingCanvas();
    const lines = layoutLines(canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
      buildSegments(runs, {} as LineLayoutEnvironment), 50, 0, 1);
    expect(lines.map(line => line.segments.filter(s => 'text' in s && s.text)
      .map(s => (s as LayoutTextSeg).text).join(''))).toEqual(['ab-', '日本']);
  });

  it('clears consumed standalone-marker joins on the continuation line', () => {
    const runs = [{ ...textRun('ab'), fitTextVal: 400, fitTextId: '1' } as DocRun,
      { ...textRun(''), optionalHyphen: true, color: 'ff0000' } as DocRun,
      { ...textRun(''), optionalHyphen: true, color: '0000ff' } as DocRun, textRun('cd')];
    const { canvas } = makeRecordingCanvas();
    const lines = layoutLines(canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
      buildSegments(runs, {} as LineLayoutEnvironment), 35, 0, 1);
    expect(lines.map(line => line.segments.filter(s => 'text' in s && s.text)
      .map(s => (s as LayoutTextSeg).text).join(''))).toEqual(['ab-', 'cd']);
    const continuation = lines[1].segments.find(s => 'text' in s) as LayoutTextSeg;
    expect(continuation.joinPrev).toBeUndefined();
    expect(continuation.hardJoinPrev).toBeUndefined();
  });

  it('an unselected large optional owner leaves ordinary and standalone line metrics unchanged', () => {
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const marker = { ...textRun(''), optionalHyphen: true, fontSize: 80 } as DocRun;
    const measure = (runs: DocRun[]) => layoutLines(ctx,
      buildSegments(runs, {} as LineLayoutEnvironment), 200, 0, 1)
      .map(line => [line.height, line.ascent, line.descent]);
    expect(measure([textRun('br'), marker, textRun('eaking')])).toEqual(measure([textRun('breaking')]));
    const fixed = { ...textRun('ab'), fitTextVal: 400, fitTextId: '1' } as DocRun;
    expect(measure([fixed, marker, textRun('cd')])).toEqual(measure([fixed, textRun('cd')]));
  });

  it('an unselected optional owner retains the existing registered-face space fit', () => {
    const route = createCanvasFontRoute('Synthetic Latin', 'registered');
    const word = (text: string): LayoutTextSeg => {
      const parsed = textRun(text);
      return {
        text: parsed.text, bold: parsed.bold, italic: parsed.italic,
        underline: parsed.underline, strikethrough: parsed.strikethrough,
        fontSize: parsed.fontSize, color: parsed.color, vertAlign: parsed.vertAlign,
        fontFamily: 'Synthetic Latin', fontRoute: route, measuredWidth: 0,
        latinSpaceAverageWidthRatio: 0.4, latinSpaceCompressionEligible: true,
      };
    };
    const { canvas } = makeRecordingCanvas(text => [...text].reduce((sum, c) => sum + (c === ' ' ? 4 : 2), 0));
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const plain = layoutLines(ctx, [word('A '), word('BC')], 8, 0, 1);
    expect(plain.map(line => line.segments.map(s => s.measuredWidth))).toEqual([[4, 4]]);
    const authored = { ...word('BC'), optionalHyphenWord: true as const,
      optionalHyphenBreaks: [{ offset: 1, glyph: word('-') }] };
    const withMarker = layoutLines(ctx, [word('A '), authored], 8, 0, 1);
    expect(withMarker.map(line => line.segments.map(s => s.measuredWidth))).toEqual([[4, 4]]);
  });
});


describe('authored optional glyph logical text boundary', () => {
  function selectedDefaultMarkerLayout() {
    const marker = { ...textRun(''), __optionalHyphen: true } as DocRun;
    const model = doc([para([textRun('br'), marker, textRun('eaking-keep')], { right: 230 })]);
    const { canvas } = makeRecordingCanvas();
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    }), { currentDateMs: 0 });
    const paragraph = layout.pages[0]!.layers.body.find(node => node.kind === 'paragraph');
    if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Expected acquired paragraph');
    return { layout, paragraph };
  }

  it('retains selected optional ink when upright-frame paint operations are projected from logical ranges', () => {
    const { paragraph } = selectedDefaultMarkerLayout();
    const glyph = paragraph.lines.flatMap(line => line.placements)
      .find(placement => placement.kind === 'text' && placement.optionalHyphenGlyph);
    if (!glyph || glyph.kind !== 'text') throw new Error('Expected selected glyph');
    const operations = eastAsianUprightPaintOps(glyph);
    expect(operations).toMatchObject([{ text: '-', range: { start: 2, end: 2 }, glyphOrientation: 'sideways' }]);
    expect(operations[0]?.offset).toEqual(glyph.clusters[0]?.offset);
    expect(glyph.advancePt).toBe(10);
  });

  it('omits selected optional ink from shape element context while retaining literal hyphens', () => {
    const { layout, paragraph } = selectedDefaultMarkerLayout();
    const bounds = { xPt: 0, yPt: 0, widthPt: 70, heightPt: 200 };
    const box: import('./layout/types.js').TextBoxLayout = {
      kind: 'textbox', id: 'optional-box', source: { story: 'textbox', storyInstance: 'optional-box', path: [] },
      flowDomainId: 'optional-box', flowBounds: bounds, inkBounds: bounds, advancePt: paragraph.advancePt,
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      ordinaryFlow: false, writingMode: 'horizontal-tb', insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      story: { story: 'textbox', flowBounds: bounds, inkBounds: bounds, blocks: [paragraph],
        advancePt: paragraph.advancePt, diagnostics: [] },
    };
    const drawing: import('./layout/types.js').DrawingLayout = {
      kind: 'drawing', id: 'optional-shape', source: paragraph.source, flowDomainId: paragraph.flowDomainId,
      flowBounds: bounds, inkBounds: bounds, advancePt: 0, ordinaryFlow: false,
      commands: [{ kind: 'fill-rect', rect: bounds, fill: '#ffffff' }], textBoxIds: [box.id],
    };
    const host = { ...paragraph, lines: [], drawings: [drawing], textBoxes: [box] };
    const retained = { ...layout, pages: [{ ...layout.pages[0]!,
      layers: buildPageLayers([{ layer: 'body', node: host }]) }] };
    const context = hitTestDocxElementContext(retained, 0, { xPt: 1, yPt: 1 }, createPaintResourceRegistry([]));
    expect(context?.elementType).toBe('shape');
    expect(context?.text).toBe('br\neaking-\nkeep');
  });
  it.each([{ name: 'same-format', color: null, fontSize: 12 },
    { name: 'real-format', color: 'ff0000', fontSize: 14 }])(
    'retains zero logical ranges and complete $name source ownership at a selected optional glyph', async style => {
    const run = (text: string): Extract<DocRun, { type: 'text' }> => ({ ...textRun(text), fontSize: 12 });
    const marker = { ...run(''), ...style, __optionalHyphen: true };
    const model = doc([para([run('br'), marker, run('eaking-keep')], { right: 216 })]);
    const { canvas, fills } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
    const services = createLayoutServices(model, { measureContext: ctx });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const owned = textRunGeometryForPage(layout, 0).map(item => item.placement);
    const glyph = owned.find(placement => placement.optionalHyphenGlyph);
    expect(glyph).toMatchObject({ text: '-', sourceRunIndex: 1,
      range: { start: 2, end: 2 }, advancePt: style.fontSize });
    expect(glyph?.clusters).toMatchObject([{ range: { start: 2, end: 2 }, advancePt: style.fontSize }]);
    expect(owned.filter(placement => placement.sourceRunIndex === 0).map(placement => placement.text).join('')).toBe('br');
    const suffix = owned.filter(placement => placement.sourceRunIndex === 2);
    expect(suffix.map(placement => placement.text).join('')).toBe('eaking-keep');
    expect(suffix[0]?.range.start).toBe(2);
    expect(suffix.at(-1)?.range.end).toBe(13);
    const wideModel = doc([para([run('br'), marker, run('eaking-keep')])]);
    const wide = layoutDocument(wideModel, createLayoutServices(wideModel, { measureContext: ctx }), { currentDateMs: 0 });
    const wideOwned = textRunGeometryForPage(wide, 0).map(item => item.placement);
    expect(wideOwned.find(placement => placement.sourceRunIndex === 1)).toMatchObject({
      text: '', range: { start: 2, end: 2 }, advancePt: 0 });
    expect(wideOwned.find(placement => placement.sourceRunIndex === 2)).toMatchObject({
      text: 'eaking-keep', range: { start: 2, end: 13 } });
    const bodyParagraph = layout.pages[0]!.layers.body.find(node => node.kind === 'paragraph');
    expect(bodyParagraph?.kind).toBe('paragraph');
    if (bodyParagraph?.kind === 'paragraph') {
      expect(bodyParagraph.lines[0]?.range).toEqual({ start: 0, end: 2 });
      expect(bodyParagraph.lines.at(-1)?.range.end).toBe(13);
    }
    // Every suffix cluster retains its original UTF-16 interval across wrap.
    expect(suffix.flatMap(placement => placement.clusters.map(cluster => cluster.range)))
      .toEqual(wideOwned.filter(placement => placement.sourceRunIndex === 2)
        .flatMap(placement => placement.clusters.map(cluster => cluster.range)));
    const projected: import('./types.js').DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(model, canvas, 0,
      { dpr: 1, width: PAGE_W, layoutServices: services, currentDate: 0, onTextRun: item => projected.push(item) });
    expect(fills.map(call => call.text).join('')).toBe('br-eaking-keep');
    const index = buildTextIndex(projected);
    expect(index.text).toBe('breaking-keep');
    expect(findMatches(index, 'breaking')).toHaveLength(1);
    expect(findMatches(index, '-keep')).toHaveLength(1);
  });
  it('keeps selected optional glyph ink out of logical copy and word find while retaining literal hyphens', async () => {
    const marker = { ...textRun(''), __optionalHyphen: true, color: 'ff0000', fontSize: 14 } as DocRun;
    const { canvas, fills } = makeRecordingCanvas();
    const projected: import('./types.js').DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(doc([para([textRun('br'), marker, textRun('eaking-keep')], { right: 230 })]),
      canvas, 0, { dpr: 1, width: PAGE_W, onTextRun: run => projected.push(run) });
    expect(fills.map(call => call.text).join('')).toBe('br-eaking-keep');
    expect(projected.map(run => run.text).join('')).toBe('breaking-keep');
    const index = buildTextIndex(projected);
    expect(index.text).toBe('breaking-keep');
    expect(findMatches(index, 'breaking')).toHaveLength(1);
    expect(findMatches(index, 'br-eaking')).toHaveLength(0);
    expect(findMatches(index, '-keep')).toHaveLength(1);
    expect(projected.find(run => run.sourceRunIndex === 1)).toMatchObject({ text: '', optionalHyphenGlyph: true });
    expect(projected.filter(run => run.sourceRunIndex === 2).map(run => run.text).join('')).toBe('eaking-keep');
  });
});
