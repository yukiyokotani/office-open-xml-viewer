import { describe, it, expect } from 'vitest';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { renderDocumentToCanvas } from './renderer.js';
import {
  splitTextForLayout,
  layoutLines,
  buildSegments,
  type LayoutTextSeg,
  type LineLayoutEnvironment,
} from './line-layout.js';
import type { DocParagraph, DocxDocumentModel, SectionProps, DocRun } from './types.js';

// ECMA-376 §17.3.3 run-content elements that were previously dropped by the
// parser's `_ => {}` arm and thus never reached the renderer:
//   §17.3.3.23 <w:ptab>        — absolute-position tab
//   §17.3.3.18 <w:noBreakHyphen> — non-breaking hyphen glyph
//   §17.3.3.29 <w:softHyphen>  — optional hyphen (invisible without hyphenation)
// These end-to-end tests record fillText() calls to pin the layout geometry the
// parser + line-layout now produce. Scale is 1 px/pt (canvas width == pageWidth)
// and every glyph is FS px wide in the recording canvas.

interface FillCall {
  text: string;
  x: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
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
        width: [...s].length * p,
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

function textRun(text: string): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocRun;
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

  it('right ptab relative to MARGIN ignores indents and aligns to the page margin', async () => {
    const fills = await render([
      para([ptabRun('right', 'margin'), textRun('99')], { left: 40, right: 20 }),
    ]);
    const f = fills.find((c) => c.text === '99');
    expect(f, '"99" must be drawn').toBeDefined();
    // relativeTo="margin" ⇒ right edge on the text margin (PAGE_W = 300), NOT the
    // indented content box; 2-glyph number starts at 300 − 20 = 280.
    expect(f!.x + 2 * FS).toBeCloseTo(PAGE_W, 3);
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
    const segs = buildSegments([textRun('lead 999-99')], {} as LineLayoutEnvironment);
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

  // §17.3.3.29: a soft hyphen "shall have zero width" and "shall not change
  // the normal display of text" unless it is the chosen break point; since
  // this renderer performs no automatic hyphenation (§17.15.1.x), it is never
  // chosen, so state (a) always applies. The parser reflects this by emitting
  // NOTHING for <w:softHyphen/> (`soft_hyphen_is_invisible`, parser.rs) — this
  // test exercises the RENDERER side of that same contract: given the exact
  // shape the parser produces for "br"+softHyphen+"eaking" (two adjacent text
  // runs, nothing in between), the renderer must draw a contiguous "breaking"
  // with no synthesized hyphen and no extra gap between the pieces.
  it('softHyphen contributes no glyph and no gap, given the shape the parser actually emits', async () => {
    // This is what parse_run_inner produces for
    // <w:r><w:t>br</w:t><w:softHyphen/><w:t>eaking</w:t></w:r>: the
    // <w:softHyphen/> arm pushes nothing, so exactly two DocRun::Text survive.
    const fills = await render([para([textRun('br'), textRun('eaking')])]);
    const drawn = fills.map((c) => c.text).join('');
    expect(drawn).not.toContain('-');
    expect(drawn.replace(/[^a-z]/g, '')).toBe('breaking');
    // No gap: the two pieces are adjacent glyph runs, not separated by a
    // dropped-but-still-spaced placeholder. "eaking" must start exactly where
    // "br" ends (2 glyphs * FS), not further right.
    const br = fills.find((c) => c.text === 'br');
    const eaking = fills.find((c) => c.text === 'eaking');
    expect(br, '"br" must be drawn').toBeDefined();
    expect(eaking, '"eaking" must be drawn').toBeDefined();
    expect(eaking!.x).toBeCloseTo(br!.x + 2 * 10, 3);
  });
});
