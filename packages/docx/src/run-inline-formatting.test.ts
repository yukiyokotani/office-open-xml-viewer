import { describe, it, expect } from 'vitest';
import { crispOffset } from '@silurus/ooxml-core';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  DocxRunBorder,
  SectionProps,
} from './types';

// End-to-end renderer verification of run-level box (`<w:bdr>` §17.3.2.4) and
// shading (`<w:shd w:fill>` §17.3.2.32) draw geometry, plus the spec's
// run-border GROUPING rule: adjacent runs whose border attribute set is
// identical render within a single frame (§17.3.2.4). These are checked through
// a recording 2D context that captures every fillRect / retained stroke edge /
// fillText in draw order with its geometry and style. (autoContrastColor's unit tests
// moved to packages/core/src/shape/paint.test.ts when the function was lifted
// into core.)

type DrawEvent =
  | { kind: 'fillRect'; x: number; y: number; w: number; h: number; style: string }
  | { kind: 'strokeLine'; x1: number; y1: number; x2: number; y2: number; style: string; lineWidth: number }
  | { kind: 'fillText'; text: string; x: number };

interface RetainedFrame {
  left: number;
  top: number;
  right: number;
  bottom: number;
  style: string;
  lineWidth: number;
}

function retainedFrames(events: readonly DrawEvent[]): RetainedFrame[] {
  const edges = events.filter(
    (event): event is Extract<DrawEvent, { kind: 'strokeLine' }> => event.kind === 'strokeLine',
  );
  expect(edges.length % 4).toBe(0);
  const frames: RetainedFrame[] = [];
  for (let index = 0; index < edges.length; index += 4) {
    const [top, right, bottom, left] = edges.slice(index, index + 4);
    expect(top.y1).toBeCloseTo(top.y2);
    expect(right.x1).toBeCloseTo(right.x2);
    expect(bottom.y1).toBeCloseTo(bottom.y2);
    expect(left.x1).toBeCloseTo(left.x2);
    const snapLeft = crispOffset(top.x1, left.lineWidth, 1);
    const snapRight = crispOffset(top.x2, right.lineWidth, 1);
    const snapTop = crispOffset(left.y1, top.lineWidth, 1);
    const snapBottom = crispOffset(left.y2, bottom.lineWidth, 1);
    expect(left.x1).toBeCloseTo(top.x1 + snapLeft);
    expect(right.x1).toBeCloseTo(top.x2 + snapRight);
    expect(left.x2).toBeCloseTo(bottom.x1 + snapLeft);
    expect(right.x2).toBeCloseTo(bottom.x2 + snapRight);
    expect(top.y1).toBeCloseTo(left.y1 + snapTop);
    expect(bottom.y1).toBeCloseTo(left.y2 + snapBottom);
    expect(top.style).toBe(right.style);
    expect(top.style).toBe(bottom.style);
    expect(top.style).toBe(left.style);
    frames.push({
      left: top.x1,
      top: left.y1,
      right: top.x2,
      bottom: left.y2,
      style: top.style,
      lineWidth: top.lineWidth,
    });
  }
  return frames;
}

/** Recording 2D context. Glyph advance = charCount × fontPx; font box 0.8/0.2
 *  em — the same synthetic metrics the numbering-marker test uses. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  events: DrawEvent[];
} {
  let font = '16px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  const events: DrawEvent[] = [];
  let path: { x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    save() {}, restore() {}, beginPath() { path = []; }, closePath() {},
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() {
      for (let index = 1; index < path.length; index += 1) {
        const from = path[index - 1];
        const to = path[index];
        events.push({
          kind: 'strokeLine', x1: from.x, y1: from.y, x2: to.x, y2: to.y,
          style: String(this.strokeStyle), lineWidth: this.lineWidth,
        });
      }
    }, fill() {}, clip() {}, rect() {},
    scale() {}, translate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillRect(x: number, y: number, w: number, h: number) {
      events.push({ kind: 'fillRect', x, y, w, h, style: String(this.fillStyle) });
    },
    strokeRect() {},
    fillText(text: string, x: number, _y: number) {
      events.push({ kind: 'fillText', text, x });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, events };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 16, color: null, fontFamily: 'Arial', fontFamilyEastAsia: 'Arial',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  };
}

function paraDoc(
  runs: DocxTextRun[],
  paraExtra: Partial<DocParagraph> = {},
  pageWidth = 400,
  sectionExtra: Partial<SectionProps> = {},
): DocxDocumentModel {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocParagraph['runs'][number]),
    defaultFontSize: 16, defaultFontFamily: 'Arial',
    widowControl: false,
    ...paraExtra,
  };
  return {
    section: {
      pageWidth, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      ...sectionExtra,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    // Arial → swiss (sans) so the segment stays single-font; not load-bearing
    // for the rect geometry these tests assert.
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

async function render(
  runs: DocxTextRun[],
  paraExtra: Partial<DocParagraph> = {},
  pageWidth = 400,
  sectionExtra: Partial<SectionProps> = {},
) {
  const { canvas, events } = makeRecordingCanvas();
  await renderDocumentToCanvas(paraDoc(runs, paraExtra, pageWidth, sectionExtra), canvas, 0, {
    dpr: 1,
    width: pageWidth, // scale = 1 (px per pt) so geometry is in pt-equivalent px
  });
  return events;
}

const border = (extra: Partial<DocxRunBorder> = {}): DocxRunBorder => ({
  style: 'single', color: '0000FF', width: 1, space: 0, ...extra,
});

describe('run box (w:bdr §17.3.2.4) + shading (w:shd §17.3.2.32) geometry', () => {
  it('insets the box outside the glyph box by w:space', async () => {
    // One run carrying BOTH shading (fillRect at the glyph box) and a border
    // with w:space — the four retained edges must sit `space*scale` OUTSIDE the
    // shading rect on every side (box bounds = glyph box + space inset). Select the RUN
    // shading rect by its colour (the first fillRect is the page-white bg).
    const sp = 4;
    const events = await render([
      textRun('AB', { background: '00FF00', border: border({ space: sp }) }),
    ]);
    const fill = events.find((e) => e.kind === 'fillRect' && e.style.toUpperCase() === '#00FF00');
    const [stroke] = retainedFrames(events);
    expect(fill).toBeDefined();
    expect(stroke).toBeDefined();
    if (fill?.kind !== 'fillRect' || !stroke) throw new Error('unreachable');
    // scale = 1 ⇒ inset = sp px on each side.
    expect(stroke.left).toBeCloseTo(fill.x - sp);
    expect(stroke.top).toBeCloseTo(fill.y - sp);
    expect(stroke.right - stroke.left).toBeCloseTo(fill.w + 2 * sp);
    expect(stroke.bottom - stroke.top).toBeCloseTo(fill.h + 2 * sp);
    // The box is the run's border colour.
    expect(stroke.style.toUpperCase()).toBe('#0000FF');
  });

  it('merges two adjacent runs with an identical w:bdr into one frame (§17.3.2.4)', async () => {
    // Two adjacent runs, identical border, no shading. The spec: identical
    // adjacent borders form one group rendered within a single set of borders.
    const events = await render([
      textRun('AB', { border: border() }),
      textRun('CD', { border: border() }),
    ]);
    const strokes = retainedFrames(events);
    expect(strokes).toHaveLength(1); // ONE four-edge frame, not one per run
    const stroke = strokes[0];
    // The frame spans both runs: width ≈ |AB| + |CD| = 2 chars × 16px × 2 runs.
    // space = 0 ⇒ no inset. Each char advance is 16px (synthetic metrics).
    expect(stroke.right - stroke.left).toBeCloseTo(4 * 16);
    // First glyph of the first run starts at the frame's left edge.
    const firstText = events.find((e) => e.kind === 'fillText');
    if (firstText?.kind !== 'fillText') throw new Error('unreachable');
    expect(stroke.left).toBeCloseTo(firstText.x);
  });

  it('does NOT merge two adjacent runs whose borders differ (separate frames)', async () => {
    const events = await render([
      textRun('AB', { border: border({ color: '0000FF' }) }),
      textRun('CD', { border: border({ color: 'FF0000' }) }),
    ]);
    const strokes = retainedFrames(events);
    expect(strokes).toHaveLength(2); // different colour ⇒ two groups
  });

  it('paints the shading fillRect behind the text (before the glyphs)', async () => {
    const events = await render([
      textRun('AB', { background: 'C0C0C0' }),
    ]);
    // Select the RUN shading rect by colour (the page-white bg fillRect runs
    // first and would otherwise win findIndex).
    const fillIdx = events.findIndex((e) => e.kind === 'fillRect' && e.style.toUpperCase() === '#C0C0C0');
    const textIdx = events.findIndex((e) => e.kind === 'fillText' && e.text.includes('A'));
    expect(fillIdx).toBeGreaterThanOrEqual(0);
    expect(textIdx).toBeGreaterThanOrEqual(0);
    // Shading is drawn FIRST so the glyphs sit on top of the fill.
    expect(fillIdx).toBeLessThan(textIdx);
    const fill = events[fillIdx];
    if (fill.kind !== 'fillRect') throw new Error('unreachable');
    // …in the run's shading colour, at the glyph box.
    expect(fill.style.toUpperCase()).toBe('#C0C0C0');
    expect(fill.w).toBeCloseTo(2 * 16); // |AB| = 2 chars × 16px
  });
});

describe('highlight fill spans justification slack (§17.3.2.15 highlight + §17.18.44 both)', () => {
  it('uses the selected font box instead of the paragraph line advance', async () => {
    const events = await render(
      [textRun('Highlighted', { highlight: 'black' })],
      { lineSpacing: { value: 30, rule: 'exact', explicit: true } },
    );
    const fill = events.find(
      (event): event is Extract<DrawEvent, { kind: 'fillRect' }> =>
        event.kind === 'fillRect' && event.style.toUpperCase() === '#000000',
    );

    expect(fill).toBeDefined();
    expect(fill?.h).toBeCloseTo(16);
  });

  it('tiles the highlight with no gaps across justified inter-word spaces', async () => {
    // A justified ('both') paragraph that wraps to 2+ lines. Line 0 is justified,
    // so its inter-word spaces are expanded by the per-gap slack. Word highlights
    // the run's spaces (incl. the expansion); our per-word highlight rects must
    // therefore TILE line 0 contiguously. The bug: each rect spans only the word's
    // natural advance (`measuredWidth`), leaving the expanded space unpainted —
    // a visible yellow gap between words.
    const events = await render(
      [textRun('aaaaa bbbbb ccccc ddddd eeeee fffff ggggg', { highlight: 'yellow' })],
      { alignment: 'both' },
      410, // not an exact multiple of the word advance, so line 0 carries slack
    );
    const yellow = events.filter(
      (e): e is Extract<DrawEvent, { kind: 'fillRect' }> =>
        e.kind === 'fillRect' && e.style.toUpperCase() === '#FFFF00',
    );
    expect(yellow.length).toBeGreaterThan(2);
    // Group rects by line via their top (y). The first line is the smallest y.
    const ys = [...new Set(yellow.map((r) => Math.round(r.y)))].sort((a, b) => a - b);
    expect(ys.length).toBeGreaterThanOrEqual(2); // wrapped to ≥2 lines
    const line0 = yellow
      .filter((r) => Math.round(r.y) === ys[0])
      .sort((a, b) => a.x - b.x);
    expect(line0.length).toBeGreaterThan(1);
    // Justification guard (fix-stable: keyed off glyph x, which the highlight-width
    // fix does NOT change): the START-to-START distance between the first two words
    // exceeds the first word's natural advance — i.e. the inter-word space really
    // is expanded on line 0. Were the line left-aligned, these would be equal.
    const line0Words = events
      .filter((e): e is Extract<DrawEvent, { kind: 'fillText' }> => e.kind === 'fillText')
      .filter((e) => /a{5}/.test(e.text)) // first word starts with aaaaa
      .sort((a, b) => a.x - b.x);
    const word0 = events.find(
      (e): e is Extract<DrawEvent, { kind: 'fillText' }> => e.kind === 'fillText',
    );
    if (!word0) throw new Error('no glyphs drawn');
    const naturalWord0 = [...word0.text].length * 16; // synthetic metrics: 16px/char
    expect(line0[1].x - line0[0].x).toBeGreaterThan(naturalWord0); // gap WAS expanded
    // Contiguity: every rect's right edge meets the next rect's left edge, so the
    // expanded space is fully painted (no yellow gap between words).
    for (let i = 0; i < line0.length - 1; i++) {
      expect(line0[i].x + line0[i].w).toBeCloseTo(line0[i + 1].x, 5);
    }
  });
});

describe('run border spans justification slack (§17.3.2.4 + §17.18.44 both)', () => {
  it('extends the border frame across justified inter-word slack', async () => {
    // The run-border frame goes through the group-accumulation path (not the
    // fillRect path the highlight test exercises), so guard it separately: on a
    // justified line the frame's right edge must fold in the widened inter-word
    // space (decoW), making the line-0 frame WIDER than the same text left-
    // aligned (where there is no slack). Before the fix both used the natural
    // glyph span and were equal.
    const text = 'aaaaa bbbbb ccccc ddddd eeeee fffff ggggg';
    const firstLineFrameWidth = async (alignment: 'both' | 'left') => {
      const events = await render([textRun(text, { border: border() })], { alignment }, 410);
      const strokes = retainedFrames(events);
      expect(strokes.length).toBeGreaterThanOrEqual(2); // wrapped to ≥2 lines ⇒ ≥2 frames
      const top = Math.min(...strokes.map((s) => s.top));
      return Math.max(...strokes
        .filter((s) => Math.round(s.top) === Math.round(top))
        .map((s) => s.right - s.left));
    };
    const justified = await firstLineFrameWidth('both');
    const leftAligned = await firstLineFrameWidth('left');
    // The justified line-0 frame absorbs the inter-word slack and is strictly wider.
    expect(justified).toBeGreaterThan(leftAligned + 1);
  });
});

describe('over-long word overflow-wrap (long URLs in a narrow column)', () => {
  it('keeps an unwrapped hyperlink as one contextual draw', async () => {
    const url = 'https://example.com/path/a-short-name.pdf';
    const events = await render([
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 800);
    const pieces = events
      .filter((event): event is Extract<DrawEvent, { kind: 'fillText' }> => event.kind === 'fillText');
    expect(pieces.map((event) => event.text)).toEqual([url]);
  });

  it('recognizes URL syntax across a formatting seam in the scheme and authority', async () => {
    const url = 'https://example.com/path/a-very-long-document-name.pdf';
    const events = await render([
      textRun('lead '),
      textRun('https', { isLink: true, hyperlink: url }),
      textRun('://example.com/path/a-very-long-document-name.pdf', {
        isLink: true,
        hyperlink: url,
        bold: true,
      }),
    ], {}, 480);
    const pieces = events
      .filter((event): event is Extract<DrawEvent, { kind: 'fillText' }> => event.kind === 'fillText')
      .filter((event) => event.text !== 'lead ');
    expect(pieces.map((event) => event.text).join('')).toBe(url);
    expect(pieces[0]?.x).toBeCloseTo(80);
    expect(pieces.map((event) => event.text).join('')).toContain('example.com');
  });

  it('uses a URL path boundary in the current line remainder', async () => {
    const url = 'https://example.com/path/a-very-long-document-name.pdf';
    const events = await render([
      textRun('lead '),
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 480);
    const texts = events.filter((event) => event.kind === 'fillText');
    const urlPieces = texts.filter((event) => event.text !== 'lead ');

    expect(urlPieces.map((event) => event.text).join('')).toBe(url);
    // The 80px "lead " leaves exactly 400px: enough for the first complete
    // path segment. Use that readable boundary instead of leaving a
    // conspicuous remainder or splitting arbitrarily inside a host.
    expect(urlPieces[0]?.x).toBeCloseTo(80);
    expect(urlPieces[0]?.text).toBe('https://example.com/path/');
    for (const piece of urlPieces) {
      expect(piece.x + [...piece.text].length * 16).toBeLessThanOrEqual(480 + 1e-6);
    }
  });

  it('keeps opening punctuation with a grapheme-safe prefix instead of overflowing the hyperlink', async () => {
    const url = 'https://example.com/a-very-long-document-name.pdf';
    const events = await render([
      textRun('('),
      textRun(url, { isLink: true, hyperlink: url }),
      textRun(')'),
    ], {}, 368);
    const texts = events.filter((event) => event.kind === 'fillText');
    const urlPieces = texts.filter((event) => event.text !== '(' && event.text !== ')');

    expect(urlPieces.map((event) => event.text).join('')).toBe(url);
    // UAX #14 LB14 keeps the first URL character with "(". That protected
    // seam must not turn the complete hyperlink run into one over-wide draw.
    expect(urlPieces[0]?.x).toBeCloseTo(16);
    expect(urlPieces[0]?.text).toBe('https://example.com/a-');
    expect(urlPieces.length).toBeGreaterThan(1);
    for (const piece of urlPieces) {
      expect(piece.x + [...piece.text].length * 16).toBeLessThanOrEqual(368 + 1e-6);
    }
  });

  it('splits the real glued group when the hyperlink alone exactly fills the line', async () => {
    const url = 'https://example.com/a-very-long-document-name.pdf';
    const events = await render([
      textRun('('),
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 352);
    const pieces = events
      .filter((event): event is Extract<DrawEvent, { kind: 'fillText' }> => event.kind === 'fillText')
      .filter((event) => event.text !== '(');
    expect(pieces.map((event) => event.text).join('')).toBe(url);
    expect(pieces[0]?.x).toBeCloseTo(16);
    expect(pieces[0]!.x + [...pieces[0]!.text].length * 16).toBeLessThanOrEqual(352);
  });

  it('moves a hyperlink when no readable URL boundary fits the current remainder', async () => {
    const url = 'https://example.com/a-very-long-document-name.pdf';
    const events = await render([
      textRun('prefix '),
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 320);
    const texts = events.filter((event) => event.kind === 'fillText');
    const urlPieces = texts.filter((event) => event.text !== 'prefix ');

    expect(urlPieces.map((event) => event.text).join('')).toBe(url);
    expect(urlPieces[0]?.x).toBeCloseTo(0);
  });

  it('does not use the current remainder for an ordinary overlong token', async () => {
    const word = 'abcdefghijklmnopqrstuvwxyz0123456789';
    const events = await render([
      textRun('lead '),
      textRun(word),
    ], {}, 160);
    const texts = events.filter((event) => event.kind === 'fillText');
    const wordPieces = texts.filter((event) => event.text !== 'lead ');

    expect(wordPieces.map((event) => event.text).join('')).toBe(word);
    expect(wordPieces[0]?.x).toBeCloseTo(0);
  });

  it('moves an overlong hyperlink when the current remainder cannot hold one grapheme', async () => {
    const url = 'https://example.com/a-very-long-document-name.pdf';
    const events = await render([
      textRun('12345678 '), // 144px, leaving only 8px in the 152px band
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 152);
    const texts = events.filter((event) => event.kind === 'fillText');
    const urlPieces = texts.filter((event) => event.text !== '12345678 ');

    expect(urlPieces.map((event) => event.text).join('')).toBe(url);
    expect(urlPieces[0]?.x).toBeCloseTo(0);
    for (const piece of urlPieces) {
      expect(piece.x + [...piece.text].length * 16).toBeLessThanOrEqual(152 + 1e-6);
    }
  });

  it('evaluates every URL boundary with negative character spacing', async () => {
    const url = 'https://example.com/path/a-long-document-name.pdf';
    const events = await render([
      textRun('lead '),
      textRun(url, { isLink: true, hyperlink: url, charSpacing: -2 }),
    ], {}, 350);
    const pieces = events
      .filter((event): event is Extract<DrawEvent, { kind: 'fillText' }> => event.kind === 'fillText')
      .filter((event) => event.text !== 'lead ');
    expect(pieces.map((event) => event.text).join('')).toBe(url);
    expect(pieces.length).toBeGreaterThan(1);
    expect(pieces.every((piece) => piece.text.length > 0)).toBe(true);
  });

  it('evaluates URL boundaries against the active snapToChars Latin block', async () => {
    const url = 'https://example.com/path/a-long-document-name.pdf';
    const events = await render([
      textRun('lead'),
      textRun(url, { isLink: true, hyperlink: url }),
    ], {}, 220, {
      docGridType: 'snapToChars',
      docGridLinePitch: 20,
      docGridCharSpace: 4096,
    });
    const pieces = events
      .filter((event): event is Extract<DrawEvent, { kind: 'fillText' }> => event.kind === 'fillText')
      .filter((event) => event.text !== 'lead');

    expect(pieces.map((event) => event.text).join('')).toBe(url);
    expect(pieces.length).toBeGreaterThan(1);
    // The leading "lead" run and this prefix share one Latin snap block.
    // Evaluating the prefix as a standalone block would admit a different
    // number of glyphs at the cell-rounding boundary.
    expect(pieces[0]?.text).toBe('https://e');
  });

  it('breaks a no-space token wider than the line at the character level', async () => {
    // pageWidth 400, margins 0, 16px/char ⇒ 25 chars per line. A 40-char URL with
    // no break opportunity must wrap (ECMA-376 prescribes no algorithm; Word
    // breaks an over-long word at the character level so it stays in the column).
    const url = 'http://example.com/aaaaaaaaaaaaaaaaaaaaa'; // 40 chars, no spaces
    expect(url.length).toBe(40);
    const events = await render([textRun(url)]);
    const texts = events.filter((e) => e.kind === 'fillText') as Array<{ text: string; x: number }>;
    // It wraps onto more than one line (the bug drew it as a single 640px line).
    expect(texts.length).toBeGreaterThan(1);
    // No drawn line exceeds the 400px content width (25 chars × 16px).
    for (const t of texts) expect(t.text.length * 16).toBeLessThanOrEqual(400 + 1e-6);
    // Every character is preserved across the wrap (character-level, lossless).
    expect(texts.map((t) => t.text).join('')).toBe(url);
  });

  it('never tears an over-long same-slot Devanagari grapheme cluster', async () => {
    const cluster = '\u0915\u093f';
    const token = cluster.repeat(30);
    const events = await render([textRun(token)]);
    const texts = events.filter((e) => e.kind === 'fillText') as Array<{ text: string; x: number }>;

    expect(texts.length).toBeGreaterThan(1);
    expect(texts.map((event) => event.text).join('')).toBe(token);
    expect(texts.every((event) => !event.text.startsWith('\u093f'))).toBe(true);
    expect(texts.every((event) => !event.text.endsWith('\u0915'))).toBe(true);
  });

  it('still wraps a normal sentence at spaces, not mid-word', async () => {
    // Guard: ordinary text must keep wrapping at spaces — the over-long path only
    // engages for a single token wider than the whole line.
    const events = await render([textRun('alpha bravo charlie delta echo foxtrot')]);
    const texts = events.filter((e) => e.kind === 'fillText') as Array<{ text: string; x: number }>;
    // Each drawn token is a whole word (plus trailing space), never a mid-word slice.
    for (const t of texts) {
      expect(t.text.trim()).toMatch(/^(alpha|bravo|charlie|delta|echo|foxtrot)$/);
    }
  });
});
