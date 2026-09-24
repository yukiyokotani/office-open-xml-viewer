import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  ParagraphBorders,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.7 — a paragraph BOTTOM border is drawn `w:space` points below
// the text ("the space after the bottom of the text … before this border is
// drawn"), and §17.3.4 gives the border its own width (`w:sz`, eighths of a point).
// drawParaBorders strokes the line CENTERED on `textBottom + space`, so its outer
// (bottom) edge is at `textBottom + space + width/2`. Word reserves that whole
// extent in the vertical flow — a bottom-bordered paragraph pushes the FOLLOWING
// paragraph BELOW the border rather than letting its first line box overlap the
// rule. The spec is silent on the flow reservation; this is Word's observed layout
// (sample-14: the reference-list rule sat ~1.75 pt too high — half a border-width
// plus its space — so "Further examples…" nearly touched the rule).
//
// The renderer used to draw the bottom border PAST `state.y` without advancing the
// flow by that extent, so the next paragraph overlapped it. This test measures the
// flow delta a bottom border introduces: the follower's baseline must drop by
// exactly `space + width/2` versus an identical layout whose leading paragraph has
// NO border.

const BORDER_COLOR = 'aa00bb';
const SPACE_PT = 1;
const WIDTH_PT = 1.5; // sz12 = 12 eighths of a point → 1.5 pt
const EXPECTED_EXTENT = SPACE_PT + WIDTH_PT / 2; // 1.75 pt

interface HStroke { y: number; strokeStyle: string; }
interface FillRectCall { x: number; y: number; w: number; h: number; fillStyle: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  hStrokes: HStroke[];
  textBaselines: number[];
  textCalls: { text: string; y: number }[];
  fillRects: FillRectCall[];
} {
  let font = '10px serif';
  let strokeStyle = '#000';
  let fillStyle = '#000';
  const hStrokes: HStroke[] = [];
  const textBaselines: number[] = [];
  const textCalls: { text: string; y: number }[] = [];
  const fillRects: FillRectCall[] = [];
  let path: { x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {},
    beginPath() { path = []; },
    closePath() {},
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() {
      for (let i = 1; i < path.length; i++) {
        if (path[i].y === path[i - 1].y) hStrokes.push({ y: path[i].y, strokeStyle });
      }
    },
    fill() {},
    fillRect(x: number, y: number, w: number, h: number) {
      fillRects.push({ x, y, w, h, fillStyle });
    },
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(t: string, _x: number, y: number) {
      textBaselines.push(y);
      textCalls.push({ text: t, y });
    },
    strokeText() {},
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  // Real CanvasRenderingContext2D has a `.canvas` back-reference; the renderer
  // reads it on some table paths (e.g. the §17.4.80 Y-axis clip).
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, hStrokes, textBaselines, textCalls, fillRects };
}

function bottomBorderOnly(): ParagraphBorders {
  return {
    top: null,
    bottom: { style: 'single', color: BORDER_COLOR, width: WIDTH_PT, space: SPACE_PT } as NonNullable<ParagraphBorders['bottom']>,
    left: null, right: null, between: null,
  };
}

function topBorderOnly(): ParagraphBorders {
  return {
    top: { style: 'single', color: BORDER_COLOR, width: WIDTH_PT, space: SPACE_PT } as NonNullable<ParagraphBorders['top']>,
    bottom: null,
    left: null, right: null, between: null,
  };
}

function para(text: string, borders: ParagraphBorders | null): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    borders,
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
          hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

const PAGE_WIDTH = 400;
function docOf(...paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_WIDTH, pageHeight: 4000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => p as unknown as BodyElement),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function followerBaseline(leadHasBorder: boolean): Promise<{ baseline: number; borderY: number | null }> {
  const { canvas, hStrokes, textBaselines } = makeRecordingCanvas();
  await renderDocumentToCanvas(
    docOf(para('', leadHasBorder ? bottomBorderOnly() : null), para('Follower', null)),
    canvas, 0, {
      dpr: 1, width: PAGE_WIDTH, // scale = 1 px per pt
      fetchImage: async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    });
  const borderStrokes = hStrokes.filter((s) => s.strokeStyle === `#${BORDER_COLOR}`);
  return {
    baseline: Math.min(...textBaselines),
    borderY: borderStrokes.length ? Math.max(...borderStrokes.map((s) => s.y)) : null,
  };
}

describe('a bottom paragraph border reserves flow so following content clears it (§17.3.1.7)', () => {
  it('reserves top border space and the outer half-stroke before its text (§17.3.1.42)', async () => {
    const topWithBorder = makeRecordingCanvas();
    await renderDocumentToCanvas(
      docOf(para('Leader', null), para('', topBorderOnly()), para('Follower', null)),
      topWithBorder.canvas, 0, { dpr: 1, width: PAGE_WIDTH },
    );
    const topWithoutBorder = makeRecordingCanvas();
    await renderDocumentToCanvas(
      docOf(para('Leader', null), para('', null), para('Follower', null)),
      topWithoutBorder.canvas, 0, { dpr: 1, width: PAGE_WIDTH },
    );
    const followerDelta = topWithBorder.textBaselines[1]! - topWithoutBorder.textBaselines[1]!;
    expect(followerDelta).toBeCloseTo(SPACE_PT + WIDTH_PT / 2, 3);
    const borderY = Math.min(...topWithBorder.hStrokes
      .filter((stroke) => stroke.strokeStyle === `#${BORDER_COLOR}`)
      .map((stroke) => stroke.y));
    expect(Number.isFinite(borderY)).toBe(true);
  });

  it('keeps the same top-border reservation in table-cell paragraphs', async () => {
    const renderFollower = async (bordered: boolean): Promise<number> => {
      const cell: DocTableCell = {
        content: [para('Leader', null), para('', bordered ? topBorderOnly() : null), para('Follower', null)],
        colSpan: 1, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        vAlign: 'top', widthPt: 300,
      } as unknown as DocTableCell;
      const table: DocTable = {
        colWidths: [300],
        rows: [{ cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow],
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
        jc: 'left',
      } as unknown as DocTable;
      const { canvas, textCalls } = makeRecordingCanvas();
      await renderDocumentToCanvas({
        ...docOf(), body: [{ type: 'table', ...table } as BodyElement],
      }, canvas, 0, { dpr: 1, width: PAGE_WIDTH });
      return textCalls.find((call) => call.text === 'Follower')!.y;
    };
    expect(await renderFollower(true) - await renderFollower(false))
      .toBeCloseTo(SPACE_PT + WIDTH_PT / 2, 3);
  });

  it('keeps the same top-border reservation in header paragraphs', async () => {
    const renderFollower = async (bordered: boolean): Promise<number> => {
      const model: DocxDocumentModel = {
        ...docOf(para('Body', null)),
        headers: {
          default: { body: [
            para('Leader', null) as BodyElement,
            para('', bordered ? topBorderOnly() : null) as BodyElement,
            para('Follower', null) as BodyElement,
          ] },
          first: null, even: null,
        },
      };
      const { canvas, textCalls } = makeRecordingCanvas();
      await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: PAGE_WIDTH });
      return textCalls.find((call) => call.text === 'Follower')!.y;
    };
    expect(await renderFollower(true) - await renderFollower(false))
      .toBeCloseTo(SPACE_PT + WIDTH_PT / 2, 3);
  });

  it('a bottom border drops the following paragraph by exactly space + width/2', async () => {
    // Baseline layouts differ ONLY in whether the leading empty paragraph carries a
    // bottom border. The border must push the follower down by its outer extent
    // (space + half the stroke width) so the follower's line box clears the rule.
    const withBorder = await followerBaseline(true);
    const noBorder = await followerBaseline(false);

    expect(withBorder.borderY).not.toBeNull();
    const delta = withBorder.baseline - noBorder.baseline;
    // Exact reservation (a sub-pixel crispness nudge on the stroke does not move the
    // baseline, which is placed by the flow cursor, so no tolerance is needed here —
    // allow a hair for float arithmetic).
    expect(delta).toBeCloseTo(EXPECTED_EXTENT, 3);

    // And the follower's line box now clears the border's OUTER edge: with a
    // 10 pt font (ascent 0.8 em = 8 pt here) the box top = baseline − 8 sits at or
    // below the border edge = borderY + width/2.
    const boxTop = withBorder.baseline - 8;
    expect(boxTop).toBeGreaterThanOrEqual((withBorder.borderY as number) + WIDTH_PT / 2 - 1e-6);
  });

  it('a bordered paragraph inside a table cell measures as tall as it paints (B2 single measurer)', async () => {
    // Retained table acquisition must include the paragraph's trailing advance
    // max(spaceAfter, bottom-border extent). Omitting it sizes the row
    // SHORT and the painted content pokes past the cell band (clipped under an
    // `exact` row rule, bleeding otherwise). A LARGE extent (space=6, sz48 → 6 pt
    // stroke → extent 9 pt) makes the pre-fix shortfall far exceed the line box's
    // internal leading, so the assertion discriminates cleanly.
    const CELL_SPACE_PT = 6;
    const CELL_WIDTH_PT = 6; // sz48 = 48 eighths of a point → 6 pt stroke
    const CELL_BG = 'ffeecc';
    const ruled: DocParagraph = {
      ...para('Ruled', null),
      borders: {
        top: null,
        bottom: { style: 'single', color: BORDER_COLOR, width: CELL_WIDTH_PT, space: CELL_SPACE_PT } as NonNullable<ParagraphBorders['bottom']>,
        left: null, right: null, between: null,
      },
    } as DocParagraph;
    const cell: DocTableCell = {
      content: [
        { type: 'paragraph', ...ruled },
        { type: 'paragraph', ...para('Below', null) },
      ],
      colSpan: 1, vMerge: null,
      borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
      background: CELL_BG,
      vAlign: 'top',
      widthPt: 300,
    } as unknown as DocTableCell;
    const row: DocTableRow = {
      cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    } as unknown as DocTableRow;
    const table: DocTable = {
      colWidths: [300], rows: [row],
      borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      jc: 'left',
    } as unknown as DocTable;
    const doc = {
      section: {
        pageWidth: PAGE_WIDTH, pageHeight: 4000,
        marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
        headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      } as SectionProps,
      body: [{ type: 'table', ...table } as unknown as BodyElement],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { 'Times New Roman': 'roman' },
    } as unknown as DocxDocumentModel;

    const { canvas, hStrokes, textBaselines, fillRects } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, {
      dpr: 1, width: PAGE_WIDTH, // scale = 1 px per pt
      fetchImage: async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    });

    // The cell background fillRect's height IS the measured row band (auto row =
    // measureCellContentHeightPx). The painted content must fit inside it.
    const bg = fillRects.filter((r) => r.fillStyle === `#${CELL_BG}`);
    expect(bg.length).toBe(1);
    const bandBottom = bg[0].y + bg[0].h;

    // Both paragraphs drew; "Below" is the lowest baseline. Its line box bottom
    // (baseline + 0.2 em descent with the mock metrics) must not poke past the
    // measured band — pre-fix the band was `space + width/2` (9 pt) short.
    expect(textBaselines.length).toBeGreaterThanOrEqual(2);
    const lastBaseline = Math.max(...textBaselines);
    expect(lastBaseline + 2).toBeLessThanOrEqual(bandBottom + 1e-6);

    // Sanity: the rule itself was drawn, inside the band.
    const rule = hStrokes.filter((s) => s.strokeStyle === `#${BORDER_COLOR}`);
    expect(rule.length).toBeGreaterThanOrEqual(1);
    expect(Math.max(...rule.map((s) => s.y))).toBeLessThanOrEqual(bandBottom + CELL_WIDTH_PT / 2 + 1e-6);
  });
});
