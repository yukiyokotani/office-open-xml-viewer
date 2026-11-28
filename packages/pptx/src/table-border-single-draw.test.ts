import { describe, it, expect } from 'vitest';
import type { Stroke } from '@silurus/ooxml-core';
import { renderTable, renderTableWithLayoutMemo, renderTextBody, type TableLayoutMemo } from './renderer.js';
import { presentFacesContext } from './test-support/present-faces.js';
import type { TableCell, TableElement, TableRow, TextBody } from './types';

// DrawingML `<a:tbl>` — end-to-end render: an interior gridline SHARED by two
// adjacent cells is drawn exactly ONCE (not once per cell), and when the two
// cells DISAGREE the drawn line is the conflict WINNER, not the later-painted
// cell. The spec is SILENT for PresentationML (no §17.4.66 analog), so this is
// the pptx leg / structural mirror of docx PR #811.
//
// We record each stroked segment (endpoints + width + colour) from a fake 2D
// context (the pptx renderer strokes borders via beginPath/moveTo/lineTo/stroke
// after applyStroke sets strokeStyle + lineWidth), then read the shared line.

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string; }
interface FillRect { x: number; y: number; width: number; height: number; color: string; }

function makeRecordingCtx(onMeasure?: () => void): { ctx: CanvasRenderingContext2D; strokes: StrokeSeg[]; fills: FillRect[]; texts: Array<{ text: string; y: number; font: string }> } {
  const strokes: StrokeSeg[] = [];
  const fills: FillRect[] = [];
  const texts: Array<{ text: string; y: number; font: string }> = [];
  let cur = { x: 0, y: 0 };
  let pending: { x1: number; y1: number; x2: number; y2: number } | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  const ctx = {
    canvas: { width: 1000, height: 1000 },
    get lineWidth() { return lineWidth; },
    set lineWidth(v: number) { lineWidth = v; },
    get strokeStyle() { return strokeStyle as string; },
    set strokeStyle(v: string) { strokeStyle = v; },
    fillStyle: '#000000',
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
    globalAlpha: 1,
    save() {}, restore() {},
    beginPath() { pending = null; }, closePath() {},
    moveTo(x: number, y: number) { cur = { x, y }; },
    lineTo(x: number, y: number) { pending = { x1: cur.x, y1: cur.y, x2: x, y2: y }; cur = { x, y }; },
    stroke() {
      if (pending) strokes.push({ ...pending, width: lineWidth, color: String(strokeStyle) });
    },
    fill() {},
    fillRect(x: number, y: number, width: number, height: number) {
      fills.push({ x, y, width, height, color: String(this.fillStyle) });
    },
    strokeRect() {}, clearRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {}, transform() {}, resetTransform() {},
    setLineDash() {}, getLineDash() { return []; },
    drawImage() {}, arc() {}, arcTo() {}, ellipse() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    createPattern() { return null; },
    // Deterministic advances: every character is half the font's px size
    // wide, so wrapping happens the same way on every host.
    measureText(text: string) {
      onMeasure?.();
      const sizePx = Number(/(\d+(?:\.\d+)?)px/.exec(String(this.font))?.[1] ?? 10);
      return { width: text.length * sizePx * 0.5, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
    },
    fillText(text: string, _x: number, y: number) { texts.push({ text, y, font: String(this.font) }); }, strokeText() {},
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    letterSpacing: '0px',
    globalCompositeOperation: 'source-over' as GlobalCompositeOperation,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, strokes, fills, texts };
}

const EMU = 12700; // 1 pt

// A stroke helper. width defaults to 1pt; colour is 6-hex without '#'.
const ln = (over: Partial<Stroke> = {}): Stroke => ({ color: '000000', width: EMU, ...over });

function cell(over: Partial<TableCell> = {}): TableCell {
  return {
    textBody: null, fill: null,
    borderL: null, borderR: null, borderT: null, borderB: null,
    diagonalTl: null, diagonalTr: null,
    gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
    ...over,
  } as TableCell;
}

function tableOf(rows: TableCell[][], cols: number[]): TableElement {
  const rowH = 20 * EMU;
  const tableRows: TableRow[] = rows.map((cells) => ({ height: rowH, cells }));
  return {
    type: 'table', x: 0, y: 0,
    width: cols.reduce((a, b) => a + b, 0), height: rowH * rows.length,
    rotation: 0, flipH: false, flipV: false,
    cols, rows: tableRows,
  };
}

// scale=1 ⇒ emuToPx is identity; a 60pt (60*EMU) column ⇒ 60px wide. Two 60pt
// columns put the shared vertical gridline at x=60. Rows are 20pt ⇒ shared
// horizontal line at y=20.
const SCALE = 1 / EMU; // makes 1 EMU → (1/EMU)px; so 60*EMU → 60px, widths in EMU → px
const COL = 60 * EMU;

function render(t: TableElement): StrokeSeg[] {
  const { ctx, strokes } = makeRecordingCtx();
  renderTable(ctx, t, SCALE, undefined, { themeMajorFont: null, themeMinorFont: null, dpr: 1 });
  return strokes;
}

function renderRecording(t: TableElement): { strokes: StrokeSeg[]; fills: FillRect[] } {
  const recording = makeRecordingCtx();
  renderTable(
    recording.ctx,
    t,
    SCALE,
    undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  );
  return recording;
}

function verticalAt(strokes: StrokeSeg[], x: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.x1 - s.x2) < 0.5 && Math.abs(s.x1 - x) <= 1);
}
function horizontalAt(strokes: StrokeSeg[], y: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.y1 - y) <= 1);
}

// applyStroke sets ctx.strokeStyle via hexToRgba, so a recorded segment colour is
// `rgba(r,g,b,a)`. Convert a 6-hex authoring colour to that form for comparison.
function rgba(hex: string): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return `rgba(${r},${g},${b},1)`;
}

describe('PowerPoint table-cell paragraph boundaries', () => {
  it('ignores first and last paragraph edge gaps even with spcFirstLastPara=1', () => {
    // PowerPoint PDF controls: a centred 24pt Arial cell with 12pt spcAft or
    // 50% spcAft keeps the same glyph position for absent/0/1. Two paragraphs
    // in the same cell retain the authored gap between them.
    const tableY = (
      spaceAfter: number | null,
      second: boolean,
      spaceBefore: number | null = null,
      spaceAfterPct: number | null = null,
    ): number[] => {
      const para = (text: string, after: number | null) => ({
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: text === 'first' ? spaceBefore : null,
        spaceAfter: after, spaceAfterPct: text === 'first' ? spaceAfterPct : null,
        spaceLine: null,
        runs: [{ type: 'text', text, fontSize: 24, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      });
      const textBody = {
        verticalAnchor: 'ctr', spcFirstLastPara: true,
        paragraphs: [para('first', spaceAfter), ...(second ? [para('second', null)] : [])],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0,
        wrap: 'square', vert: 'horz', autoFit: 'none',
      } as unknown as TextBody;
      const t = tableOf([[cell({ textBody })]], [COL]);
      t.rows[0].height = 100 * EMU;
      const { ctx, texts } = makeRecordingCtx();
      renderTable(ctx, t, SCALE);
      return texts.map(({ y }) => y);
    };

    expect(tableY(1200, false)[0]).toBeCloseTo(tableY(null, false)[0], 5);
    expect(tableY(null, false, 1200)[0]).toBeCloseTo(tableY(null, false)[0], 5);
    expect(tableY(null, false, null, 50000)[0]).toBeCloseTo(tableY(null, false)[0], 5);
    const plain = tableY(null, true);
    const spaced = tableY(1200, true);
    expect(spaced[1] - spaced[0] - (plain[1] - plain[0])).toBeCloseTo(12, 5);
  });

  it('does not grow an automatic row from the last paragraph gap when the flag is true', () => {
    // PowerPoint PDF control with a:tr@h=0: the following row starts at the
    // same position for 0pt and 12pt spcAft, including flag=1.
    const boundaryY = (spaceAfter: number | null): number => {
      const textBody = {
        verticalAnchor: 't', spcFirstLastPara: true,
        paragraphs: [{
          alignment: 'l', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter, spaceLine: null,
          runs: [{ type: 'text', text: 'first', fontSize: 24, fontFamily: 'Arial' }],
          bullet: { type: 'none' }, eaLnBrk: true,
        }],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0,
        wrap: 'square', vert: 'horz', autoFit: 'none',
      } as unknown as TextBody;
      const t = tableOf([
        [cell({ textBody, borderB: ln() })],
        [cell({ borderT: ln() })],
      ], [COL]);
      t.rows[0].height = 0;
      const border = render(t).find((s) => s.y1 === s.y2 && s.y1 > 0);
      expect(border).toBeDefined();
      return (border as StrokeSeg).y1;
    };

    expect(boundaryY(1200)).toBeCloseTo(boundaryY(null), 5);
  });
});

describe('DrawingML <a:tbl> — shared interior gridline drawn once (spec-silent)', () => {
  it('grows a positive header row for explicit percentage line spacing and cell margins', () => {
    // PowerPoint PDF of a table with 16pt text, 36.85pt header minimum,
    // 8.5pt top/bottom margins and 120% lnSpc grows the header. The same
    // table at 100% lnSpc, or with zero vertical margins, keeps the minimum.
    const boundary = (pct: number, marginPt: number): number => {
      const textBody = {
        verticalAnchor: 'b', paragraphs: [{
          alignment: 'ctr', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter: 300,
          spaceLine: { type: 'pct', val: pct },
          runs: [{ type: 'text', text: 'Header', fontSize: 16, fontFamily: 'Yu Gothic' }],
          bullet: { type: 'none' }, eaLnBrk: true,
        }],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 108000, rIns: 108000,
        tIns: marginPt * EMU, bIns: marginPt * EMU,
        wrap: 'square', vert: 'horz', autoFit: 'none',
      } as unknown as TextBody;
      const t = tableOf([
        [cell({ textBody, borderB: ln() })],
        [cell({ borderT: ln() })],
      ], [200 * EMU]);
      t.rows[0].height = 36.85 * EMU;
      t.height = 60 * EMU; // authored frame exceeds the row minima
      return render(t).find((s) => s.y1 === s.y2 && s.y1 > 0)?.y1 ?? -1;
    };

    expect(boundary(120000, 8.5) - boundary(100000, 8.5)).toBeCloseTo(4, 2);
    expect(boundary(120000, 0)).toBeCloseTo(boundary(100000, 8.5), 2);
  });

  it('composes a table background below translucent cell bands', () => {
    const t = tableOf([[cell({ fill: { fillType: 'solid', color: '4472C466' } })]], [COL]);
    t.background = { fillType: 'solid', color: '22BBCC' };
    const { fills } = renderRecording(t);
    expect(fills.map((f) => f.color)).toEqual([rgba('22BBCC'), 'rgba(68,114,196,0.4)']);
    expect(fills[0].height).toBe(20);
  });

  it('does not grow an authored row from substituted-font design metrics', () => {
    // ECMA-376 §21.1.2.2.5/.11: 120% line spacing is based on the largest
    // authored natural line box. The terminal 3pt spcAft is suppressed:
    // 16pt * 120% natural * 120% authored + 2 * 8.5pt insets = 40.04pt,
    // so the authored 40.6pt row already fits. The old measure-only path used
    // Meiryo's ~1.596em design line box and moved this boundary down to 42.5pt.
    const textBody = {
      verticalAnchor: 'ctr',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: 300,
        spaceLine: { type: 'pct', val: 120000 },
        runs: [{ type: 'text', text: '流動負債', fontSize: 16, fontFamily: 'Meiryo' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 108000, rIns: 108000, tIns: 108000, bIns: 108000,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const rowH = 40.6 * EMU;
    const t = tableOf([
      [cell({ textBody, borderB: ln() })],
      [cell({ borderT: ln() })],
    ], [COL]);
    t.rows[0].height = rowH;

    expect(horizontalAt(render(t), 40.6)).toHaveLength(1);
  });

  it('uses the natural 120% line box when table line spacing is omitted', () => {
    // ECMA-376 leaves omitted line spacing to the text font. Current
    // PowerPoint grows a zero-height table row from the natural single-line
    // box: 16pt * 120% + 1pt top + 1pt bottom = 21.2pt. An authored
    // `<a:spcPct val="120000"/>` scales that natural line box again.
    const textBody = {
      verticalAnchor: 'ctr',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'auto', fontSize: 16, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: EMU, bIns: EMU,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([[cell({ textBody, borderB: ln() })]], [COL]);
    t.rows[0].height = 0;

    expect(horizontalAt(render(t), 21.2)).toHaveLength(1);
  });

  it('grows a positive row to the 120% line box when line spacing is omitted', () => {
    // PowerPoint PDF controls (Windows and Mac) grow a 20pt minimum holding
    // one 20pt line to 24pt: an omitted lnSpc line takes 1.2 x size in a
    // positive row too. Here 16pt * 1.2 + 1pt top + 1pt bottom = 21.2pt
    // exceeds the 18pt minimum.
    const textBody = {
      verticalAnchor: 'ctr',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'fixed', fontSize: 16, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: EMU, bIns: EMU,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([[cell({ textBody, borderB: ln() })]], [COL]);
    t.rows[0].height = 18 * EMU;

    const lines = render(t);
    expect(horizontalAt(lines, 18)).toHaveLength(0);
    expect(horizontalAt(lines, 21.2)).toHaveLength(1);
  });

  it('grows a positive row for every painted line box', () => {
    // [MS-OE376] §2.1.1347 gives Office's minimum-row rule. Every line,
    // the last included, consumes the line box used by paint. Three 18pt lines at explicit 100% use 3 * 21.6pt plus
    // 3.6pt top/bottom insets = 72pt.
    const textBody = {
      verticalAnchor: 'ctr',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null,
        spaceLine: { type: 'pct', val: 100000 },
        runs: [
          { type: 'text', text: 'one', fontSize: 18, fontFamily: 'Arial' },
          { type: 'break' },
          { type: 'text', text: 'two', fontSize: 18, fontFamily: 'Arial' },
          { type: 'break' },
          { type: 'text', text: 'three', fontSize: 18, fontFamily: 'Arial' },
        ],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 3.6 * EMU, bIns: 3.6 * EMU,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([[cell({ textBody, borderB: ln() })]], [COL]);
    t.rows[0].height = 60 * EMU;
    t.height = 72 * EMU;

    expect(horizontalAt(render(t), 72)).toHaveLength(1);
  });

  it('grows a positive row even when the row minima already fill the frame', () => {
    // #2345: generators such as python-pptx write equal a:tr@h minima whose
    // sum equals the graphic-frame extent. PowerPoint still treats each
    // a:tr@h as a minimum ([MS-OE376] §2.1.1347) and grows the row to its
    // wrapped text, pushing the next row down; the frame extent is not a
    // reason to let the text overflow into the following row.
    // One unbroken run that must wrap: at 9px per 18px character,
    // "aaaa bbbb cccc" breaks into three lines in the 60px column, so
    // 3 * 21.6pt + 3.6pt top and bottom = 72pt.
    const textBody = {
      verticalAnchor: 'ctr',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null,
        spaceLine: { type: 'pct', val: 100000 },
        runs: [{ type: 'text', text: 'aaaa bbbb cccc', fontSize: 18, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 3.6 * EMU, bIns: 3.6 * EMU,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([
      [cell({ textBody, borderB: ln() })],
      [cell({ borderT: ln(), borderB: ln() })],
    ], [COL]);
    t.rows[0].height = 60 * EMU;
    t.rows[1].height = 40 * EMU;
    t.height = 100 * EMU;

    const { ctx, strokes, texts } = makeRecordingCtx();
    renderTable(ctx, t, SCALE, undefined, { themeMajorFont: null, themeMinorFont: null, dpr: 1 });
    expect(texts.map((entry) => entry.text.trim())).toEqual(['aaaa', 'bbbb', 'cccc']);
    expect(horizontalAt(strokes, 60)).toHaveLength(0);
    expect(horizontalAt(strokes, 72)).toHaveLength(1);
    // The next row keeps its 40pt minimum below the grown row.
    expect(horizontalAt(strokes, 100)).toHaveLength(0);
    expect(horizontalAt(strokes, 112)).toHaveLength(1);
  });

  it('keeps a positive row for a warped body, whose paint does not wrap', () => {
    // a:prstTxWarp textInflate (ECMA-376 §20.1.9.19): paint fits the text
    // into the warp envelope without wrapping, so the flat wrapped measure
    // is not its extent. With the minima filling the frame the 24pt row
    // stays, as before this growth rule.
    const textBody = {
      verticalAnchor: 't',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'aaaa bbbb cccc dddd eeee', fontSize: 24, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'horz', autoFit: 'none',
      textWarp: { preset: 'textInflate' },
    } as unknown as TextBody;
    const t = tableOf([
      [cell({ textBody, borderB: ln() })],
      [cell({ borderT: ln() })],
    ], [COL]);
    t.rows[0].height = 24 * EMU;
    t.height = 44 * EMU;
    const strokes = render(t);
    expect(horizontalAt(strokes, 24)).toHaveLength(1);
    expect(horizontalAt(strokes, 144)).toHaveLength(0);
    // textNoShape is no warp: it grows like ordinary text.
    textBody.textWarp = { preset: 'textNoShape' };
    expect(horizontalAt(render(t), 24)).toHaveLength(0);
  });

  it('paints a fontAlgn paragraph by the model its body falls back to', () => {
    // A 20pt Arial paragraph with fontAlgn="t" resolves in the metric pass,
    // but the next paragraph's face has no reference metrics, so the whole
    // body is laid out again with the ordinary model (baseline 0.8 x 24pt
    // per 24pt line). The discarded pass's alignment offset must not move
    // the first baseline (it would be 37.2pt).
    const para = (text: string, fontFamily: string, fontAlgn?: string) => ({
      alignment: 'l', marL: 0, marR: 0, indent: 0,
      spaceBefore: null, spaceAfter: null, spaceLine: null,
      ...(fontAlgn ? { fontAlgn } : {}),
      runs: [{ type: 'text', text, fontSize: 20, fontFamily }],
      bullet: { type: 'none' }, eaLnBrk: true,
    });
    const textBody = {
      verticalAnchor: 't',
      paragraphs: [para('top', 'Arial', 't'), para('next', 'No Reference Metrics Face')],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([[cell({ textBody })]], [COL]);
    t.rows[0].height = 100 * EMU;
    t.height = 100 * EMU;
    const { ctx, texts } = makeRecordingCtx();
    // Arial present (its reference metrics apply); the other face is not.
    renderTable(ctx, t, SCALE, undefined, presentFacesContext({
      ...textBody, paragraphs: textBody.paragraphs.slice(0, 1),
    } as TextBody));
    expect(texts.map((entry) => entry.text)).toEqual(['top', 'next']);
    expect(texts[0].y).toBeCloseTo(19.2, 5);
    expect(texts[1].y).toBeCloseTo(43.2, 5);
  });

  it('keeps only row-sizing layouts, and only until their cell is painted', () => {
    // Row sizing records each measured paragraph layout for the paint of
    // that cell (one per cell here); paint drops it. A shrink-to-fit cell
    // (ODF style:shrink-to-fit) is skipped by row sizing, and the font
    // scales its paint tries are never retained; a warped cell, whose paint
    // does not reuse its flat layout, is not recorded either.
    const shrinking = {
      verticalAnchor: 't', shrinkOnOverflow: true,
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'aaaa bbbb cccc dddd eeee ffff gggg', fontSize: 24, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'horz', autoFit: 'norm',
    } as unknown as TextBody;
    const warped = { ...lines(2), textWarp: { preset: 'textInflate' } } as TextBody;
    const t = tableOf([
      [cell({ textBody: lines(1) }), cell({ textBody: lines(1) }), cell({ textBody: shrinking })],
      [cell({ textBody: lines(1) }), cell({ textBody: warped }), cell()],
    ], [COL, COL, COL]);
    t.height = 60 * EMU; // frame slack: the warped cell is measured too
    const memo: TableLayoutMemo = { layouts: new Map(), lines: 0 };
    const sizes: number[] = [];
    const { ctx, texts } = makeRecordingCtx(() => sizes.push(memo.lines));
    renderTableWithLayoutMemo(memo, ctx, t, SCALE, undefined);
    expect(texts.some((entry) => entry.text.startsWith('aaaa'))).toBe(true);
    // Three measured one-line cells, never more; each painted cell gives its
    // layout up, so nothing is held once the table is painted.
    expect(Math.max(...sizes)).toBe(3);
    expect(memo.lines).toBe(0);
    expect(memo.layouts.size).toBe(0);
  });

  it('bounds the layouts a large table retains', () => {
    // 30 x 30 cells of four lines each: 3600 measured lines. Recording stops
    // at the retention bound, so the first cells reuse their layouts and the
    // rest are laid out again, painting the same rows.
    const rows = Array.from({ length: 30 }, () =>
      Array.from({ length: 30 }, () => cell({ textBody: lines(4) })));
    const t = tableOf(rows, Array(30).fill(COL));
    const memo: TableLayoutMemo = { layouts: new Map(), lines: 0 };
    let peak = 0;
    const { ctx, texts } = makeRecordingCtx(() => { peak = Math.max(peak, memo.lines); });
    renderTableWithLayoutMemo(memo, ctx, t, SCALE, undefined);
    expect(peak).toBeGreaterThan(1000);
    expect(peak).toBeLessThanOrEqual(2048);
    expect(memo.lines).toBe(0);
    const reference = makeRecordingCtx();
    renderTable(reference.ctx, t, SCALE);
    expect(texts).toEqual(reference.texts);
    expect(texts.length).toBe(3600);
  });

  it('lets no other text body see the measured layouts of a table cell', () => {
    // An onTextRun consumer renders the same paragraph on a second canvas
    // with another theme font while the cell paints. It must get its own
    // layout (Courier New), not the table's measured one (Arial).
    const textBody = {
      verticalAnchor: 't',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'theme', fontSize: 18, fontFamily: '+mn-lt' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([[cell({ textBody })]], [COL]);
    const outer = makeRecordingCtx();
    const nested = makeRecordingCtx();
    let rendered = false;
    renderTable(outer.ctx, t, SCALE, undefined, { themeMajorFont: null, themeMinorFont: 'Arial', dpr: 1 }, () => {
      if (rendered) return;
      rendered = true;
      renderTextBody(nested.ctx, textBody, 0, 0, COL * SCALE, 20, SCALE, null, 0, false, false,
        '#000000', undefined, { themeMajorFont: null, themeMinorFont: 'Courier New', dpr: 1 });
    });
    expect(rendered).toBe(true);
    expect(outer.texts[0].font).toContain('Arial');
    expect(nested.texts.map((entry) => entry.text)).toEqual(['theme']);
    expect(nested.texts[0].font).toContain('Courier New');
    expect(nested.texts[0].font).not.toContain('Arial');
  });

  it('keeps a positive row for rotated cell text, whose measure is the box width', () => {
    // measureOnly reports a rotated or stacked body by its cell width, not by
    // its laid-out extent. Growing a positive row to that width would invent
    // height, so a positive row keeps its minimum for these bodies.
    const textBody = {
      verticalAnchor: 't',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: [{ type: 'text', text: 'v', fontSize: 10, fontFamily: 'Arial' }],
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'vert', autoFit: 'none',
    } as unknown as TextBody;
    const t = tableOf([
      [cell({ textBody, borderB: ln() })],
      [cell({ borderT: ln() })],
    ], [COL]);

    expect(horizontalAt(render(t), 20)).toHaveLength(1);
  });

  // n 18pt lines at explicit 100%: n * 21.6pt plus 3.6pt top and bottom.
  const lines = (n: number, vert = 'horz') => ({
    verticalAnchor: 't',
    paragraphs: [{
      alignment: 'l', marL: 0, marR: 0, indent: 0,
      spaceBefore: null, spaceAfter: null,
      spaceLine: { type: 'pct', val: 100000 },
      runs: Array.from({ length: n }, (_, i) => [
        ...(i > 0 ? [{ type: 'break' }] : []),
        { type: 'text', text: `l${i}`, fontSize: 18, fontFamily: 'Arial' },
      ]).flat(),
      bullet: { type: 'none' }, eaLnBrk: true,
    }],
    defaultFontSize: null, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 3.6 * EMU, bIns: 3.6 * EMU,
    wrap: 'square', vert, autoFit: 'none',
  } as unknown as TextBody);
  const threeLines = (vert = 'horz') => lines(3, vert);

  it('keeps every non-horizontal vert token on the earlier rule', () => {
    // ST_TextVerticalType (ECMA-376 §20.1.10.83): mongolianVert is vertical
    // text too, and an xsd:token value padded with collapsible whitespace is
    // the same token. Only an effectively horizontal body (omitted or `horz`,
    // padded or not) takes the growth rule; with the minima filling the
    // frame every other value keeps the 20pt row.
    for (const vert of ['mongolianVert', ' vert ', '\teaVert\n', ' wordArtVert', 'vert270 ', 'unknown']) {
      const t = tableOf([
        [cell({ textBody: threeLines(vert), borderB: ln() })],
        [cell({ borderT: ln() })],
      ], [COL]);
      expect(horizontalAt(render(t), 20), vert).toHaveLength(1);
    }
    for (const vert of [' horz', 'horz\t']) {
      const t = tableOf([
        [cell({ textBody: threeLines(vert), borderB: ln() })],
        [cell({ borderT: ln() })],
      ], [COL]);
      expect(horizontalAt(render(t), 72), vert).toHaveLength(1);
    }
  });

  it('adds a rowSpan deficit to the last spanned row', () => {
    // PowerPoint PDF controls: a merged cell taller than its rows keeps every
    // earlier row at its own height and grows only the last. Two 20pt minima
    // (frame = their sum) under a 72pt merged cell: the neighbour's shared
    // edge stays at 20pt and the merged area ends at 72pt.
    const t = tableOf([
      [cell({ textBody: threeLines(), rowSpan: 2, borderB: ln() }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell({ borderT: ln() })],
      [cell({ borderT: ln() }), cell({ borderT: ln() })],
    ], [COL, COL]);
    t.rows[2].height = 0;
    t.height = 40 * EMU;
    const strokes = render(t);
    expect(horizontalAt(strokes, 20)).toHaveLength(1);
    expect(horizontalAt(strokes, 36)).toHaveLength(0);
    // One segment per column.
    expect(horizontalAt(strokes, 72)).toHaveLength(2);
  });

  it('grows rotated cell text only under frame slack, as before', () => {
    // The box-width measure keeps the earlier rule: with the frame extent
    // above the row sum the row grows to that measure (the cell width here).
    const textBody = { ...threeLines('vert') } as TextBody;
    const t = tableOf([
      [cell({ textBody, borderB: ln() })],
      [cell({ borderT: ln() })],
    ], [COL]);
    t.height = 50 * EMU;
    expect(horizontalAt(render(t), 20)).toHaveLength(0);
  });

  it('keeps the earlier rule for an ODF table', () => {
    // ODF row heights (OpenDocument 1.3 Part 3 §20.348) are not covered by
    // the PowerPoint controls: with the minima filling the frame the row
    // keeps its height, and with frame slack it grows as before.
    const t = tableOf([
      [cell({ textBody: threeLines(), borderB: ln() })],
      [cell({ borderT: ln() })],
    ], [COL]);
    t.odfRowHeights = true;
    expect(horizontalAt(render(t), 20)).toHaveLength(1);
    t.height = 100 * EMU;
    const grown = render(t);
    expect(horizontalAt(grown, 20)).toHaveLength(0);
    expect(horizontalAt(grown, 72)).toHaveLength(1);
  });

  it('resolves nested merged cells by the row they end on, whatever the column order', () => {
    // A five-line cell over rows 1-3 (115.2pt) and a four-line cell over
    // rows 1-2 (93.6pt), three 20pt minima. The inner merge grows row 2 to
    // 73.6pt; the outer one then needs only 1.6pt more in row 3, so the table
    // ends at 115.2pt in either column order.
    const outer = () => cell({ textBody: lines(5), rowSpan: 3 });
    const inner = () => cell({ textBody: lines(4), rowSpan: 2 });
    const edge = () => cell({ borderB: ln() });
    for (const order of [[outer, inner, edge], [inner, outer, edge]]) {
      const merged = (make: () => TableCell, row: number) =>
        make === edge ? edge() : make === inner && row === 2 ? cell() : cell({ vMerge: true });
      const t = tableOf([
        order.map((make) => make()),
        order.map((make) => merged(make, 1)),
        order.map((make) => merged(make, 2)),
      ], [COL, COL, COL]);
      const strokes = render(t);
      expect(horizontalAt(strokes, 115.2)).toHaveLength(1);
      expect(horizontalAt(strokes, 168.8)).toHaveLength(0);
    }
  });

  it('resolves staggered merges by end row, not by span length (control4 J/K)', () => {
    // PowerPoint for Mac control4 (2026-11-29), four rows of 24pt minima,
    // 24pt leading (20pt text), zero margins, frame equal to the row sum:
    // eight lines over rows 1-3 beside five lines over rows 3-4 exports as
    // 24/24/144/24 in both column orders (J, and K with the columns
    // swapped). Shortest span first would give 24/24/144/96.
    const body = (n: number): TextBody => ({
      verticalAnchor: 't',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        runs: Array.from({ length: n }, (_, i) => [
          ...(i > 0 ? [{ type: 'break' }] : []),
          { type: 'text', text: `l${i}`, fontSize: 20, fontFamily: 'Arial' },
        ]).flat(),
        bullet: { type: 'none' }, eaLnBrk: true,
      }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'square', vert: 'horz', autoFit: 'none',
    } as unknown as TextBody);
    // Column a: eight lines over rows 1-3. Column b: five lines over rows 3-4.
    const a = [cell({ textBody: body(8), rowSpan: 3 }), cell({ vMerge: true }), cell({ vMerge: true }), cell()];
    const b = [cell(), cell(), cell({ textBody: body(5), rowSpan: 2 }), cell({ vMerge: true })];
    for (const swapped of [false, true]) {
      const t = tableOf([0, 1, 2, 3].map((r) => swapped
        ? [b[r], a[r], cell({ borderB: ln() })]
        : [a[r], b[r], cell({ borderB: ln() })]), [COL, COL, COL]);
      for (const row of t.rows) row.height = 24 * EMU;
      t.height = 96 * EMU;
      const edges = render(t)
        .filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && s.x1 >= 120)
        .map((s) => Math.floor(s.y1)); // strokes sit on half-pixel centres
      expect(edges, swapped ? 'K' : 'J').toEqual([24, 48, 192, 216]);
    }
  });

  it('orders a group with an earlier-rule merge the same way whatever the column order', () => {
    // A one-line merge over rows 1-4 (fits, adds nothing) and a rotated
    // merge over rows 1-2 share row 1; a four-line merge over rows 2-3
    // links them into one group. A group with an earlier-rule cell keeps
    // reading order for all of its merges, so swapping the fitting and the
    // rotated merge changes nothing: the rotated cell spreads its 12pt
    // deficit (60pt box width over 48pt) evenly, then the four-line merge
    // adds its 39.6pt deficit (93.6pt over 30 + 24pt) to row 3.
    const fitting = [cell({ textBody: lines(1), rowSpan: 4 }), cell({ vMerge: true }), cell({ vMerge: true }), cell({ vMerge: true })];
    const rotated = [cell({ textBody: threeLines('vert'), rowSpan: 2 }), cell({ vMerge: true }), cell(), cell()];
    const tall = [cell(), cell({ textBody: lines(4), rowSpan: 2 }), cell({ vMerge: true }), cell()];
    const result = (swapped: boolean) => {
      const t = tableOf([0, 1, 2, 3].map((r) => [
        ...(swapped ? [rotated[r], fitting[r]] : [fitting[r], rotated[r]]),
        tall[r], cell({ borderB: ln() }),
      ]), [COL, COL, COL, COL]);
      for (const row of t.rows) row.height = 24 * EMU;
      t.height = 100 * EMU; // frame slack: the rotated cell is measured
      return render(t)
        .filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && s.x1 >= 180)
        .map((s) => Math.floor(s.y1)); // strokes sit on half-pixel centres
    };
    // 30 / 30 / 63.6 / 24pt rows: edges 30, 60, 123.6, 147.6.
    expect(result(false)).toEqual([30, 60, 123, 147]);
    expect(result(true)).toEqual(result(false));
  });

  it('keeps a multi-column cell on the earlier rule, painting into its columns', () => {
    // numCol=2 flows four 21.6pt lines into two columns of a 54pt row
    // (two lines and the 7.2pt insets fit one column); a stacked measure
    // would ask for 93.6pt.
    const textBody = { ...lines(4), numCol: 2 } as TextBody;
    const t = tableOf([[cell({ textBody, borderB: ln() })], [cell()]], [2 * COL]);
    t.rows[0].height = 54 * EMU;
    t.height = 74 * EMU;
    const { ctx, strokes, texts } = makeRecordingCtx();
    renderTable(ctx, t, SCALE, undefined, { themeMajorFont: null, themeMinorFont: null, dpr: 1 });
    expect(horizontalAt(strokes, 54)).toHaveLength(1);
    const y = (text: string) => texts.find((entry) => entry.text === text)?.y;
    expect(y('l2')).toBe(y('l0'));
    expect(y('l3')).toBe(y('l1'));
  });

  it('keeps row-major order for interacting merges on the earlier rule', () => {
    // An ODF table with frame slack: a six-line merge over rows 1-3, then a
    // five-line merge over rows 1-2 beside it. The earlier rule resolves
    // them in reading order, as before.
    const t = tableOf([
      [cell({ textBody: lines(6), rowSpan: 3 }), cell({ textBody: lines(5), rowSpan: 2 }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell({ vMerge: true }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell(), cell({ borderB: ln() })],
    ], [COL, COL, COL]);
    t.height = 70 * EMU;
    t.odfRowHeights = true;
    const strokes = render(t);
    // Row edges in the right-hand column (main and this renderer agree).
    expect(strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && s.x1 >= 120).map((s) => s.y1))
      .toEqual([57.5, 115.5, 160.5]);
  });

  it('spreads a rotated merged cell evenly, as before, over a zero-height row', () => {
    // The box-width measure (60pt, the cell width) keeps the earlier even
    // split when a spanned row has h=0: rows of 20pt and 0pt each take half
    // of the 40pt deficit, so the boundary moves to 40pt.
    const t = tableOf([
      [cell({ textBody: threeLines('vert'), rowSpan: 2 }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell()],
    ], [COL, COL]);
    t.rows[1].height = 0;
    t.height = 20 * EMU;
    expect(horizontalAt(render(t), 40)).toHaveLength(1);
  });

  it('shared VERTICAL gridline is drawn exactly ONCE (not once per cell)', () => {
    // Left cell right = 1pt; right cell left = 1pt (same). Previously TWO strokes
    // at x=60 (one per cell); now exactly one.
    const strokes = render(tableOf([
      [cell({ borderR: ln() }), cell({ borderL: ln() })],
    ], [COL, COL]));
    expect(verticalAt(strokes, 60)).toHaveLength(1);
  });

  it('wider line wins the shared gridline (rule #1), ignoring paint order', () => {
    // Left cell right = 1pt red; right cell left = 3pt blue ⇒ blue (wider) wins.
    const strokes = render(tableOf([
      [cell({ borderR: ln({ width: EMU, color: 'ff0000' }) }),
       cell({ borderL: ln({ width: 3 * EMU, color: '0000ff' }) })],
    ], [COL, COL]));
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('0000ff')); // wider blue wins
    expect(Math.abs(shared[0].width - 3)).toBeLessThan(0.6); // winner's 3pt width
  });

  it('null on one side leaves the OTHER side visible (rule #0)', () => {
    // Left cell right = null (no line); right cell left = 1pt. The real line shows.
    const strokes = render(tableOf([
      [cell({ borderR: null }), cell({ borderL: ln({ color: '112233' }) })],
    ], [COL, COL]));
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('112233'));
  });

  it('null on BOTH sides suppresses the shared gridline entirely', () => {
    const strokes = render(tableOf([
      [cell({ borderR: null }), cell({ borderL: null })],
    ], [COL, COL]));
    expect(verticalAt(strokes, 60)).toHaveLength(0);
  });

  it('shared HORIZONTAL gridline resolves between stacked rows (wider wins)', () => {
    // Row0 cell bottom = 1pt red; row1 cell top = 3pt green ⇒ green wins. Shared
    // line at y=20 (rows 20pt tall).
    const strokes = render(tableOf([
      [cell({ borderB: ln({ width: EMU, color: 'ff0000' }) })],
      [cell({ borderT: ln({ width: 3 * EMU, color: '00ff00' }) })],
    ], [COL]));
    const shared = horizontalAt(strokes, 20);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('00ff00')); // wider green wins
  });

  it('outer table edges are unaffected (non-regression: each drawn once)', () => {
    // Single-cell table with all four outer borders → 4 edges, each once.
    const strokes = render(tableOf([
      [cell({ borderT: ln(), borderB: ln(), borderL: ln(), borderR: ln() })],
    ], [COL]));
    expect(verticalAt(strokes, 0)).toHaveLength(1);   // outer left
    expect(verticalAt(strokes, 60)).toHaveLength(1);  // outer right
    expect(horizontalAt(strokes, 0)).toHaveLength(1); // outer top
    expect(horizontalAt(strokes, 20)).toHaveLength(1);// outer bottom
  });

  it('paints cmpd=dbl table borders as two separated rails', () => {
    const recording = renderRecording(tableOf([
      [cell({ borderB: ln({ width: 3 * EMU, color: '1D6FA8', cmpd: 'dbl' }) })],
    ], [COL]));

    expect(horizontalAt(recording.strokes, 20)).toHaveLength(0);
    const rails = recording.fills.filter((fill) => fill.color === rgba('1D6FA8'));
    expect(rails).toEqual([
      { x: 0, y: 19, width: 60, height: 1, color: rgba('1D6FA8') },
      { x: 0, y: 21, width: 60, height: 1, color: rgba('1D6FA8') },
    ]);
  });

  it('gridSpan: a horizontally-merged cell spans the shared vertical line region', () => {
    // Row0: one cell spanning both columns (gridSpan=2) + hMerge continuation.
    // Row1: two separate cells sharing a vertical line at x=60.
    // The spanning cell has NO interior vertical line inside it (x=60 within the
    // merged cell must not be drawn from row0), and row1's shared line at x=60 is
    // drawn once. Also the horizontal line between the rows at y=20 under x∈[0,60]
    // and x∈[60,120] must not double up where row0's single wide cell meets the
    // two row1 cells.
    const strokes = render(tableOf([
      [cell({ gridSpan: 2, borderT: ln(), borderB: ln() }), cell({ hMerge: true })],
      [cell({ borderR: ln(), borderB: ln() }), cell({ borderL: ln(), borderB: ln() })],
    ], [COL, COL]));
    // Row1 interior vertical shared line at x=60 drawn exactly once.
    expect(verticalAt(strokes, 60)).toHaveLength(1);
  });

  it('rtl (§21.1.3.13): shared vertical gridline still drawn once with the winner', () => {
    // A right-to-left table places logical column 0 at the physical RIGHT. Two
    // cells, logical [a, b]: a is physically RIGHT (x=60..120), b physically LEFT
    // (x=0..60), so the shared interior line is at x=60. At that line a's PHYSICAL-
    // left edge is its logical-RIGHT border (a.borderR), and b's PHYSICAL-right
    // edge is its logical-LEFT border (b.borderL) — these face each other.
    // a.borderR=1pt vs b.borderL=3pt ⇒ the wider (b's) wins, drawn once.
    const t = tableOf([
      [cell({ borderR: ln({ width: EMU, color: 'ff0000' }) }),
       cell({ borderL: ln({ width: 3 * EMU, color: '0000ff' }) })],
    ], [COL, COL]);
    t.rtl = true;
    const strokes = render(t);
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('0000ff')); // wider (b's borderR) wins
    // Outer physical edges (x=0 = b's physical left / logical right side of grid,
    // x=120 = a's physical right) carry no authored border here, so none drawn.
    expect(verticalAt(strokes, 0)).toHaveLength(0);
    expect(verticalAt(strokes, 120)).toHaveLength(0);
  });

  it('rowSpan: a vertically-merged cell spans the shared horizontal line region', () => {
    // Col0 row0 rowSpan=2 (+ vMerge continuation in row1); col1 has two cells that
    // share a horizontal line at y=20. The merged cell must not draw an interior
    // horizontal line at y=20 within its own footprint.
    const strokes = render(tableOf([
      [cell({ rowSpan: 2, borderR: ln() }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell({ borderT: ln() })],
    ], [COL, COL]));
    // Col1's shared horizontal line at y=20 (x in [60,120]) drawn exactly once.
    const shared = horizontalAt(strokes, 20).filter((s) => Math.min(s.x1, s.x2) >= 59);
    expect(shared).toHaveLength(1);
  });
  it('#824: a gridSpan cell resolves EACH below sub-segment against its own neighbour', () => {
    // A cell spanning both columns (gridSpan=2, no bottom border) faces TWO below
    // cells with DIFFERENT top borders. The shared horizontal edge must be split at
    // the column boundary: left half → the left neighbour's border, right half →
    // the right neighbour's. Before the fix only the span-origin (left) neighbour
    // was consulted and the right half's border was dropped.
    const strokes = render(tableOf([
      [cell({ gridSpan: 2 }), cell({ hMerge: true })],
      [cell({ borderT: ln({ width: EMU, color: '111111' }) }),
       cell({ borderT: ln({ width: 3 * EMU, color: '222222' }) })],
    ], [COL, COL]));
    const seg = horizontalAt(strokes, 20);
    const left = seg.filter((s) => Math.min(s.x1, s.x2) < 30);
    const right = seg.filter((s) => Math.max(s.x1, s.x2) > 90);
    expect(left.some((s) => s.color === rgba('111111'))).toBe(true);
    expect(right.some((s) => s.color === rgba('222222'))).toBe(true);
  });

  it('#824: a rowSpan cell resolves EACH right sub-segment against its own neighbour', () => {
    // A cell spanning both rows (rowSpan=2, no right border) faces TWO right
    // neighbours with DIFFERENT left borders. Its shared vertical edge must be split
    // at the row boundary: top half → the upper neighbour's border, bottom half →
    // the lower neighbour's. Before the fix only the span-origin (upper) neighbour
    // was consulted and the bottom half's border was dropped.
    const strokes = render(tableOf([
      [cell({ rowSpan: 2 }), cell({ borderL: ln({ width: EMU, color: '111111' }) })],
      [cell({ vMerge: true }), cell({ borderL: ln({ width: 3 * EMU, color: '222222' }) })],
    ], [COL, COL]));
    const seg = verticalAt(strokes, 60);
    const top = seg.filter((s) => Math.min(s.y1, s.y2) < 10);
    const bot = seg.filter((s) => Math.max(s.y1, s.y2) > 30);
    expect(top.some((s) => s.color === rgba('111111'))).toBe(true);
    expect(bot.some((s) => s.color === rgba('222222'))).toBe(true);
  });

});

describe('PowerPoint table-cell diagonal borders', () => {
  it('strokes the parser-named diagonalTl / diagonalTr lines', () => {
    // The parser JSON names these `diagonalTl` / `diagonalTr` (a:lnTlToBr /
    // a:lnBlToTr); a renderer reading any other spelling draws nothing.
    const t = tableOf([[cell({ diagonalTl: ln({ color: 'FF0000' }), diagonalTr: ln({ color: '0000FF' }) })]], [COL]);
    const diagonals = render(t).filter((s) => Math.abs(s.x1 - s.x2) > 0.5 && Math.abs(s.y1 - s.y2) > 0.5);
    expect(diagonals.map((s) => s.color).sort()).toEqual([rgba('0000FF'), rgba('FF0000')].sort());
    const tl = diagonals.find((s) => s.color === rgba('FF0000'));
    expect(tl && Math.sign(tl.x2 - tl.x1) === Math.sign(tl.y2 - tl.y1)).toBe(true);
  });
});
