import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocumentLayout, LayoutServices } from './layout/types.js';
import { layoutLines, type LayoutSeg } from './line-layout.js';
import { stableFingerprint } from './layout/fingerprint.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
  } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Phase 4-1 B2 Stage 2 — ZOOM-INVARIANT line breaking.
//
// Word lays text out in the document's own coordinate space and treats the
// display scale as a viewport transform: the line PARTITION (which glyphs land
// on which line) is the same at every zoom. Stage 2 realises this by reusing the
// paginator's scale-1 line PARTITION at ANY paint scale — mapping ordinary body
// glyph geometry through a viewport transform without re-deciding wrap points —
// instead of re-running layoutLines' break
// decisions at the paint scale, which under a real (hinted) font whose glyph
// advance is NOT proportional to the font px size would shift wrap points as the
// zoom changes (documented in layout-lines-scale-invariance.test.ts).
//
// This suite drives the REAL render path (renderParagraph) through a SUB-LINEAR
// mock canvas — glyph advance shrinks per-px as the size grows, the direction
// real hinting bends — and asserts the drawn per-line text partition is
// IDENTICAL at scale 1 and scale 0.75. A control proves the test is non-vacuous:
// fed the SAME segments, layoutLines' OWN break decisions wrap differently at the
// two scales under this font, so the invariance is a real property of the reuse.
// ─────────────────────────────────────────────────────────────────────────────

/** Sub-linear glyph advance: `px·(perPx − shrink·px)` per code point. NOT
 *  proportional to px, so a scale-1 layout and a scale-s re-layout would wrap
 *  differently — the exact hinting non-linearity Stage 2 must make invisible to
 *  the line partition. The SAME metric backs the paginate ctx (OffscreenCanvas)
 *  and every paint ctx, so scale-1 paginate and scale-1 paint agree. */
function subLinearWidth(text: string, fontPx: number): number {
  const per = Math.max(0.01, fontPx * (0.5 - 0.006 * fontPx));
  return [...text].length * per;
}

function makeMeasureStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: subLinearWidth(s, p),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeMeasureStubCtx(); }
};

interface Draw { text: string; x: number; y: number; right: number; font: string; scaleX: number; }

/** A recording canvas backed by the SAME sub-linear metric. Records every text
 *  draw with its baseline y so the caller can reconstruct the per-line partition
 *  (glyphs sharing a baseline belong to one line). */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; draws: Draw[]; measures: () => number } {
  let font = '10px serif';
  let letterSpacing = '0px';
  let measures = 0;
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const stack: typeof transform[] = [];
  const draws: Draw[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      measures++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: subLinearWidth(s, p),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ ...transform }); },
    restore() { transform = stack.pop() ?? transform; },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(x: number, y: number) {
      transform.scaleX *= x;
      transform.scaleY *= y;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) {
      if (!s) return;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const worldX = transform.translateX + transform.scaleX * x;
      const worldY = transform.translateY + transform.scaleY * y;
      const spacing = parseFloat(letterSpacing) || 0;
      const localAdvance =
        subLinearWidth(s, p) + Math.max(0, [...s].length - 1) * spacing;
      draws.push({
        text: s,
        x: worldX,
        y: worldY,
        right: worldX + transform.scaleX * localAdvance,
        font,
        scaleX: transform.scaleX,
      });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, draws, measures: () => measures };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

// A SHORT page (content height ≈ 40pt ⇒ ~3 lines) so the long paragraph SPLITS
// across pages — the paginator only stamps its scale-1 lines on a split
// (splitParagraphAcrossPages), which is what arms the reuse path.
function doc(body: BodyElement[], pageHeight = 60): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function oneCellTable(paragraph: DocParagraph): BodyElement {
  const border = { style: 'single', color: '000000', width: 0.5 };
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...paragraph } as CellElement],
    colSpan: 1,
    vMerge: null,
    borders: { top: border, bottom: border, left: border, right: border, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: 180,
  } as DocTableCell;
  const row: DocTableRow = {
    cells: [cell],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as DocTableRow;
  const table: DocTable = {
    type: 'table',
    colWidths: [180],
    rows: [row],
    borders: { top: border, bottom: border, left: border, right: border, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    // A negative leading indent gate-excludes fragment paint, so the legacy
    // table layout must share the canonical paragraph contract.
    tblInd: -5,
    layout: 'fixed',
  } as DocTable;
  return table as unknown as BodyElement;
}

/** Reconstruct the per-line TEXT partition from a paint stream: group draws by
 *  their baseline y (rounded to absorb sub-px), then concatenate each line's text
 *  in draw order. Independent of the absolute y values, so it compares across
 *  scales. */
function linePartition(draws: Draw[]): string[] {
  const byLine = new Map<number, { x: number; text: string }[]>();
  for (const d of draws) {
    const key = Math.round(d.y * 100) / 100;
    if (!byLine.has(key)) byLine.set(key, []);
    byLine.get(key)!.push({ x: d.x, text: d.text });
  }
  return [...byLine.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, segs]) => segs.sort((a, b) => a.x - b.x).map((s) => s.text).join(''));
}

/** Render every page of `model` at `width` (⇒ paint scale = width/pageWidth) and
 *  return the concatenated per-line text partition across all pages. */
async function partitionAtWidth(
  model: DocxDocumentModel,
  layout: DocumentLayout,
  services: LayoutServices,
  width: number,
): Promise<string[]> {
  const all: string[] = [];
  for (let p = 0; p < layout.pages.length; p++) {
    const rec = makeRecordingCanvas();
    await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, width, layoutServices: services });
    all.push(...linePartition(rec.draws));
  }
  return all;
}

function canonical(model: DocxDocumentModel) {
  const services = createLayoutServices(model);
  return { services, layout: layoutDocument(model, services, { currentDateMs: 0 }) };
}

describe('zoom-invariant line breaking (Phase 4-1 B2 Stage 2)', () => {
  it('a long paragraph draws the SAME per-line partition at scale 1 and scale 0.75', async () => {
    // 160 single-letter words → many wrap points; page width 200pt so it wraps
    // and splits across pages. Sub-linear font ⇒ a re-layout at 0.75 would wrap
    // differently, but the scale-1 stamp reuse pins the partition.
    const text = Array.from({ length: 160 }, (_, i) => `w${i % 10}`).join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(1); // actually split

    const at1 = await partitionAtWidth(model, layout, services, 200);       // scale 1.0
    const at075 = await partitionAtWidth(model, layout, services, 150);      // scale 0.75
    expect(at075).toEqual(at1);
    // Non-vacuous: the paragraph really wrapped to multiple lines.
    expect(at1.length).toBeGreaterThan(3);
  });

  it('reuses the retained external-link partition at every paint zoom', async () => {
    const url = [
      'https://example.test/path/a-long-document-name.pdf',
      'more/path/with-separators-and-query.pdf?part=one&part=two',
      'final/path/another-long-document-name.pdf',
    ].join('/');
    const paragraph = para(url);
    Object.assign(paragraph.runs[0]!, {
      isLink: true,
      underline: true,
      hyperlink: url,
    });
    const model = doc([paragraph as unknown as BodyElement]);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(1);

    const at1 = await partitionAtWidth(model, layout, services, 200);
    const at05 = await partitionAtWidth(model, layout, services, 100);
    expect(at05).toEqual(at1);
    expect(at1.join('')).toBe(url);
    expect(at1.length).toBeGreaterThan(3);
  });

  it('CJK per-glyph wrap: same partition at scale 1 and scale 0.5', async () => {
    const text = 'あ'.repeat(240);
    const model = doc([para(text) as unknown as BodyElement]);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const at1 = await partitionAtWidth(model, layout, services, 200);   // scale 1.0
    const at05 = await partitionAtWidth(model, layout, services, 100);  // scale 0.5
    expect(at05).toEqual(at1);
  });

  it('keeps a justified CJK line inside its right paragraph border at every scale', async () => {
    const borderSpacePt = 4;
    const p = para('あ'.repeat(240), {
      alignment: 'justify',
      borders: {
        top: null,
        bottom: null,
        left: null,
        right: { style: 'single', color: '000000', width: 0.5, space: borderSpacePt },
        between: null,
      },
    });
    const model = doc([p as unknown as BodyElement]);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(1);

    const paintScale = 0.5;
    const runBoxes: { x: number; w: number; font: string; fontSize: number }[] = [];
    const glyphDraws: Draw[] = [];
    for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
      const rec = makeRecordingCanvas();
      await renderDocumentToCanvas(model, rec.canvas, pageIndex, {
        dpr: 1,
        width: model.section.pageWidth * paintScale,
        layoutServices: services,
        onTextRun: ({ x, w, font, fontSize }) => runBoxes.push({ x, w, font, fontSize }),
      });
      glyphDraws.push(...rec.draws);
    }

    expect(runBoxes.length).toBeGreaterThan(0);
    const rightBorderX =
      (model.section.pageWidth - model.section.marginRight + borderSpacePt) * paintScale;
    const rightmostRunX = Math.max(...runBoxes.map(({ x, w }) => x + w));
    expect(rightmostRunX).toBeLessThanOrEqual(rightBorderX + 1e-6);
    expect(Math.max(...glyphDraws.map(({ right }) => right))).toBeLessThanOrEqual(rightBorderX + 1e-6);
    // Top-level body paragraphs intentionally have an empty story-container
    // chain and still use the canonical document-space glyph transform.
    expect(glyphDraws.some(({ font, scaleX }) =>
      font.includes('10px') && Math.abs(scaleX - paintScale) < 1e-9)).toBe(true);
    expect(runBoxes.every(({ font, fontSize }) => font.includes('5px') && fontSize === 5)).toBe(true);
  });

  it('keeps justified CJK inside a negative-indent legacy body table cell at every scale', async () => {
    const model = doc([
      oneCellTable(para('あ'.repeat(240), { alignment: 'justify' })),
    ], 600);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(0);

    const paintScale = 0.5;
    const glyphDraws: Draw[] = [];
    const runRights: number[] = [];
    for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
      const rec = makeRecordingCanvas();
      await renderDocumentToCanvas(model, rec.canvas, pageIndex, {
        dpr: 1,
        width: model.section.pageWidth * paintScale,
        layoutServices: services,
        onTextRun: ({ x, w }) => runRights.push(x + w),
      });
      glyphDraws.push(...rec.draws);
    }

    expect(glyphDraws.length).toBeGreaterThan(0);
    const tableRightX =
      (model.section.marginLeft - 5 + 180) * paintScale;
    expect(Math.max(...runRights)).toBeLessThanOrEqual(tableRightX + 1e-6);
    expect(Math.max(...glyphDraws.map(({ right }) => right))).toBeLessThanOrEqual(tableRightX + 1e-6);
    expect(glyphDraws.some(({ font, scaleX }) =>
      font.includes('10px') && Math.abs(scaleX - paintScale) < 1e-9)).toBe(true);
  });

  it.each([
    ['numbering marker', () => para('body', {
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 18,
        tab: 18, suff: 'tab', jc: 'left',
      },
    })],
    ['tab segment', () => para('body\ttail', {
      tabStops: [{ pos: 60, alignment: 'left', leader: 'dot' }],
    })],
    ['math segment', () => {
      const paragraph = para('body');
      paragraph.runs.push({
        type: 'math', display: false, fontSize: 10,
        nodes: [{ kind: 'run', text: 'x+1', style: 'italic' }],
      } as DocParagraph['runs'][number]);
      return paragraph;
    }],
  ])('paints retained %s geometry only through the device transform', async (_name, makeParagraph) => {
    const model = doc([makeParagraph() as unknown as BodyElement], 600);
    const { layout, services } = canonical(model);
    const retained = layout.pages[0].layers.body[0];
    const fingerprint = stableFingerprint('retained-zoom', retained);

    const point = makeRecordingCanvas();
    await renderDocumentToCanvas(model, point.canvas, 0, {
      dpr: 1,
      width: model.section.pageWidth,
      layoutServices: services,
    });
    const device = makeRecordingCanvas();
    await renderDocumentToCanvas(model, device.canvas, 0, {
      dpr: 1,
      width: model.section.pageWidth * 0.5,
      layoutServices: services,
    });

    expect(point.measures()).toBe(0);
    expect(device.measures()).toBe(0);
    expect(stableFingerprint('retained-zoom', retained)).toBe(fingerprint);
    const pointBody = point.draws.find((draw) => draw.text === 'body');
    const deviceBody = device.draws.find((draw) => draw.text === 'body');
    expect(pointBody).toBeDefined();
    expect(deviceBody).toBeDefined();
    expect(pointBody!.font).toContain('10px');
    expect(deviceBody!.font).toBe(pointBody!.font);
    expect(pointBody!.scaleX).toBe(1);
    expect(deviceBody!.scaleX).toBe(0.5);
    expect(deviceBody!.x).toBe(pointBody!.x * 0.5);
    expect(deviceBody!.y).toBe(pointBody!.y * 0.5);
  });

  it('CONTROL: the mock font is genuinely NON-linear — a paint-scale re-layout WOULD wrap differently', () => {
    // Non-vacuity: proves the invariance above is a real property of Stage 2, not
    // an artefact of a metric that happens to be scale-clean. Feeding the SAME
    // segments to layoutLines at scale 1 vs scale 0.75 with this sub-linear
    // advance yields a DIFFERENT line PARTITION (fewer glyphs fit per line at the
    // smaller px size) — exactly the paint-scale re-break Stage 2 removed by
    // reusing the scale-1 stamp. If this ever passes trivially the render-level
    // assertions above would be meaningless.
    const seg = (text: string): LayoutSeg => ({
      text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', vertAlign: null,
      measuredWidth: 0,
    } as unknown as LayoutSeg);
    const text = Array.from({ length: 200 }, (_, i) => `w${i % 10}`).join(' ');

    // Same real content width (180pt) at two scales: scale-1 box 180, scale-0.75
    // box 135, first-indent 0. A linear font would give the SAME partitions; the
    // sub-linear font does not.
    const a = layoutLines(makeMeasureStubCtx(), [seg(text)], 180, 0, 1);
    const b = layoutLines(makeMeasureStubCtx(), [seg(text)], 135, 0, 0.75);
    expect(a.length).toBeGreaterThan(1);
    const partitions = (lines: typeof a) => lines.map((line) => line.segments
      .filter((segment): segment is Extract<LayoutSeg, { text: string }> => 'text' in segment)
      .map((segment) => segment.text)
      .join(''));
    // The exact partitions differ under the non-linear advance. They may happen
    // to produce the same total line count after emergency wrapping consumes a
    // preceding line remainder, so compare the retained line content directly.
    expect(partitions(b)).not.toEqual(partitions(a));
  });

  it('an ANCHORED image in the paragraph adds no inline advance at any scale', async () => {
    // Anchored images live out of inline flow: layoutLines pins measuredWidth=0
    // and never adds them to line.segments (they are drawn by renderAnchorImages).
    // Characterizes that the reuse/rehydration path preserves this: the partition
    // is zoom-invariant AND every line's first glyph starts at the paragraph
    // origin (∝ scale) — not shifted right by imageWidth·scale, which is what a
    // rescale that conjured an inline advance for the anchor would produce.
    const text = Array.from({ length: 120 }, (_, i) => `w${i % 10}`).join(' ');
    const p = para(text);
    p.runs.unshift({
      type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png',
      widthPt: 50, heightPt: 40, anchor: true, anchorXPt: 30, anchorYPt: 20,
    } as unknown as DocParagraph['runs'][number]);
    const model = doc([p as unknown as BodyElement]);
    const { layout, services } = canonical(model);
    expect(layout.pages.length).toBeGreaterThan(1);

    const collect = async (width: number): Promise<{ partition: string[]; firstX: number[] }> => {
      const partition: string[] = [];
      const firstX: number[] = [];
      for (let pg = 0; pg < layout.pages.length; pg++) {
        const rec = makeRecordingCanvas();
        await renderDocumentToCanvas(model, rec.canvas, pg, { dpr: 1, width, layoutServices: services });
        const byLine = new Map<number, Draw[]>();
        for (const d of rec.draws) {
          const key = Math.round(d.y * 100) / 100;
          if (!byLine.has(key)) byLine.set(key, []);
          byLine.get(key)!.push(d);
        }
        for (const [, draws] of [...byLine.entries()].sort((a, b) => a[0] - b[0])) {
          draws.sort((a, b) => a.x - b.x);
          partition.push(draws.map((d) => d.text).join(''));
          firstX.push(draws[0].x);
        }
      }
      return { partition, firstX };
    };

    const at1 = await collect(200);   // scale 1.0
    const at075 = await collect(150); // scale 0.75
    expect(at075.partition).toEqual(at1.partition);
    // Line starts scale linearly (page origin ∝ scale): x@0.75 = x@1 × 0.75.
    // A conjured 50pt inline advance would add 37.5px here and fail.
    expect(at075.firstX.length).toBe(at1.firstX.length);
    for (let i = 0; i < at1.firstX.length; i++) {
      expect(at075.firstX[i]).toBeCloseTo(at1.firstX[i] * 0.75, 6);
    }
  });
});
