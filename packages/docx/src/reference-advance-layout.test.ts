import { describe, expect, it } from 'vitest';
import { TEXT_ADVANCE_REFERENCE_PX } from '@silurus/ooxml-core';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import { createLayoutServices } from './layout-runtime.js';
import { verticalGlyphMeasurementServiceOf } from './layout/runtime-state.js';
import { buildDocxTextLayer } from './text-layer.js';
import { buildDocxHighlightLayer } from './find-highlight-layer.js';
import {
  gridFittedWidth,
  linearAdvancePx,
  makeGridFittingCanvas,
  type PaintedText,
} from './test-support/grid-fitting-canvas.test-support.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// Hosts that grid-fit small sizes (core reference-advances.ts) must still lay
// out from proportional advances, as Word does, and paint the glyphs on those
// same advances. The double rounds every half-width cell below 64px, so a
// full-width glyph at 11px measures 12, while Word's advance is 11.

/** ECMA-376 §17.6.5 docGrid linesAndChars: pitch = 11pt + 1382/4096pt. */
const CHAR_SPACE = 1382;
const COLUMN_PT = 453.5;

function paragraph(text: string, fontSize: number): BodyElement {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize, color: null, fontFamily: 'GridFitted', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: fontSize, defaultFontFamily: 'GridFitted', widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function model(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: COLUMN_PT + 144, pageHeight: 400,
    marginTop: 36, marginRight: 72, marginBottom: 36, marginLeft: 72,
    headerDistance: 18, footerDistance: 18, titlePage: false, evenAndOddHeaders: false,
    docGridType: 'linesAndChars', docGridLinePitch: 18, docGridCharSpace: CHAR_SPACE,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(doc: DocxDocumentModel) {
  const host = makeGridFittingCanvas();
  const runs: DocxTextRunInfo[] = [];
  // 1 CSS px per pt: run geometry is in page points.
  await renderDocumentToCanvas(doc, host.canvas, 0, {
    dpr: 1, width: doc.section.pageWidth, onTextRun: (run) => runs.push(run),
  });
  return { ...host, runs };
}

/** Page-space x of a painted string's anchor and its painted advance. */
function paintedExtent(fill: PaintedText): { xPt: number; advancePt: number } {
  return {
    xPt: fill.matrix.a * fill.x + fill.matrix.c * fill.y + fill.matrix.e,
    advancePt: gridFittedWidth(fill.text, fill.font, fill.letterSpacing) * fill.matrix.a,
  };
}

describe('reference-em text advances on a grid-fitting host', () => {
  it('fits 38 full-width characters of an 11pt grid line into a 453.5pt column', async () => {
    const text = '申請者氏名住所'.repeat(5) + '申請者'; // 38 full-width characters
    expect([...text]).toHaveLength(38);
    const { runs, fills } = await render(model([paragraph(text, 11)]));

    // One line: the grid pitch is Word's 11 + 1382/4096 pt, not 12 + 1382/4096.
    expect(new Set(runs.map((run) => run.y)).size).toBe(1);
    const pitch = 11 + CHAR_SPACE / 4096;
    const width = runs.reduce((sum, run) => sum + run.w, 0);
    expect(width).toBeCloseTo(38 * pitch, 6);
    expect(width).toBeLessThanOrEqual(COLUMN_PT);

    // Paint draws at the reference em under the linear scale, and each painted
    // string starts on its run and advances exactly by the run's width.
    const textFills = fills.filter((fill) => runs.some((run) => run.text === fill.text));
    expect(textFills.length).toBe(runs.length);
    for (const fill of textFills) {
      const run = runs.find((candidate) => candidate.text === fill.text)!;
      expect(fill.font).toContain(`${TEXT_ADVANCE_REFERENCE_PX}px`);
      expect(fill.matrix.a).toBeCloseTo(11 / TEXT_ADVANCE_REFERENCE_PX, 12);
      const painted = paintedExtent(fill);
      expect(painted.xPt).toBeCloseTo(run.x, 6);
      expect(painted.advancePt).toBeCloseTo(run.w, 6);
    }
  });

  it('keeps measuring and painting at the authored size where the host is proportional', async () => {
    // Every cell is a whole pixel at 12px, so the host is proportional there.
    const { runs, fills } = await render(model([paragraph('申請者', 12)]));
    const fill = fills.find((candidate) => candidate.text === '申請者')!;
    expect(fill.font).toContain('12px');
    expect(fill.matrix.a).toBe(1);
    expect(runs[0]!.w).toBeCloseTo(3 * (12 + CHAR_SPACE / 4096), 6);
    // Overlays follow the same decision: the run keeps its plain font.
    expect(runs[0]!.referenceFont).toBeUndefined();
  });

  it('plans vertical glyph cells from the same reference em', () => {
    const host = makeGridFittingCanvas();
    const services = createLayoutServices(model([paragraph('申', 11)]), { measureContext: host.ctx });
    const cells = verticalGlyphMeasurementServiceOf(services).planRun({
      text: '申請',
      font: 'normal 400 11px GridFitted',
      fontKerning: 'none',
      fontSizePt: 11,
      letterSpacingPt: 0,
      charScale: 1,
      growTrRotateInk: false,
      writingMode: 'vertical-rl',
    });
    expect(cells.map((cell) => cell.advancePt)).toEqual([
      linearAdvancePx('申', 11),
      linearAdvancePx('請', 11),
    ]);
  });

  it('lays out the selection and find overlays from the reference em', async () => {
    const { runs } = await render(model([paragraph('申請者', 11)]));
    const run = runs[0]!;
    const k = 11 / TEXT_ADVANCE_REFERENCE_PX;
    expect(run.referenceFont).toEqual({
      font: run.font.replace('11px', `${TEXT_ADVANCE_REFERENCE_PX}px`),
      scale: k,
    });

    const spans: { style: { cssText: string } }[] = [];
    const layer = {
      ownerDocument: {
        createElement: () => ({ style: { cssText: '' }, dataset: {}, addEventListener() {} }),
      },
      appendChild: (child: { style: { cssText: string } }) => spans.push(child),
      dataset: {},
      set innerHTML(_value: string) { spans.length = 0; },
    };
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [run], 600, 400);
    expect(spans[0]!.style.cssText).toContain(`font:${run.referenceFont!.font};`);
    expect(spans[0]!.style.cssText).toContain(`line-height:${run.h / k}px;`);
    expect(spans[0]!.style.cssText).toContain(`transform:scale(${k});transform-origin:top left;`);

    // Find highlight extents measure the reference font and scale it back.
    const boxes: { style: { cssText: string } }[] = [];
    const highlight = {
      appendChild: (child: { style: { cssText: string } }) => boxes.push(child),
      set innerHTML(_value: string) { boxes.length = 0; },
    };
    const previousDocument = (globalThis as { document?: unknown }).document;
    (globalThis as { document?: unknown }).document = {
      createElement: () => ({ style: { cssText: '' } }),
    };
    try {
      buildDocxHighlightLayer(
        highlight as unknown as HTMLDivElement,
        [run],
        [{ slices: [{ runIndex: 0, start: 1, end: 3 }], active: false }],
        1, 1,
        (font) => (s) => gridFittedWidth(s, font),
      );
    } finally {
      (globalThis as { document?: unknown }).document = previousDocument;
    }
    // Two linear 11pt cells after one, plus the grid pitch carried as tracking
    // (as percentages of a 1px box).
    const pitch = run.letterSpacingPx!;
    expect(pitch).toBeCloseTo(CHAR_SPACE / 4096, 12);
    const left = Number(/left:([\d.]+)%/.exec(boxes[0]!.style.cssText)![1]) / 100;
    const width = Number(/width:([\d.]+)%/.exec(boxes[0]!.style.cssText)![1]) / 100;
    expect(left).toBeCloseTo(run.x + 11 + pitch, 9);
    expect(width).toBeCloseTo(22 + pitch, 9);
  });
});
