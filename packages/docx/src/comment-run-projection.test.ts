import { describe, it, expect } from 'vitest';
import type {
  BodyElement, DocParagraph, DocxTextRun, DocxDocumentModel, SectionProps,
} from './types';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunsForPage } from './text-run-projection.js';

// ECMA-376 §17.13.4 consumer UIs join projected run geometry back to
// model runs via (source.path, sourceRunIndex). Pin that both survive the
// projection for every text run, so `commentAnchorRanges()` intervals
// can be mapped onto `DocxTextRunInfo` geometry without layout threading.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(...runs: DocRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

describe('comment-overlay run projection join keys', () => {
  it('every projected text run carries its paragraph path and sourceRunIndex', () => {
    const section: SectionProps = {
      pageWidth: 600, pageHeight: 400,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const model = {
      section,
      body: [
        para(textRun('alpha '), textRun('beta '), textRun('gamma')),
        para(textRun('second')),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(
      model,
      createLayoutServices(model, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    );
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const byText = new Map(runs.map((run) => [run.text.trim(), run]));
    expect(byText.get('alpha')?.source?.path).toEqual([0]);
    expect(byText.get('alpha')?.sourceRunIndex).toBe(0);
    expect(byText.get('beta')?.sourceRunIndex).toBe(1);
    expect(byText.get('gamma')?.sourceRunIndex).toBe(2);
    expect(byText.get('second')?.source?.path).toEqual([1]);
    expect(byText.get('second')?.sourceRunIndex).toBe(0);
    for (const run of runs) {
      expect(run.sourceRunIndex).toBeTypeOf('number');
      expect(run.source?.story).toBe('body');
    }
  });
});
