import { describe, it, expect } from 'vitest';
import type {
  BodyElement, DocParagraph, DocxTextRun, DocxDocumentModel, RunRevision, SectionProps,
} from './types';
import { attachDocumentLayoutVariants, selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';

// ECMA-376 §17.13.5 tracked-change views. The DEFAULT layout variant
// (showTrackedChanges absent/false) is the FINAL view: deleted (`w:del`,
// §17.13.5.14) and moved-away (`w:moveFrom`, §17.13.5.22) runs produce no
// segments, so line breaking and pagination see the accepted document state.
// The markup variant (`showTrackedChanges: true`) keeps every revision run
// visible. A document without revisions must produce identical geometry in
// both variants — the axis is inert unless revision runs exist.

// Deterministic stub canvas: glyph advance = charCount × fontPx, font box =
// 0.8/0.2 em (a single line is exactly fontPx tall). Copied from
// per-section-page-geometry.test.ts.
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
function textRun(text: string, revision?: RunRevision): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
    ...(revision ? { revision } : {}),
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

function doc(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 600, pageHeight: 400,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function revisionDoc(): DocxDocumentModel {
  return doc([
    para(
      textRun('kept '),
      textRun('added', { kind: 'insertion', author: 'Alice' }),
      textRun('gone', { kind: 'deletion', author: 'Alice' }),
      textRun('moved-away', { kind: 'moveFrom', author: 'Bob' }),
      textRun('moved-in', { kind: 'moveTo', author: 'Bob' }),
    ),
  ]);
}

function pageTexts(model: DocxDocumentModel, showTrackedChanges?: boolean): string[] {
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: makeCtx() }),
    { currentDateMs: 0, ...(showTrackedChanges === undefined ? {} : { showTrackedChanges }) },
  );
  return layout.pages.flatMap((page) =>
    page.layers.body
      .filter((node) => node.kind === 'paragraph')
      .map((node) => node.kind === 'paragraph'
        ? node.lines.flatMap((line) => line.placements)
            .filter((placement) => placement.kind === 'text')
            .map((placement) => placement.text).join('')
        : ''));
}

describe('showTrackedChanges layout axis (§17.13.5)', () => {
  it('defaults to the final view: deletions and moveFrom are hidden, insertions and moveTo render plain', () => {
    expect(pageTexts(revisionDoc())).toEqual(['kept addedmoved-in']);
  });

  it('an explicit false matches the absent default', () => {
    expect(pageTexts(revisionDoc(), false)).toEqual(['kept addedmoved-in']);
  });

  it('the markup view keeps every revision run visible', () => {
    expect(pageTexts(revisionDoc(), true)).toEqual(['kept addedgonemoved-awaymoved-in']);
  });

  it('a document without revisions lays out identically in both variants', () => {
    const model = doc([para(textRun('one two three')), para(textRun('four'))]);
    const services = createLayoutServices(model, { measureContext: makeCtx() });
    const finalView = layoutDocument(model, services, { currentDateMs: 0 });
    const markupView = layoutDocument(model, services, {
      currentDateMs: 0,
      showTrackedChanges: true,
    });
    expect(JSON.parse(JSON.stringify(markupView.pages)))
      .toEqual(JSON.parse(JSON.stringify(finalView.pages)));
  });

  it('the two views are distinct cached layout variants selected per render', () => {
    const model = revisionDoc();
    const services = createLayoutServices(model, { measureContext: makeCtx() });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model),
      services,
      defaultCurrentDateMs: 0,
      buildLayout: (options) => layoutDocument(model, services, options),
    });
    const finalView = selectDocumentLayoutPage(services, { defaultCurrentDateMs: 0 }, 0);
    const markupView = selectDocumentLayoutPage(
      services,
      { defaultCurrentDateMs: 0, showTrackedChanges: true },
      0,
    );
    expect(markupView.key).not.toBe(finalView.key);
    // Same-flag re-selection is a cache hit (object identity), and the default
    // variant survives markup selection (default + one non-default retained).
    expect(selectDocumentLayoutPage(services, { defaultCurrentDateMs: 0 }, 0).layout)
      .toBe(finalView.layout);
    expect(
      selectDocumentLayoutPage(services, { defaultCurrentDateMs: 0, showTrackedChanges: true }, 0)
        .layout,
    ).toBe(markupView.layout);
  });
});
