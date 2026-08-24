import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
  SectionGeom, HeaderFooter,
  DocTable, DocTableRow, DocTableCell,
} from './types';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';

// ECMA-376 §17.6.13 `<w:pgSz>` + §17.6.11 `<w:pgMar>` — page geometry is
// PER-SECTION. A mid-body SectionBreak carries its ending section's `geom`; the
// paginator stamps each element's `sectionGeom` (upcoming SectionBreak's geom, or
// the body-level section for the final section) so the renderer sizes each page
// from its own section. Single-section documents stamp the body-level geometry
// everywhere (byte-identical layout).

// Deterministic stub canvas: glyph advance = charCount × fontPx, font box =
// 0.8/0.2 em (a single line is exactly fontPx tall). Copied from pagination.test.ts.
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

// Canonical layout builds its default measure ctx from `new OffscreenCanvas(...)`, which
// the node test env lacks. Polyfill it with the SAME deterministic stub (line
// height = fontPx) so the HF-reserve paginator runs headless and its page-fit
// arithmetic is exact. (The paginator-stamp/renderer tests above inject their ctx
// directly through production layout/paint and don't need this.)
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(text: string, fontSize = 20): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(text, fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

// A portrait first section (200×140) ended by a nextPage break, then a landscape
// body section (140×200) via doc.section. `geom` on the break carries the portrait
// section's page size; doc.section carries the landscape body size.
function mixedDoc(): DocxDocumentModel {
  const portrait: SectionGeom = {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0,
  };
  const bodySection: SectionProps = {
    pageWidth: 140, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  const body: BodyElement[] = [
    para('PORTRAIT_SECTION'),
    { type: 'sectionBreak', kind: 'nextPage', geom: portrait } as BodyElement,
    para('LANDSCAPE_SECTION'),
  ];
  return {
    section: bodySection, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function canonicalLayout(doc: DocxDocumentModel) {
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: makeCtx() }), { currentDateMs: 0 });
}

const retainedText = (node: import('./layout/types.js').PaintNode): string => node.kind === 'paragraph'
  ? node.lines.flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text).join('')
  : '';

describe('per-section page geometry (§17.6.13/§17.6.11) — paginator', () => {
  it('stamps each element with its section geometry', () => {
    const doc = mixedDoc();
    const pages = canonicalLayout(doc).pages;
    // Page 0 = portrait section (the first section, ended by the nextPage break).
    expect(pages[0].section.geometry.pageWidth).toBe(200);
    expect(pages[0].section.geometry.pageHeight).toBe(140);
    // Page 1 = landscape body section (no following break ⇒ body-level geometry).
    expect(pages[1].section.geometry.pageWidth).toBe(140);
    expect(pages[1].section.geometry.pageHeight).toBe(200);
  });

  // Height-sensitive spill — proves the const→arrow conversion of the page frame.
  // A first section of 200×140 with margins 20 has a content height of 100pt ⇒ five
  // 20pt paragraphs per page, so a SIXTH paragraph spills to a second page WITHIN the
  // first section. If the frame were still read from the body-level section (140×200,
  // content 160 ⇒ eight 20pt lines fit), all six would stay on page 0 and this test
  // would fail — i.e. it regresses the moment the per-section frame reverts to a const.
  it('paginates each section against ITS OWN page height (const→arrow)', () => {
    const portrait: SectionGeom = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0,
    };
    const bodySection: SectionProps = {
      pageWidth: 140, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    // Six 20pt paragraphs in the portrait first section, then a nextPage break into
    // the landscape body section carrying one paragraph.
    const body: BodyElement[] = [
      para('A1'), para('A2'), para('A3'), para('A4'), para('A5'), para('A6'),
      { type: 'sectionBreak', kind: 'nextPage', geom: portrait } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: bodySection, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    const texts = pages.map((page) =>
      page.layers.body.filter((node) => node.kind === 'paragraph').map(retainedText),
    );
    // Portrait content height 100 ⇒ A1..A5 on page 0, A6 spills to page 1 (still the
    // portrait section), then the break opens page 2 for the landscape section.
    expect(texts).toEqual([['A1', 'A2', 'A3', 'A4', 'A5'], ['A6'], ['B1']]);
    // The spilled A6 still carries the portrait section geometry (it precedes the
    // break); B1 carries the body-level landscape geometry.
    expect(pages[1].section.geometry.pageWidth).toBe(200);
    expect(pages[1].section.geometry.pageHeight).toBe(140);
    expect(pages[2].section.geometry.pageWidth).toBe(140);
    expect(pages[2].section.geometry.pageHeight).toBe(200);
  });

  // pgSz/pgMar are optional children with no normative Letter/one-inch default.
  // Preserve the document page box for omitted fields without changing the
  // final section's authored occurrence.
  it('inherits omitted non-continuous geometry without changing the final section', () => {
    const geom1: SectionGeom = {
      pageWidth: 300, pageHeight: 400,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0,
    };
    const bodySection: SectionProps = {
      pageWidth: 140, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const body: BodyElement[] = [
      para('S1'),
      { type: 'sectionBreak', kind: 'nextPage', geom: geom1 } as BodyElement,
      para('S2'),
      // No `geom`: this section retains the following document page box.
      { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
      para('S3'),
    ];
    const doc = {
      section: bodySection, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    // S1: section ending at break1 ⇒ break1's geom.
    expect(pages[0].section.geometry.pageWidth).toBe(300);
    expect(pages[0].section.geometry.pageHeight).toBe(400);
    // S2: section ending at break2, which has NO geom ⇒ following page box.
    expect(pages[1].section.geometry.pageWidth).toBe(140);
    expect(pages[1].section.geometry.pageHeight).toBe(200);
    // S3: final section ⇒ body-level geometry.
    expect(pages[2].section.geometry.pageWidth).toBe(140);
    expect(pages[2].section.geometry.pageHeight).toBe(200);
  });

  it('does not chain an omitted non-continuous page box through an authored intermediate section', () => {
    const intermediate: SectionGeom = {
      pageWidth: 300, pageHeight: 400,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0,
    };
    const bodySection: SectionProps = {
      pageWidth: 140, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section: bodySection,
      body: [
        para('DOCUMENT_FALLBACK'),
        { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
        para('AUTHORED_INTERMEDIATE'),
        { type: 'sectionBreak', kind: 'nextPage', geom: intermediate } as BodyElement,
        para('FINAL'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;

    const pages = canonicalLayout(doc).pages;
    expect(pages.map(({ section }) => ({
      width: section.geometry.pageWidth,
      height: section.geometry.pageHeight,
    }))).toEqual([
      { width: 140, height: 200 },
      { width: 300, height: 400 },
      { width: 140, height: 200 },
    ]);
  });

  it('inherits omitted non-continuous geometry across an incoming continuous boundary', () => {
    const bodySection: SectionProps = {
      pageWidth: 140, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'continuous',
    } as SectionProps;
    const doc = {
      section: bodySection,
      body: [
        para('INHERITED'),
        { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
        para('CONTINUOUS'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;

    const pages = canonicalLayout(doc).pages;
    expect(pages[0].section.geometry).toMatchObject({ pageWidth: 140, pageHeight: 200 });
    expect(pages).toHaveLength(1);
    expect(pages[0].sectionRegions).toHaveLength(2);
    expect(pages[0].sectionRegions[1]?.section.geometry)
      .toMatchObject({ pageWidth: 140, pageHeight: 200 });
  });

  it('single-section document stamps the body-level geometry on every element', () => {
    const section: SectionProps = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section, body: [para('A'), para('B')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    for (const page of pages) {
      expect(page.section.geometry.pageWidth).toBe(200);
      expect(page.section.geometry.pageHeight).toBe(140);
    }
  });
});

// A single 40pt paragraph measures exactly 40pt tall under the mock ctx (line
// height = fontPx). Same footer shape as per-section-headers-footers.test.ts's
// `footer()` — { body: [para(...)] } — but sized to 40pt so the reserve arithmetic
// is round. Verified empirically: measureHeaderFooterHeight(FOOTER_40PT) = 40.
const FOOTER_40PT: HeaderFooter = { body: [para('F', 40)] };

describe('per-section HF reserves (§17.6.11) — paginator', () => {
  it('reserves footer overflow against the PAGE section margins, not the body section', () => {
    // Mid-body section sec1 (page 0): pageHeight 200, marginTop 20, marginBottom 60,
    // footerDistance 10 ⇒ footer extent 10+40=50 FITS the 60pt bottom margin ⇒ reserve
    // 0. Its frame = 200−20−60 = 120 ⇒ five 20pt lines (100pt) fit.
    // Body section (page 1): marginBottom 20 ⇒ footer extent 50 overflows ⇒ reserve 30
    // (its OWN page). B1 fits regardless.
    // WITHOUT the fix, computeFooterReserves reads the body-level marginBottom (20) for
    // EVERY page ⇒ page 0 gets a phantom reserve 30 ⇒ usable 90 ⇒ only 4 lines ⇒ A5
    // spills to page 1. WITH the fix, page 0 reads sec1's stamped marginBottom (60) ⇒
    // reserve 0 ⇒ all five A* stay on page 0.
    const sec1: SectionGeom = {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 60, marginLeft: 20,
      headerDistance: 0, footerDistance: 10,
    };
    const bodySection: SectionProps = {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 10, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section: bodySection,
      body: [
        para('A1'), para('A2'), para('A3'), para('A4'), para('A5'),
        {
          // The mid-body section must carry the footer itself — otherwise it resolves
          // no footer on its pages and the (phantom) reserve is never applied there,
          // so the bug can't surface. Mirror doc.footers on the break.
          type: 'sectionBreak', kind: 'nextPage', geom: sec1,
          headers: { default: null, first: null, even: null },
          footers: { default: FOOTER_40PT, first: null, even: null },
          titlePage: false,
        } as BodyElement,
        para('B1'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: FOOTER_40PT, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    const textsOn = (i: number) =>
      (pages[i]?.layers.body ?? []).filter((node) => node.kind === 'paragraph').map(retainedText);
    // All five A-paragraphs fit page 0 (reserve 0 under ITS section's marginBottom 60).
    expect(textsOn(0)).toEqual(['A1', 'A2', 'A3', 'A4', 'A5']);
    expect(textsOn(1)).toEqual(['B1']);
  });
});

// A recording canvas that exposes width/height + records fillText, so a test can
// assert both the page pixel size and which paragraph text landed on a page.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[] } {
  let font = '10px serif';
  const texts: string[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string) { texts.push(s); }, strokeText(s: string) { texts.push(s); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

describe('per-section page geometry (§17.6.13/§17.6.11) — renderer', () => {
  it('sizes each page from its own section (portrait then landscape)', async () => {
    const doc = mixedDoc();
    // dpr:1 so canvas.width == cssWidth. No `width` override ⇒ each page sizes from
    // its OWN section's pageWidth × PT_TO_PX. Page 0 = portrait 200×140,
    // page 1 = landscape 140×200. The bug was every page sizing from doc.section
    // (140×200), so page 0 would come out 140-wide.
    const { canvas: c0, texts: t0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { dpr: 1 });
    const { canvas: c1, texts: t1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { dpr: 1 });

    // Join with '' — the deterministic mock ctx wraps "PORTRAIT_SECTION" into two
    // fillText calls at the content-width boundary, so the text arrives split.
    expect(t0.join('')).toContain('PORTRAIT_SECTION');
    expect(t1.join('')).toContain('LANDSCAPE_SECTION');
    // Portrait page is wider than tall; landscape page is taller than wide. This is
    // the load-bearing discriminator: before the fix page 0 sized from the body-level
    // (landscape 140×200) section, so c0 would be TALLER than wide (test Red).
    expect(c0.width).toBeGreaterThan(c0.height);
    expect(c1.height).toBeGreaterThan(c1.width);
    // Aspect ratios match the two sections (200/140 vs 140/200). Precision 2: canvas
    // dims are integer-pixel (Math.round(cssW/H*dpr)), so the ratio carries a ~7e-4
    // rounding residual — precision 2 still cleanly separates 1.43 from the buggy 0.70.
    expect(c0.width / c0.height).toBeCloseTo(200 / 140, 2);
    expect(c1.width / c1.height).toBeCloseTo(140 / 200, 2);
  });

  it('honours opts.width per CALL: same pixel width, per-page height from aspect', async () => {
    // The scroll viewer's contract: `width` is the canvas CSS width for THIS call.
    // Mixed-width pages rendered at a constant width fill the same pixels at
    // different px-per-pt scales; height still follows each page's OWN aspect.
    // Page 0 (portrait 200×140): height = 400·140/200 = 280.
    // Page 1 (landscape 140×200): height = 400·200/140 ≈ 571.
    const doc = mixedDoc();
    const { canvas: c0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { width: 400, dpr: 1 });
    const { canvas: c1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { width: 400, dpr: 1 });
    expect(c0.width).toBe(400);
    expect(c1.width).toBe(400);
    expect(c0.height).toBe(280);
    expect(c1.height).toBe(Math.round(400 * (200 / 140)));
  });

  it('single-section document sizes every page from doc.section (no regression)', async () => {
    const section: SectionProps = {
      pageWidth: 300, pageHeight: 400,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section, body: [para('ONLY_SECTION')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const { canvas } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1 });
    // Precision 2: integer-pixel rounding gives 400/533 = 0.7505, ~5e-4 off 0.75.
    expect(canvas.width / canvas.height).toBeCloseTo(300 / 400, 2);
  });
});

describe('page-size fact (§17.6.13/§17.6.11) — canonical page geometry', () => {
  it('exposes per-page width/height directly on each retained page', () => {
    const doc = mixedDoc();
    const pages = canonicalLayout(doc).pages;
    const sizeOf = (i: number) => {
      const geometry = pages[i]?.geometry;
      return { widthPt: geometry?.widthPt ?? doc.section.pageWidth, heightPt: geometry?.heightPt ?? doc.section.pageHeight };
    };
    expect(sizeOf(0)).toEqual({ widthPt: 200, heightPt: 140 });
    expect(sizeOf(1)).toEqual({ widthPt: 140, heightPt: 200 });
  });

  // Direct unit coverage of DocxDocument.pageSize itself. The class needs a
  // Worker/WASM to construct normally, so inject the private state onto a bare
  // prototype instance (TS-private fields are runtime-assignable) — covering the
  // worker branch (_meta), the main branch (_document + pagination), clamping,
  // the not-loaded {0,0} contract, and the fresh-object (no aliasing) guarantee.
  it('DocxDocument.pageSize: both branches, clamp, not-loaded, no aliasing', async () => {
    const { DocxDocument } = await import('./document.js');
    // The constructor is private, so derive the instance type from the factory.
    type Doc = Awaited<ReturnType<typeof DocxDocument.load>>;
    const make = (state: Record<string, unknown>) => {
      const d = Object.create(DocxDocument.prototype) as Doc;
      attachDocumentLayoutRuntime(d, 0);
      Object.assign(d, { _meta: null, _document: null, _pages: null, _mode: 'main' }, state);
      const model = (d as unknown as { _document: DocxDocumentModel | null })._document;
      if (model) {
        const runtime = documentLayoutRuntimeOf(d);
        const services = createLayoutServices(model);
        runtime.services = services;
        attachDocumentLayoutVariants({
          source: layoutSourceStore(model),
          services,
          defaultCurrentDateMs: runtime.defaultCurrentDateMs,
          buildLayout: (options) => layoutDocument(model, services, options),
        });
      }
      return d;
    };

    // Worker branch: reads _meta.pageSizes with clamp; returns a COPY.
    const meta = { pageCount: 2, revisions: [], comments: [], footnotes: [], endnotes: [],
      pageSizes: [{ widthPt: 200, heightPt: 140 }, { widthPt: 140, heightPt: 200 }] };
    const w = make({ _meta: meta });
    expect(w.pageSize(0)).toEqual({ widthPt: 200, heightPt: 140 });
    expect(w.pageSize(1)).toEqual({ widthPt: 140, heightPt: 200 });
    expect(w.pageSize(-5)).toEqual({ widthPt: 200, heightPt: 140 }); // clamp low
    expect(w.pageSize(99)).toEqual({ widthPt: 140, heightPt: 200 }); // clamp high
    const returned = w.pageSize(0);
    returned.widthPt = 9999; // mutating the return value must not corrupt the meta
    expect(w.pageSize(0)).toEqual({ widthPt: 200, heightPt: 140 });

    // Main branch: paginates the model and reads the stamped sectionGeom.
    const m = make({ _document: mixedDoc() });
    expect(m.pageSize(0)).toEqual({ widthPt: 200, heightPt: 140 });
    expect(m.pageSize(1)).toEqual({ widthPt: 140, heightPt: 200 });
    expect(m.pageSize(99)).toEqual({ widthPt: 140, heightPt: 200 }); // clamp high

    // Not loaded (pre-load / post-destroy): {0,0}.
    expect(make({}).pageSize(0)).toEqual({ widthPt: 0, heightPt: 0 });

    // Engine fact (WS4 O1): the render mode is publicly readable so a viewer
    // receiving an injected engine can route main vs bitmap paths without probing.
    const workerish = make({ _mode: 'worker' });
    expect(workerish.mode).toBe('worker');
    const mainish = make({ _mode: 'main' });
    expect(mainish.mode).toBe('main');
  });
});

// ECMA-376 §17.6.11 (pgMar) + §17.6.13 (pgSz) — a newspaper-column band's WIDTH
// and x-origin derive from the OWNING section's page geometry, not the body-level
// section. `computeColumns` is fed by `sectionColumnsFrom`; the paginator sizes a
// table to the active column band (`colW()`), so the table fragment's stored column
// widths are the observable proxy for the band width the fix corrects.
function cellPara(text: string): DocParagraph {
  return {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false,
      background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
}

// A single-row, single-column FIXED-layout table whose `<w:tblGrid>` width is
// `gridWidthPt`. Fixed layout uses the grid verbatim and only scales DOWN when the
// grid overflows the available band (resolveColumnWidths §17.4.52), so the fragment's
// `columnWidthsPt` equals `gridWidthPt` iff the band is ≥ that width.
function fixedTable(gridWidthPt: number): DocTable {
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...cellPara('x') }],
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: gridWidthPt,
  } as unknown as DocTableCell;
  const row: DocTableRow = {
    cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
  } as unknown as DocTableRow;
  return {
    colWidths: [gridWidthPt], rows: [row],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', widthPt: gridWidthPt,
  } as unknown as DocTable;
}

describe('per-section newspaper-column band (§17.6.11/§17.6.13) — paginator', () => {
  it('sizes a table to the OWNING section band width, not the body-level band', () => {
    // First section (ended by the nextPage break): page 600 wide, margins L=10 R=40
    // ⇒ content band 600−10−40 = 550. Its single fixed table has a tblGrid of 550,
    // so it fits EXACTLY and its columns stay 550.
    const sec1: SectionGeom = {
      pageWidth: 600, pageHeight: 500,
      marginTop: 20, marginRight: 40, marginBottom: 20, marginLeft: 10,
      headerDistance: 0, footerDistance: 0,
    };
    // Body (final) section: page 600 wide, margins L=75 R=75 ⇒ content band 450.
    // BEFORE the fix, `sectionColumnsFrom` builds the first section's band from THIS
    // body-level width (450), so the 550-wide grid is clamped to 450.
    const bodySection: SectionProps = {
      pageWidth: 600, pageHeight: 500,
      marginTop: 20, marginRight: 75, marginBottom: 20, marginLeft: 75,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const body: BodyElement[] = [
      { type: 'table', ...fixedTable(550) } as BodyElement,
      { type: 'sectionBreak', kind: 'nextPage', geom: sec1 } as BodyElement,
      para('BODY_SECTION'),
    ];
    const doc = {
      section: bodySection, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const table = canonicalLayout(doc).pages[0].layers.body.find((node) => node.kind === 'table');
    expect(table?.kind).toBe('table');
    const resolvedTotal = table?.kind === 'table'
      ? table.columnWidthsPt.reduce((s, w) => s + w, 0)
      : 0;
    // The table belongs to the first section (band 550), so its columns sum to 550 —
    // NOT the body-level band 450 the bug clamps them to.
    expect(resolvedTotal).toBeCloseTo(550, 5);
  });
});
