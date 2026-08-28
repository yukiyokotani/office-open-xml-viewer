/**
 * Deterministic synthetic DOCX models for layout performance work.
 *
 * `paginateBody` is not one pass: it is a seed pass wrapped in up to four
 * whole-document fixed-point solvers (continuous-section column balancing,
 * page-owned anchor convergence, header/footer reserve convergence, and
 * PAGE/NUMPAGES field feedback). A single "big document" fixture cannot tell
 * which of those multipliers a change actually moved, so the shapes below
 * isolate them:
 *
 * - `plain`          — no headers/footers, no fields. Exercises the seed pass
 *                      alone, so it measures raw block-loop + measurement cost.
 * - `header-footer`  — field-free header and footer stories. Adds one reserve
 *                      measurement over every page plus the confirming
 *                      convergence iteration.
 * - `fields`         — a footer carrying PAGE and NUMPAGES. Turns on the
 *                      pagination-field feedback edge
 *                      (`paginatedFlowHasPaginationDependentFields`), which
 *                      re-paginates the whole document until the page contexts
 *                      reach a fixed point.
 * - `tables`         — auto-fit tables, whose intrinsic-width measurement runs
 *                      an extra line-breaking pass per cell paragraph.
 * - `long-paragraphs`— few, very long single-script paragraphs, where cluster
 *                      geometry's per-grapheme prefix measurement dominates.
 * - `tracked`        — a mix of ordinary and deleted (`w:del`) runs. The two
 *                      `showTrackedChanges` views paginate DIFFERENTLY: the
 *                      final view hides deletions, the markup view keeps them
 *                      visible, so this is the fixture for anything that must
 *                      distinguish layout variants.
 * - `tracked-fields` — the same, plus the PAGE/NUMPAGES footer, i.e. a document
 *                      that is both variant-sensitive and inexact under a
 *                      prefix preview (the shape of a real reviewed contract).
 *
 * Everything here is derived from a fixed seed: the same `shape` and
 * `paragraphs` always produce byte-identical models, so a benchmark or an
 * equivalence test can compare runs (and machines) meaningfully. Test-only —
 * never imported by production code.
 */
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocxDocumentModel,
  HeaderFooter,
  SectionProps,
} from '../types.js';

export type SyntheticDocumentShape =
  | 'plain'
  | 'header-footer'
  | 'fields'
  | 'tables'
  | 'long-paragraphs'
  | 'tracked'
  | 'tracked-fields';

export interface SyntheticDocumentOptions {
  /** 0-based indices of body paragraphs to mark `w:keepNext` (§17.3.1.15).
   *  A run of consecutive indices forms one keep-set whose terminal block is
   *  the first unmarked paragraph after it — the shape that exercises the
   *  paginator's unbounded successor scan across a progressive truncation
   *  cut. Applied to paragraph-bodied shapes (`plain`, `header-footer`,
   *  `fields`); ignored for `tables`. */
  readonly keepNextIndices?: readonly number[];
  /** Body paragraph count (or table count for `tables`). */
  readonly paragraphs?: number;
  /** Words per generated paragraph. `long-paragraphs` overrides this upward. */
  readonly wordsPerParagraph?: number;
}

/**
 * Deterministic 32-bit LCG (Numerical Recipes constants). `Math.random` is
 * deliberately avoided: a benchmark that re-shapes its own input between runs
 * measures nothing.
 */
function lcg(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 0x100000000;
  };
}

/** A closed vocabulary keeps the shaping/measurement cache hit-rate realistic:
 *  real prose reuses words heavily, and a purely random alphabet would make the
 *  measurement caches look far less effective than they are on real documents. */
const WORDS = [
  'document', 'layout', 'paragraph', 'section', 'measure', 'render', 'page',
  'column', 'anchor', 'footnote', 'heading', 'table', 'border', 'spacing',
  'baseline', 'justify', 'kerning', 'glyph', 'canvas', 'viewer', 'the', 'of',
  'and', 'to', 'a', 'in', 'that', 'is', 'for', 'with', 'as', 'by', 'on',
];

function sentence(next: () => number, words: number): string {
  const parts: string[] = [];
  for (let index = 0; index < words; index += 1) {
    parts.push(WORDS[Math.floor(next() * WORDS.length)]);
  }
  const text = parts.join(' ');
  return `${text.charAt(0).toUpperCase()}${text.slice(1)}.`;
}

function textRun(text: string, fontSize = 10): DocParagraph['runs'][number] {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize,
    color: null,
    fontFamily: 'Times New Roman',
    fontFamilyEastAsia: '',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  } as DocParagraph['runs'][number];
}

/** A PAGE / NUMPAGES field run — the two dependencies
 *  {@link paginationFieldDependency} recognizes as pagination-feedback edges. */
function fieldRun(
  fieldType: 'page' | 'numPages',
  fallbackText: string,
): DocParagraph['runs'][number] {
  return {
    type: 'field',
    fieldType,
    instruction: fieldType === 'page' ? 'PAGE' : 'NUMPAGES',
    fallbackText,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Times New Roman',
    background: null,
    vertAlign: null,
  } as DocParagraph['runs'][number];
}

function paragraph(
  runs: readonly DocParagraph['runs'][number][],
  over: Partial<DocParagraph> = {},
): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 6,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [...runs],
    defaultFontSize: 10,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function textParagraph(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return paragraph([textRun(text)], over);
}

/** A run marked as deleted (ECMA-376 §17.13.5.14 `w:del`). Hidden by the
 *  final-view default; shown struck through in the markup view. */
function deletedRun(text: string): DocParagraph['runs'][number] {
  return {
    ...(textRun(text) as Record<string, unknown>),
    revision: { kind: 'deletion', author: 'Reviewer A', date: '2024-01-01T00:00:00Z' },
  } as DocParagraph['runs'][number];
}

/**
 * A paragraph whose length depends on which tracked-change view is selected.
 *
 * Roughly half the text sits in a deleted run, so the final view lays out
 * materially less content than the markup view and the two variants reach
 * different page counts — which is the whole point of the fixture. Every other
 * paragraph is left untouched so the difference accumulates gradually rather
 * than doubling the document.
 */
function trackedParagraph(next: () => number, words: number, index: number): DocParagraph {
  if (index % 2 === 1) return textParagraph(sentence(next, words));
  return paragraph([
    textRun(`${sentence(next, Math.max(1, Math.round(words / 2)))} `),
    deletedRun(sentence(next, Math.max(1, Math.round(words / 2)))),
  ]);
}

/** All-unset edges. Acquisition dereferences the edge record itself, so an
 *  absent border set is `{ top: null, ... }` rather than a bare `null`. */
const NO_BORDERS = {
  top: null, bottom: null, left: null, right: null, insideH: null, insideV: null,
};

/**
 * A 3x3 table left to ECMA-376 §17.4.52 `autofit` layout, so cell content — not
 * an authored `tblW` — drives the column widths and the intrinsic-width
 * measurement pass runs for every cell paragraph.
 */
function autoFitTable(next: () => number, wordsPerCell: number): BodyElement {
  const rows = [0, 1, 2].map(() => ({
    cells: [0, 1, 2].map(() => ({
      content: [textParagraph(sentence(next, wordsPerCell))] as CellElement[],
      colSpan: 1,
      vMerge: null,
      borders: { ...NO_BORDERS },
      background: null,
      vAlign: 'top' as const,
      widthPt: null,
    })),
    rowHeight: null,
    rowHeightRule: 'auto' as const,
    isHeader: false,
  }));
  return {
    type: 'table',
    colWidths: [],
    rows,
    borders: { ...NO_BORDERS },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 5.4,
    cellMarginRight: 5.4,
    jc: 'left',
    layout: 'autofit',
  } as unknown as BodyElement;
}

function headerFooterStory(runs: readonly DocParagraph['runs'][number][]): HeaderFooter {
  return { body: [paragraph(runs, { spaceAfter: 0 }) as BodyElement] };
}

function section(): SectionProps {
  return {
    pageWidth: 612,
    pageHeight: 792,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
    headerDistance: 36,
    footerDistance: 36,
    titlePage: false,
    evenAndOddHeaders: false,
    sectionStart: 'nextPage',
    columns: null,
  } as SectionProps;
}

/**
 * Build one synthetic model. The returned object is a plain
 * {@link DocxDocumentModel} — the same shape the WASM parser materializes — so
 * it drives `layoutDocument` / `paginateBody` exactly as a real file would,
 * without needing WASM artifacts or a real canvas.
 */
export function syntheticDocxModel(
  shape: SyntheticDocumentShape,
  options: SyntheticDocumentOptions = {},
): DocxDocumentModel {
  const next = lcg(0x5eed);
  const count = options.paragraphs ?? 200;
  const words = options.wordsPerParagraph ?? 40;

  let body: BodyElement[];
  if (shape === 'tables') {
    body = Array.from({ length: count }, () => autoFitTable(next, 8));
  } else if (shape === 'long-paragraphs') {
    body = Array.from(
      { length: count },
      () => textParagraph(sentence(next, options.wordsPerParagraph ?? 1200)) as BodyElement,
    );
  } else if (shape === 'tracked' || shape === 'tracked-fields') {
    body = Array.from(
      { length: count },
      (_unused, index) => trackedParagraph(next, words, index) as BodyElement,
    );
  } else {
    const keepNext = new Set(options.keepNextIndices ?? []);
    body = Array.from(
      { length: count },
      (_unused, index) => textParagraph(
        sentence(next, words),
        keepNext.has(index) ? { keepNext: true } : {},
      ) as BodyElement,
    );
  }

  const headers: DocxDocumentModel['headers'] = { default: null, first: null, even: null };
  const footers: DocxDocumentModel['footers'] = { default: null, first: null, even: null };
  if (shape === 'header-footer') {
    headers.default = headerFooterStory([textRun('Synthetic layout benchmark')]);
    footers.default = headerFooterStory([textRun('Confidential draft')]);
  } else if (shape === 'fields' || shape === 'tracked-fields') {
    headers.default = headerFooterStory([textRun('Synthetic layout benchmark')]);
    footers.default = headerFooterStory([
      textRun('Page '),
      fieldRun('page', '1'),
      textRun(' of '),
      fieldRun('numPages', '1'),
    ]);
  }

  return {
    section: section(),
    body,
    headers,
    footers,
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

/**
 * Install the linear-metric `OffscreenCanvas` stub the layout suites use
 * (glyph advance = fontPx × 0.5). Layout measurement goes through Canvas 2D
 * `measureText`, so a deterministic stub is what makes layout output —  and
 * therefore `layoutFingerprint` equivalence — reproducible in Node.
 *
 * Returns a counter of `measureText` calls, which the benchmark reports: the
 * call count is the machine-independent half of the measurement cost, so it
 * stays comparable across runs where wall-clock is noisy.
 */
export function installStubCanvas(): { measureTextCalls: () => number } {
  let calls = 0;
  const makeContext = (): CanvasRenderingContext2D => {
    let font = '10px serif';
    const context = {
      get font() { return font; },
      set font(value: string) { font = value; },
      letterSpacing: '0px',
      fontKerning: 'normal',
      measureText: (text: string) => {
        calls += 1;
        const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
        const per = px * 0.5;
        return {
          width: [...text].length * per,
          fontBoundingBoxAscent: px * 0.8,
          fontBoundingBoxDescent: px * 0.2,
          actualBoundingBoxAscent: px * 0.8,
          actualBoundingBoxDescent: px * 0.2,
        } as TextMetrics;
      },
      save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
      stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
      scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
      quadraticCurveTo() {}, bezierCurveTo() {},
      createLinearGradient() { return { addColorStop() {} }; },
      drawImage() {}, fillText() {}, strokeText() {},
      fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'left' as CanvasTextAlign,
      direction: 'ltr' as CanvasDirection,
    };
    return context as unknown as CanvasRenderingContext2D;
  };
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeContext(); }
  };
  return { measureTextCalls: () => calls };
}
