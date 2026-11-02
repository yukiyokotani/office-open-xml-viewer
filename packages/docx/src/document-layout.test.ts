import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// PR 5 Task 12 — body layout fragments.
//
// `layoutDocument(doc)` produces an immutable `DocumentLayout`: pages of
// `PlacedFragment`s wrapping `ParagraphFragment`s. Each fragment references its
// SOURCE paragraph (never a mutated copy), a placement-aware `MeasuredParagraph`,
// and a `[lineStart, lineEnd)` line range. This suite pins the fragment model
// contract (design doc §"Measured Fragment Model" / §"Pagination and paint
// invariants"): source identity, immutable line ranges, placement coordinates,
// page geometry, section context, paragraph continuation across pages, and the
// spacing-ownership invariant
//   cursor advancement == leadingSpacePt + measured line advances + trailingSpacePt.
// ─────────────────────────────────────────────────────────────────────────────

/** OffscreenCanvas polyfill with a linear glyph metric (width = fontPx * 0.5),
 *  matching the other renderer suites so `layoutDocument`'s scale-1 measurement is
 *  deterministic in node. */
function makeStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
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

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeStubCtx(); }
  };
});

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

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
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

/** Every paragraph fragment on every page, in document order. */
function allFragments(model: DocxDocumentModel) {
  const layout = layoutDocument(model);
  return layout.pages.flatMap((page) =>
    page.layers.body
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .map((fragment) => ({ page, fragment })),
  );
}

function sizedTextRun(
  text: string,
  fontSize: number,
  ruby?: Readonly<{ text: string; fontSizePt: number }>,
): DocParagraph['runs'][number] {
  return {
    type: 'text', text, bold: false, italic: false, underline: false,
    strikethrough: false, fontSize, color: null, fontFamily: 'Times New Roman',
    fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
    hyperlink: null, ...(ruby ? { ruby } : {}),
  } as DocParagraph['runs'][number];
}

function cjkParagraph(
  text: string,
  fontSize: number,
  over: Partial<DocParagraph> = {},
): DocParagraph {
  return para('', {
    runs: [sizedTextRun(text, fontSize)],
    defaultFontSize: fontSize,
    ...over,
  });
}

function columnDocument(
  body: BodyElement[],
  widthsPt: readonly number[],
  spacesPt: readonly number[],
  pageHeight: number,
): DocxDocumentModel {
  const model = doc(body, pageHeight);
  const contentWidthPt = widthsPt.reduce((sum, width) => sum + width, 0)
    + spacesPt.reduce((sum, space) => sum + space, 0);
  (model.section as SectionProps).pageWidth = contentWidthPt + 20;
  (model.section as SectionProps).columns = {
    count: widthsPt.length,
    spacePt: 0,
    equalWidth: false,
    sep: false,
    cols: widthsPt.map((widthPt, index) => ({
      widthPt,
      spacePt: spacesPt[index] ?? 0,
    })),
  };
  return model;
}

function sourceFragments(layout: ReturnType<typeof layoutDocument>, bodyIndex: number) {
  return layout.pages.flatMap((page) => page.layers.body.filter(
    (node): node is ParagraphLayout => node.kind === 'paragraph'
      && node.source.story === 'body'
      && node.source.storyInstance === 'body'
      && node.source.path[0] === bodyIndex,
  ));
}

function fragmentText(fragments: readonly ParagraphLayout[]): string {
  return fragments.flatMap((fragment) => fragment.lines)
    .flatMap((line) => line.placements)
    .filter((placement) => placement.kind === 'text')
    .map((placement) => placement.text)
    .join('');
}

function lineInlineAdvance(line: ParagraphLayout['lines'][number]): number {
  return line.placements.reduce(
    (sum, placement) => sum + ('advancePt' in placement ? placement.advancePt : 0),
    0,
  );
}

function fragmentLineAdvancesPt(fragment: ParagraphLayout): number {
  return fragment.lines.reduce((sum, line) => sum + line.advancePt, 0);
}

describe('layoutDocument — body paragraph fragments (PR 5 Task 12)', () => {
  it('assigns positional source paths and deterministic layout ids to repeated object references', () => {
    const repeated = para('shared object');
    const model = doc([
      repeated as unknown as BodyElement,
      repeated as unknown as BodyElement,
    ]);

    const first = layoutDocument(model).pages.flatMap((page) => page.layers.body);
    const second = layoutDocument(model).pages.flatMap((page) => page.layers.body);

    expect(first.map((node) => node.source.path)).toEqual([[0], [1]]);
    expect(new Set(first.map((node) => node.id)).size).toBe(2);
    expect(second.map((node) => node.id)).toEqual(first.map((node) => node.id));
  });

  it('emits one self-contained retained paragraph per body paragraph', () => {
    const p1 = para('alpha');
    const p2 = para('beta');
    const model = doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]);
    const layout = layoutDocument(model);

    expect(layout.pages.length).toBe(1);
    const frags = layout.pages[0].layers.body;
    expect(frags.length).toBe(2);
    expect(frags[0].kind).toBe('paragraph');
    expect(frags[1].kind).toBe('paragraph');
    if (frags[0].kind !== 'paragraph' || frags[1].kind !== 'paragraph') {
      throw new Error('expected paragraph layouts');
    }
    expect(frags[0]).not.toHaveProperty('measured');
    expect(frags[0]).not.toHaveProperty('runs');
    expect(frags[0].source.path).toEqual([0]);
    expect(frags[1].source.path).toEqual([1]);
  });

  it('keeps fragment state OFF the source paragraph (no fragment fields added to DocParagraph)', () => {
    // Historical paginator runtime stamps (column index,
    // colGeom, sectionGeom, ...) are pre-existing and out of scope; this pins that
    // the NEW fragment model never writes its fields onto the parsed paragraph — the
    // measurement/line-range/spacing live in the layout result, keyed off-object.
    const p1 = para('gamma');
    layoutDocument(doc([p1 as unknown as BodyElement]));
    const record = p1 as unknown as Record<string, unknown>;
    for (const field of ['measured', 'lineStart', 'lineEnd', 'leadingSpacePt', 'trailingSpacePt', 'fragment', 'placedFragment']) {
      expect(record[field]).toBeUndefined();
    }
  });

  it('records immutable line ranges covering the whole paragraph', () => {
    const p = para(Array.from({ length: 40 }, () => 'w').join(' '));
    const layout = layoutDocument(doc([p as unknown as BodyElement]));
    const frag = layout.pages[0].layers.body[0] as ParagraphLayout;
    expect(frag.lines[0]?.range.start).toBe(0);
    expect(frag.lines.at(-1)?.range.end).toBeGreaterThan(0);
    expect(frag.lines.length).toBeGreaterThan(1); // actually wrapped
  });

  it('places fragments with page-absolute coordinates and the content-band width', () => {
    const p1 = para('one');
    const p2 = para('two');
    const layout = layoutDocument(doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]));
    const [f1, f2] = layout.pages[0].layers.body;
    // contentX = marginLeft, width = pageWidth - marginLeft - marginRight.
    expect(f1.flowBounds.xPt).toBeCloseTo(10, 6);
    expect(f1.flowBounds.widthPt).toBeCloseTo(180, 6);
    // The first fragment starts at the top content inset; the second stacks below it.
    expect(f1.flowBounds.yPt).toBeCloseTo(10, 6);
    expect(f2.flowBounds.yPt).toBeGreaterThan(f1.flowBounds.yPt);
    expect(f2.flowBounds.yPt).toBeCloseTo(f1.flowBounds.yPt + f1.advancePt, 6);
  });

  it('exposes page geometry and the resolved section context', () => {
    const layout = layoutDocument(doc([para('x') as unknown as BodyElement]));
    const page = layout.pages[0];
    expect(page.pageIndex).toBe(0);
    expect(page.geometry.widthPt).toBe(200);
    expect(page.geometry.heightPt).toBe(400);
    expect(page.section.geometry.marginLeft).toBe(10);
    // SectionLayoutContext carries the resolved grid policy (docGrid absent => none).
    expect(page.section.grid.kind).toBe('none');
    // M-2 — a single-column section exposes one column spanning the content band.
    expect(page.section.columns.length).toBe(1);
    expect(page.section.columns[0].wPt).toBeCloseTo(180, 6);
  });

  it('M-2: exposes the §17.6.4 per-page column geometry the paginator resolved', () => {
    const model = doc([para('x') as unknown as BodyElement, para('y') as unknown as BodyElement]);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [],
    } as SectionProps['columns'];
    const page = layoutDocument(model).pages[0];
    // Two newspaper columns: content band 180pt minus 20pt gutter → 80pt each.
    expect(page.section.columns.length).toBe(2);
    expect(page.section.columns[0].wPt).toBeCloseTo(80, 6);
    expect(page.section.columns[1].wPt).toBeCloseTo(80, 6);
    // A fragment records which column it is placed in.
    const domains = new Set(page.sectionRegions[0]?.flowDomainIds ?? []);
    for (const node of page.layers.body) expect(domains.has(node.flowDomainId)).toBe(true);
  });

  it('splits a long paragraph into continuation fragments over one source', () => {
    // Force several pages: short page height, a long wrapping paragraph.
    const p = para(Array.from({ length: 300 }, () => 'w').join(' '));
    const layout = layoutDocument(doc([p as unknown as BodyElement], 60));
    const frags = allFragments(doc([p as unknown as BodyElement], 60));

    expect(layout.pages.length).toBeGreaterThan(1);
    for (const f of frags) expect(f.fragment.source).toEqual(frags[0].fragment.source);
    for (let index = 1; index < frags.length; index += 1) {
      expect(frags[index].fragment.continuation?.continuesFromPrevious).toBe(true);
      expect(frags[index - 1].fragment.lines.at(-1)!.range.end)
        .toBe(frags[index].fragment.lines[0]!.range.start);
    }

    // Leading spacing only on the first fragment; trailing only on the last.
    expect(frags[frags.length - 1].fragment.spacing.afterPt).toBeGreaterThanOrEqual(0);
    for (let i = 1; i < frags.length; i++) {
      expect(frags[i].fragment.spacing.beforePt).toBe(0);
    }
    for (let i = 0; i < frags.length - 1; i++) {
      expect(frags[i].fragment.spacing.afterPt).toBe(0);
    }
  });

  it('remeasures at a changed wrap width: a fragment measurement matches its own placement', () => {
    const p = para(Array.from({ length: 40 }, () => 'w').join(' '));
    const wide = layoutDocument(doc([p as unknown as BodyElement])); // width 180
    const narrowModel = doc([p as unknown as BodyElement]);
    (narrowModel.section as SectionProps).pageWidth = 120; // width 100
    const narrow = layoutDocument(narrowModel);

    const wf = wide.pages[0].layers.body[0] as ParagraphLayout;
    const nf = narrow.pages[0].layers.body[0] as ParagraphLayout;
    // Each fragment's measurement reflects its own available width (no stale reuse).
    expect(wf.flowBounds.widthPt).toBeCloseTo(180, 6);
    expect(nf.flowBounds.widthPt).toBeCloseTo(100, 6);
    // A narrower band wraps to more lines.
    expect(nf.lines.length).toBeGreaterThan(wf.lines.length);
  });

  it('INVARIANT: cursor advancement == leadingSpacePt + line advances + trailingSpacePt', () => {
    const p1 = para(Array.from({ length: 30 }, () => 'w').join(' '), { spaceBefore: 6, spaceAfter: 8 });
    const p2 = para('tail', { spaceBefore: 4, spaceAfter: 3 });
    const frags = allFragments(doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]));
    for (const { fragment } of frags) {
      expect(fragment.advancePt).toBeCloseTo(
        fragment.spacing.beforePt
          + fragmentLineAdvancesPt(fragment)
          + fragment.spacing.afterPt,
        6,
      );
    }
  });

  it('freezes the layout result and its fragment arrays', () => {
    const layout = layoutDocument(doc([para('x') as unknown as BodyElement]));
    expect(Object.isFrozen(layout)).toBe(true);
    expect(Object.isFrozen(layout.pages)).toBe(true);
    expect(Object.isFrozen(layout.pages[0])).toBe(true);
    expect(Object.isFrozen(layout.pages[0].layers.body)).toBe(true);
    expect(Object.isFrozen(layout.pages[0].layers.body[0])).toBe(true);
    // The inner ParagraphFragment is frozen too — its measured lines, line range and
    // spacing are immutable, so paint can never mutate the layout result (design
    // §"Pagination and paint invariants" 4).
    expect(Object.isFrozen(layout.pages[0].layers.body[0])).toBe(true);
  });
});

describe('layoutDocument — exact source-boundary paragraph continuation', () => {
  it('moves a break-only grid heading intact when its first full line does not fit', () => {
    // Windows Word 16.0.20430: both ends of the break-only heading land on
    // page 2 even with keepNext/keepLines/widowControl explicitly disabled.
    // Unlike visible glyph lines, blank break lines get no grid-leading relief.
    const target = para('', {
      runs: [{ type: 'break', breakType: 'line' }],
      defaultFontSize: 16,
      lineSpacing: { rule: 'auto', value: 1.15, explicit: true },
      keepNext: false, keepLines: false, widowControl: false,
      spaceAfter: 2,
    });
    const model = doc([
      para('F', { lineSpacing: { rule: 'exact', value: 60 } }) as BodyElement,
      target as BodyElement,
      para('After', { lineSpacing: { rule: 'exact', value: 10 } }) as BodyElement,
    ], 100);
    model.section.docGridType = 'lines';
    model.section.docGridLinePitch = 18;
    const layout = layoutDocument(model);
    const fragments = sourceFragments(layout, 1);

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0]!.layers.body.map(node => node.source.path[0])).toEqual([0]);
    expect(layout.pages[1]!.layers.body.map(node => node.source.path[0])).toEqual([1, 2]);
    expect(fragments).toHaveLength(1);
    expect(fragments[0]!.lines).toHaveLength(2);
    expect(fragments[0]!.lines.map(line => line.advancePt)).toEqual([20.7, 20.7]);
    expect(fragments[0]!.spacing.afterPt).toBe(2);
    expect(fragmentText(sourceFragments(layout, 0))).toBe('F');
    expect(fragmentText(sourceFragments(layout, 2))).toBe('After');
    expect(sourceFragments(layout, 2)[0]!.flowBounds.yPt).toBeCloseTo(53.4);
  });

  it.each(['', 'A'])('retains the empty line after a paginated trailing break (%j)', (text) => {
    const target = para('', {
      runs: [...(text ? [sizedTextRun(text, 10)] : []), { type: 'break', breakType: 'line' }],
      lineSpacing: { rule: 'exact', value: 20 },
      spaceAfter: 3,
    });
    const model = doc([
      para('F', { lineSpacing: { rule: 'exact', value: 60 } }) as BodyElement,
      target as BodyElement,
      para('After', { lineSpacing: { rule: 'exact', value: 10 } }) as BodyElement,
    ], 100);
    const layout = layoutDocument(model);
    const fragments = sourceFragments(layout, 1);

    expect(layout.pages).toHaveLength(2);
    expect(fragments.map(fragment => fragment.lines.length)).toEqual([1, 1]);
    expect(fragments.map(fragment => fragment.lines[0]!.advancePt)).toEqual([20, 20]);
    expect(fragments.map(fragment => fragment.lines[0]!.bounds.yPt)).toEqual([70, 10]);
    expect(fragments.map(fragment => fragment.spacing.afterPt)).toEqual([0, 3]);
    expect(fragmentText(fragments)).toBe(text);
    const after = sourceFragments(layout, 2);
    expect(fragmentText(after)).toBe('After');
    expect(after[0]!.lines[0]!.bounds.yPt).toBe(33);
  });

  it('continues exactly from a 100pt column into a 48pt destination measurement', () => {
    const filler = cjkParagraph('う'.repeat(30), 20);
    const text = 'あ'.repeat(20);
    const target = cjkParagraph(text, 20, { indentFirst: 20, spaceBefore: 20 });
    const layout = layoutDocument(columnDocument(
      [filler as unknown as BodyElement, target as unknown as BodyElement],
      [100, 48], [12, 0], 120,
    ));
    const fragments = sourceFragments(layout, 1);

    expect(fragments.map((fragment) => fragment.flowBounds.widthPt)).toEqual([100, 48]);
    expect(fragmentText(fragments)).toBe(text);
    for (let index = 1; index < fragments.length; index += 1) {
      expect(fragments[index - 1]!.lines.at(-1)!.range.end)
        .toBe(fragments[index]!.lines[0]!.range.start);
    }
    for (const fragment of fragments.slice(0, -1)) {
      expect(fragment.lines.at(-1)!.range.end).toBeGreaterThan(fragment.lines[0]!.range.start);
    }
    const narrow = fragments[1]!;
    expect(narrow.lines.every((line) => lineInlineAdvance(line) <= 48)).toBe(true);
    expect(narrow.spacing.beforePt).toBe(0);
    expect(new Set(narrow.lines.map((line) => line.bounds.xPt)).size).toBe(1);
    expect(fragments[0]!.lines[0]!.bounds.xPt).toBeGreaterThan(fragments[0]!.flowBounds.xPt);
  });

  it('composes original-segment boundaries across 100pt, 48pt, and 30pt columns', () => {
    const filler = cjkParagraph('う'.repeat(40), 20);
    const text = 'あ'.repeat(35);
    const target = cjkParagraph(text, 20);
    const layout = layoutDocument(columnDocument(
      [filler as unknown as BodyElement, target as unknown as BodyElement],
      [100, 48, 30], [6, 6, 0], 120,
    ));
    const fragments = sourceFragments(layout, 1);

    expect(fragments.map((fragment) => fragment.flowBounds.widthPt)).toEqual([100, 48, 30]);
    expect(fragmentText(fragments)).toBe(text);
    expect(fragments.map((fragment) => ({
      start: fragment.lines[0]!.range.start,
      end: fragment.lines.at(-1)!.range.end,
    }))).toEqual([
      { start: 0, end: 10 },
      { start: 10, end: 30 },
      { start: 30, end: 35 },
    ]);
    expect(fragments[1]!.lines.every((line) => lineInlineAdvance(line) <= 48)).toBe(true);
    expect(fragments[2]!.lines.every((line) => lineInlineAdvance(line) <= 30)).toBe(true);
  });

  it('keeps a same-width suffix immediately after the accepted source boundary', () => {
    const filler = cjkParagraph('う'.repeat(40), 20);
    const text = 'あ'.repeat(15);
    const target = cjkParagraph(text, 20);
    const layout = layoutDocument(columnDocument(
      [filler as unknown as BodyElement, target as unknown as BodyElement],
      [100, 100], [10, 0], 120,
    ));
    const fragments = sourceFragments(layout, 1);

    expect(fragmentText(fragments)).toBe(text);
    expect(fragments.map((fragment) => ({
      start: fragment.lines[0]!.range.start,
      end: fragment.lines.at(-1)!.range.end,
    }))).toEqual([{ start: 0, end: 10 }, { start: 10, end: 15 }]);
  });

  it('preserves the uniform ruby advance when the exact suffix enters a narrower column', () => {
    const filler = cjkParagraph('う'.repeat(140), 10);
    const text = 'あ'.repeat(10) + 'い'.repeat(20);
    const target = para('', {
      runs: [
        sizedTextRun('あ'.repeat(10), 10, { text: 'ルビ', fontSizePt: 20 }),
        sizedTextRun('い'.repeat(20), 10),
      ],
      defaultFontSize: 10,
    });
    const model = columnDocument(
      [filler as unknown as BodyElement, target as unknown as BodyElement],
      [100, 48], [12, 0], 140,
    );
    (model.section as SectionProps).docGridType = 'lines';
    (model.section as SectionProps).docGridLinePitch = 10;
    const fragments = sourceFragments(layoutDocument(model), 1);
    const advances = fragments.flatMap((fragment) => fragment.lines.map((line) => line.advancePt));

    expect(fragments.length).toBeGreaterThan(1);
    expect(fragmentText(fragments)).toBe(text);
    expect(new Set(advances)).toEqual(new Set([advances[0]]));
  });
});
