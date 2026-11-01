import { existsSync, readFileSync } from 'node:fs';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { storeZip } from '../conformance/generate.js';
import { describe, expect, it, vi } from 'vitest';

// Counts body pagination passes: each pass opens exactly one kernel session.
const passes = vi.hoisted(() => ({ opened: 0, finalizedPages: 0, sessionCalls: 0 }));
vi.mock('./page-factory.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./page-factory.js')>();
  return {
    ...actual,
    finalizeLayoutPage: (...args: Parameters<typeof actual.finalizeLayoutPage>) => {
      passes.finalizedPages += 1;
      return actual.finalizeLayoutPage(...args);
    },
  };
});
vi.mock('./runtime-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./runtime-state.js')>();
  return {
    ...actual,
    bodyLayoutKernelOf: (services: Parameters<typeof actual.bodyLayoutKernelOf>[0]) => {
      const kernel = actual.bodyLayoutKernelOf(services);
      return kernel && {
        openBodyLayoutSession: (...args: Parameters<typeof kernel.openBodyLayoutSession>) => {
          passes.opened += 1;
          // Every session call is layout work; counting them between host
          // turns measures how long the host waited.
          const session = kernel.openBodyLayoutSession(...args);
          return Object.fromEntries(Object.entries(session).map(([key, value]) => [
            key,
            typeof value === 'function'
              ? (...callArgs: unknown[]) => { passes.sessionCalls += 1; return value.apply(session, callArgs); }
              : value,
          ])) as typeof session;
        },
      };
    },
  };
});

import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import type { BodyElement, DocxDocumentModel } from '../types.js';
import { paginateBody } from './body-paginator.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { PaginationAbortError } from './pagination-scheduler.js';
import { layoutDocumentProgressively, type ProgressiveLayoutPreview } from './progressive.js';
import type { DocumentLayout, LayoutPage } from './types.js';

// Page-owned anchors (§20.4.3.5 positionV relativeFrom="margin") and
// page-positioned floating tables (§17.4.57 tblpPr vertAnchor="page") are
// resolved by exact-state convergence: the first pass prescans anchors in
// source order, later passes apply the destinations and table page floors the
// previous pass proved. A page published from a pass that a later pass
// supersedes would be revised. These cases pin that such documents still
// publish progressively and that every published page is already the page of
// the final layout.
//
// Pagination here is independent of glyph metrics, so the pass counts and
// publication points are identical on every host: each paragraph is a single
// short glyph on an exact 24 pt line (§17.3.1.33), measured by an injected
// linear-metric context, and every anchor and table has explicit geometry.
// The 468 pt text column (US Letter, 1 in margins) holds 27 lines per page.

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const context = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return context as unknown as CanvasRenderingContext2D;
}

function line(extraRuns: readonly unknown[] = []): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: 24, rule: 'exact', explicit: true },
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'x', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'NotInMetrics',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }, ...extraRuns],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as BodyElement;
}

/** A full-width, margin-relative image that text must clear (wrapTopAndBottom,
 * §20.4.2.20). Its paragraph is a page-owned anchor owner. */
function anchoredImageLine(heightPt: number): BodyElement {
  return line([{
    type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
    widthPt: 468, heightPt, anchor: true, anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: true, anchorYFromPara: false,
    wrapMode: 'topAndBottom', wrapSide: 'bothSides',
    anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'margin',
  }]);
}

/** A page-positioned floating table of empty exact-height rows. */
function pageFloatingTable(
  widthPt: number,
  tblpY: number,
  rows: number,
  rowHeightPt: number,
  topFromText: number,
): BodyElement {
  const noBorders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
  return {
    type: 'table',
    colWidths: [widthPt],
    rows: Array.from({ length: rows }, () => ({
      cells: [{
        content: [], colSpan: 1, vMerge: null, borders: noBorders,
        background: null, vAlign: 'top', widthPt: null,
      }],
      rowHeight: rowHeightPt, rowHeightRule: 'exact', isHeader: false,
    })),
    borders: noBorders,
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText, bottomFromText: 0,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY,
    },
  } as unknown as BodyElement;
}

/** An ordinary (in-flow) table of empty exact-height rows across the column. */
function ordinaryTable(rows: number, rowHeightPt: number): BodyElement {
  const { tblpPr: _tblpPr, ...table } = pageFloatingTable(468, 0, rows, rowHeightPt, 0) as unknown as
    Record<string, unknown>;
  return table as unknown as BodyElement;
}

/** `lines` single-line paragraphs; each insert lands at its body index, in order. */
function documentModel(
  lines: number,
  inserts: readonly (readonly [number, BodyElement])[],
): DocxDocumentModel {
  const body = Array.from({ length: lines }, () => line());
  for (const [index, element] of inserts) body.splice(index, 0, element);
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null,
    },
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function open(model: DocxDocumentModel) {
  const source = layoutSourceStore(model);
  return {
    input: source.bodyLayoutInput,
    services: createLayoutServices(source, { measureContext: measureContext() }),
    options: normalizeLayoutOptions(undefined, 1_700_000_000_000),
  };
}

function pageFingerprint(page: LayoutPage): string {
  return layoutFingerprint({ pages: [page], diagnostics: [] } as DocumentLayout);
}

function blockingLayout(model: DocxDocumentModel): Readonly<{ layout: DocumentLayout; passes: number }> {
  const before = passes.opened;
  const blocking = open(model);
  const layout = paginateBody(blocking.input, blocking.services, blocking.options);
  return { layout, passes: passes.opened - before };
}

async function progressiveRun(model: DocxDocumentModel) {
  const previews: ProgressiveLayoutPreview[] = [];
  const progressive = open(model);
  const final = await layoutDocumentProgressively(
    progressive.input,
    progressive.services,
    progressive.options,
    { onPreview: (preview) => { previews.push(preview); } },
  );
  return { previews, final };
}

/** The page holding the root of the body-level table at `bodyIndex`, and the
 * page where the flow reached it: the page holding the block before it. */
function tablePages(layout: DocumentLayout, bodyIndex: number) {
  const pagesWith = (predicate: (node: LayoutPage['layers']['body'][number]) => boolean) => (
    layout.pages.filter((page) => page.layers.body.some(predicate)).map((page) => page.pageIndex)
  );
  const placed = pagesWith((node) => node.kind === 'table' && node.source.path[0] === bodyIndex);
  const preceding = pagesWith((node) => node.source.path[0] === bodyIndex - 1);
  return { reachedPage: preceding.at(-1)!, placedPage: placed[0]! };
}

function expectPublishedPagesFinal(
  previews: readonly ProgressiveLayoutPreview[],
  final: DocumentLayout,
): void {
  const counts = previews.map((preview) => preview.layout.pages.length);
  expect(counts).toEqual([...counts].sort((left, right) => left - right));
  for (const preview of previews) {
    expect(preview.exact).toBe(false);
    expect(preview.layout.pages.length).toBeLessThanOrEqual(final.pages.length);
    preview.layout.pages.forEach((page, index) => {
      expect(pageFingerprint(page as LayoutPage))
        .toBe(pageFingerprint(final.pages[index] as LayoutPage));
    });
  }
}

// Line 60 sits on page 2 in source order. The first pass prescans its image
// on pages 0-2, the second applies it on the page it reached, and the third
// confirms the moved destination.
const anchoredImageModel = () => documentModel(300, [[60, anchoredImageLine(200)]]);

describe('progressive layout with page-owned anchors', () => {
  it('publishes only pages every later convergence pass reproduces', async () => {
    const model = anchoredImageModel();
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(3);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // The anchored document now publishes before layout completes.
    expect(previews.length).toBeGreaterThan(1);
    expect(previews[0]!.layout.pages.length).toBeLessThan(final.pages.length);
    expectPublishedPagesFinal(previews, final);
  }, 300_000);

  it('keeps pages before a floored table publishable while its floor changes', async () => {
    // The full-width table's exclusion sends the lines before it on its page
    // past that page. The next pass therefore proves a minimum page for the
    // table (none, then one past its first page), and the converged pass reads
    // that floor where the flow reaches the table and advances past it.
    const tableIndex = 70;
    const model = documentModel(300, [[tableIndex, pageFloatingTable(468, 300, 4, 100, 20)]]);
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(3);
    const { reachedPage, placedPage } = tablePages(blocking.layout, tableIndex);
    expect(placedPage).toBe(reachedPage + 1);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    expect(previews.length).toBeGreaterThan(0);
    expectPublishedPagesFinal(previews, final);
    // Every page before the reach page is published. A floor that acts at the
    // reach page bounds the pass there, so that page is never published early.
    expect(Math.max(...previews.map((preview) => preview.layout.pages.length)))
      .toBe(reachedPage);
  }, 300_000);

  it('lowers the bound to the reach page of a table placed at or past it', async () => {
    // A page-positioned table that ends below the bottom margin, with text
    // before it on the page where the flow reaches it, moves to the next page
    // (Windows Word #1902 controls). A checkpoint whose live page holds the
    // moved table must not publish its reach page, because a later floor
    // could still act there.
    const secondTableIndex = 62;
    const model = documentModel(300, [
      [secondTableIndex, pageFloatingTable(300, 400, 3, 110, 0)],
    ]);
    const blocking = blockingLayout(model);
    const { reachedPage, placedPage } = tablePages(blocking.layout, secondTableIndex);
    expect(placedPage).toBe(reachedPage + 1);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    expectPublishedPagesFinal(previews, final);
    const counts = previews.map((preview) => preview.layout.pages.length);
    // The checkpoint while the deferred table's page is live stops at the
    // reach page instead of publishing it (which would be `placedPage` pages),
    // and later checkpoints extend past both once that page has closed.
    expect(counts).toContain(reachedPage);
    expect(counts).not.toContain(placedPage);
    expect(counts.at(-1)!).toBeGreaterThan(placedPage);
  }, 300_000);

  it('publishes the opening pages of a document that opens with a page table during discovery', async () => {
    // A page-positioned table on the first page: the source-order first pass
    // cannot have registered it at the page start, so none of its pages is
    // final. The second pass, applying the plan of the pages the first has
    // closed, runs behind it and publishes before discovery finishes. A
    // three-page ordinary table makes one body entry of the second pass open
    // pages whose plan the first pass has not closed yet.
    const lines = 300;
    const model = documentModel(lines, [
      [2, pageFloatingTable(468, 300, 2, 100, 0)],
      [40, ordinaryTable(3 * 27, 24)],
    ]);
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(2);
    expect(tablePages(blocking.layout, 2).placedPage).toBe(0);

    let steps = 0;
    let stepsAtFirstPreview: number | null = null;
    const previews: ProgressiveLayoutPreview[] = [];
    const run = open(model);
    const before = passes.opened;
    const final = await layoutDocumentProgressively(run.input, run.services, run.options, {
      scheduler: { onProgress: () => { steps += 1; } },
      onPreview: (preview) => {
        stepsAtFirstPreview ??= steps;
        previews.push(preview);
      },
    });
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // The overlapped second pass is the solver's second pass, not extra work.
    expect(passes.opened - before).toBe(blocking.passes);
    expectPublishedPagesFinal(previews, final);
    // The first pass alone suspends once per body entry; the opening preview
    // arrives well before it could have finished.
    expect(stepsAtFirstPreview).not.toBeNull();
    expect(stepsAtFirstPreview!).toBeLessThan(run.input.sequence.length / 4);
    expect(previews.at(-1)!.layout.pages.length).toBeGreaterThan(5);
  }, 300_000);

  it('keeps host turns frequent while the second pass waits on the first', async () => {
    // The ordinary table makes one second-pass entry wait for pages the first
    // pass has not closed; the many one-line paragraphs after it are what the
    // first pass must lay out meanwhile. Each must stay a suspension point.
    const model = documentModel(300, [
      [2, pageFloatingTable(468, 72, 3, 200, 0)],
      [6, pageFloatingTable(468, 72, 3, 200, 0)],
      [10, pageFloatingTable(468, 72, 3, 200, 0)],
      [40, ordinaryTable(4 * 27, 24)],
    ]);
    const blocking = blockingLayout(model);
    const run = open(model);
    let since = 0;
    let widest = 0;
    let last = passes.sessionCalls;
    const final = await layoutDocumentProgressively(run.input, run.services, run.options, {
      onPreview: () => {},
      scheduler: {
        now: () => Number.MAX_SAFE_INTEGER,
        sliceMs: 0,
        yieldToHost: () => {
          since = passes.sessionCalls - last;
          widest = Math.max(widest, since);
          last = passes.sessionCalls;
          return Promise.resolve();
        },
      },
    });
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // One body entry here is at most about fifteen session calls. Waiting
    // synchronously for the first pass instead ran a whole page of entries
    // (over 180 calls) inside one host turn.
    expect(widest).toBeLessThan(40);
  }, 300_000);

  it('discovers first-pass destinations in work linear in the page count', async () => {
    const finalizations = async (pages: number) => {
      const model = documentModel(pages * 27, [[2, pageFloatingTable(468, 300, 2, 100, 0)]]);
      const run = open(model);
      const before = passes.finalizedPages;
      await layoutDocumentProgressively(run.input, run.services, run.options, { onPreview: () => {} });
      return passes.finalizedPages - before;
    };
    const small = await finalizations(11);
    const medium = await finalizations(21);
    const large = await finalizations(41);
    // Doubling the document may at most roughly double the finalized pages;
    // a full-prefix snapshot per closed page would quadruple them.
    expect(medium / small).toBeLessThan(2.5);
    expect(large / medium).toBeLessThan(2.5);
  }, 300_000);

  it('discovers destinations on closed pages of a section with page borders', async () => {
    // #1949: the opening pass finalizes each closed page to read its anchor
    // destinations; a bordered section needs that page's section-owned
    // identity (§17.6.10 display) there as in final composition.
    const model = anchoredImageModel();
    const edge = { style: 'single', width: 1.5, space: 24 };
    model.section.pageBorders = {
      offsetFrom: 'page', display: 'firstPage', zOrder: 'front',
      top: edge, bottom: edge, left: edge, right: edge,
    };
    const blocking = blockingLayout(model);
    expect(blocking.layout.pages.map((page) => page.pageBorder !== null).slice(0, 2))
      .toEqual([true, false]);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    expect(previews.length).toBeGreaterThan(0);
    expectPublishedPagesFinal(previews, final);
  }, 300_000);

  it('emits no stale page when cancelled during convergence', async () => {
    const model = anchoredImageModel();
    const { layout: final } = blockingLayout(model);
    const everyStep = (onYield: () => void) => ({
      now: () => Number.MAX_SAFE_INTEGER,
      sliceMs: 0,
      yieldToHost: () => { onYield(); return Promise.resolve(); },
    });

    let totalYields = 0;
    const complete = open(model);
    await layoutDocumentProgressively(complete.input, complete.services, complete.options, {
      onPreview: () => {},
      scheduler: everyStep(() => { totalYields += 1; }),
    });

    // Abort points inside the second and third passes, after publications
    // have started, as a destroyed viewer or a replaced load would.
    for (const fraction of [0.5, 0.7, 0.9]) {
      const abortAt = Math.floor(totalYields * fraction);
      const controller = new AbortController();
      const previews: ProgressiveLayoutPreview[] = [];
      let yields = 0;
      let publishedAfterAbort = false;
      const cancelled = open(model);
      await expect(layoutDocumentProgressively(
        cancelled.input,
        cancelled.services,
        cancelled.options,
        {
          onPreview: (preview) => {
            if (controller.signal.aborted) publishedAfterAbort = true;
            previews.push(preview);
          },
          scheduler: {
            ...everyStep(() => {
              yields += 1;
              if (yields === abortAt) controller.abort();
            }),
            signal: controller.signal,
          },
        },
      )).rejects.toBeInstanceOf(PaginationAbortError);
      expect(previews.length).toBeGreaterThan(0);
      expect(publishedAfterAbort).toBe(false);
      expectPublishedPagesFinal(previews, final);
    }
  }, 300_000);
});

/** Six nextPage sections, each a page of text and then a near-page-tall page
 * table whose exclusion pushes that text off its page (#2012). */
function chainedTableCycleSections(): Readonly<{ model: DocxDocumentModel; tableIndexes: number[] }> {
  const model = documentModel(0, []);
  model.section.pageHeight = 842;
  const sectionBreak = { type: 'sectionBreak', kind: 'nextPage', columns: null } as unknown as BodyElement;
  const tableIndexes: number[] = [];
  for (let section = 0; section < 6; section += 1) {
    if (section > 0) model.body.push(sectionBreak);
    model.body.push(...Array.from({ length: 26 }, () => line()));
    tableIndexes.push(model.body.length);
    model.body.push(pageFloatingTable(524, 0, 1, 667, 0), line(), line());
  }
  return { model, tableIndexes };
}

/** One parsed paragraph holding a top-and-bottom paragraph-relative 200x60 pt
 * picture whose top is 6 pt above its anchor line. */
async function parsedTightPictureParagraph(): Promise<BodyElement> {
  await init({ module_or_path: readFileSync(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const EMU = 12700;
  const picture = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" '
    + 'simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
    + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    + `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${-6 * EMU}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${200 * EMU}" cy="${60 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + '<wp:wrapTopAndBottom/><wp:docPr id="1" name="Picture 1"/>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
    + '<pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
    + '<a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>'
    + `<a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${60 * EMU}"/></a:xfrm><a:prstGeom prst="rect">`
    + '<a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const document = `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" `
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'
    + '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>'
    + `<w:r><w:t>x</w:t></w:r>${picture}</w:p><w:sectPr/></w:body></w:document>`;
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const encoder = new TextEncoder();
  const archive = new DocxArchive(storeZip(new Map(Object.entries({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdImg" Type="${rel}/image" Target="media/p.png"/></Relationships>`),
    'word/document.xml': encoder.encode(document),
  }))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document.body[0]!;
  } finally {
    archive.free();
  }
}

/** The lines around the picture of `layout`, relative to the picture's top. */
function pictureNeighbourhood(layout: DocumentLayout): readonly number[] {
  for (const page of layout.pages) {
    const paragraphs = page.layers.body.filter((node) => node.kind === 'paragraph');
    const drawing = paragraphs.flatMap((node) => node.drawings)[0];
    if (!drawing) continue;
    return paragraphs.flatMap((node) => node.lines)
      .map((entry) => entry.bounds.yPt - drawing.flowBounds.yPt)
      .filter((delta) => Math.abs(delta) < 150);
  }
  throw new Error('No page holds the picture');
}

describe('page-table destination cycle settlement', () => {
  const coupledTables = () => documentModel(33, [
    [13, pageFloatingTable(468, 72, 1, 360, 0)],
    [15, pageFloatingTable(468, 72, 1, 600, 0)],
    [33, pageFloatingTable(300, 72, 1, 600, 0)],
  ]);

  it('settles a near-page-tall table on the later page of its two-state cycle', () => {
    const model = documentModel(56, [
      [26, pageFloatingTable(524, 0, 1, 667, 0)],
      [55, pageFloatingTable(524, 0, 1, 667, 0)],
    ]);
    model.section.pageHeight = 842;
    const { layout } = blockingLayout(model);
    // Main alternates (0, 1) and (2, 3). Each table's exclusion would push
    // the text before it off its page, so it moves to the next page with that
    // text (Windows Word #1902 controls); the text resumes below the table.
    expect(tablePages(layout, 26).placedPage).toBe(1);
    expect(tablePages(layout, 55).placedPage).toBe(3);
    for (const index of [26, 55]) {
      const page = layout.pages[tablePages(layout, index).placedPage]!;
      const table = page.layers.body.find((node) => node.kind === 'table')!;
      const lines = page.layers.body.flatMap((node) => node.kind === 'paragraph' ? node.lines : []);
      expect(lines.length).toBeGreaterThan(0);
      expect(lines.every((line) =>
        line.bounds.yPt >= table.flowBounds.yPt + table.flowBounds.heightPt)).toBe(true);
    }
  });

  it('settles near-page-tall table cycles in consecutive sections', async () => {
    // Each table belongs on the page after its section's text. Moving one table
    // shifts every later section's table by a page; that shift must not hide
    // the earlier table's proven floor, or the run alternates until its pass
    // limit (#2012; Word places the controls' tables the same way).
    const { model, tableIndexes } = chainedTableCycleSections();
    const { layout } = blockingLayout(model);
    expect(tableIndexes.map((index) => tablePages(layout, index)))
      .toEqual([0, 2, 4, 6, 8, 10].map((page) => ({ reachedPage: page, placedPage: page + 1 })));
    for (const index of tableIndexes) {
      const page = layout.pages[tablePages(layout, index).placedPage]!;
      const table = page.layers.body.find((node) => node.kind === 'table')!;
      const lines = page.layers.body.flatMap((node) => node.kind === 'paragraph' ? node.lines : []);
      expect(lines.every((line) =>
        line.bounds.yPt >= table.flowBounds.yPt + table.flowBounds.heightPt)).toBe(true);
    }
    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(layout));
    expectPublishedPagesFinal(previews, final);
  }, 30_000);

  it('keeps those table floors through the later text-frame carry', () => {
    // A text frame after the tables starts a frame-carry run, which settles the
    // same tables again and needs the same page-scoped floor proof.
    const { model, tableIndexes } = chainedTableCycleSections();
    const framePr = {
      dropCap: 'none', lines: 1, wrap: 'notBeside', hAnchor: 'page', vAnchor: 'text',
      hRule: 'exact', h: 60, hSpace: 0, vSpace: 10, w: 100, xAlign: 'center',
    };
    model.body.push(...Array.from({ length: 50 }, () => line()));
    model.body.push({ ...line(), framePr } as unknown as BodyElement);
    model.body.push(...Array.from({ length: 50 }, () => line()));
    const { layout } = blockingLayout(model);
    expect(tableIndexes.map((index) => tablePages(layout, index).placedPage)).toEqual([1, 3, 5, 7, 9, 11]);
  }, 30_000);

  it('keeps those table floors through the later host-drawing carry', async () => {
    // A paragraph-relative top-and-bottom picture reaching 6 pt above its anchor
    // line is carried to the start of its page (WORD_LATER_ANCHOR_EARLIER_LINE_WRAP).
    // That run settles the same tables again, so it starts from their floors:
    // the picture keeps its carry and the earlier lines wrap around it, exactly
    // as in the same document without the tables.
    const pictureParagraph = await parsedTightPictureParagraph();
    const withPicture = (model: DocxDocumentModel) => {
      model.body.push(...Array.from({ length: 50 }, () => line()), pictureParagraph,
        ...Array.from({ length: 50 }, () => line()));
      return model;
    };
    const { model, tableIndexes } = chainedTableCycleSections();
    const { layout } = blockingLayout(withPicture(model));
    expect(tableIndexes.map((index) => tablePages(layout, index).placedPage)).toEqual([1, 3, 5, 7, 9, 11]);
    const plain = documentModel(0, []);
    plain.section.pageHeight = 842;
    const { layout: plainLayout } = blockingLayout(withPicture(plain));
    expect(pictureNeighbourhood(layout)).toEqual(pictureNeighbourhood(plainLayout));
  }, 30_000);

  it('keeps those table floors through header reserve repagination', () => {
    // A header that extends below the top margin repaginates the body; that
    // rerun settles the same tables and starts with the same floor proof.
    // Six sections fail the reserve-free run's strict proof first; four pass
    // it and fail only in the reserve run.
    for (const sections of [6, 4]) {
      const { model, tableIndexes } = chainedTableCycleSections();
      model.body = model.body.slice(0, sections * 30 - 1);
      model.headers.default = { body: [line()].map((paragraph) => ({
        ...paragraph, lineSpacing: { value: 48, rule: 'exact', explicit: true },
      })) } as unknown as DocxDocumentModel['headers']['default'];
      const { layout } = blockingLayout(model);
      for (const index of tableIndexes.slice(0, sections)) {
        const { reachedPage, placedPage } = tablePages(layout, index);
        expect(placedPage).toBe(reachedPage + 1);
      }
    }
  }, 30_000);

  it('settles the coupled T1/T2/T3 cycle deterministically without overlapping body text', () => {
    const model = coupledTables();
    const { layout } = blockingLayout(model);
    expect(layoutFingerprint(blockingLayout(model).layout)).toBe(layoutFingerprint(layout));
    expect(layout.pages.flatMap((page) => page.layers.body)
      .filter((node) => node.kind === 'table')).toHaveLength(3);
    // T1 and T2 each move off the page whose text they would cover. Both
    // permit overlap, so they may share page 1 overlapping each other
    // (§17.4.56; Windows Word #1902 controls); body text never sits under one.
    expect([13, 15, 33].map((index) => tablePages(layout, index).placedPage)).toEqual([1, 1, 2]);
    for (const page of layout.pages) {
      const tables = page.layers.body.filter((node) => node.kind === 'table');
      const lines = page.layers.body.flatMap((node) => node.kind === 'paragraph' ? node.lines : []);
      for (const table of tables) {
        for (const b of lines.map((line) => line.bounds)) {
          const a = table.flowBounds;
          expect(a.xPt >= b.xPt + b.widthPt || b.xPt >= a.xPt + a.widthPt
            || a.yPt >= b.yPt + b.heightPt || b.yPt >= a.yPt + a.heightPt).toBe(true);
        }
      }
    }
  });

  it('keeps cycle-settled progressive pages equal to the blocking result', async () => {
    const model = coupledTables();
    const { final, previews } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blockingLayout(model).layout));
    expectPublishedPagesFinal(previews, final);
  });
});

// Local Word controls exercise the parser-to-paginator path without publishing
// the non-redistributable fixtures. Placement fidelity is informational here:
// this regression guarantees termination, not Word's logical anchor rule.
describe('local page-table boundary controls', () => {
  for (const name of ['heights', 'source-boundaries', 'wrap-widths', 'zero-boundaries']) {
    const path = new URL(`../../public/private/docx/controls-1659/page-anchor-${name}.docx`, import.meta.url);
    it.skipIf(!existsSync(path))(`lays out the ${name} controls without error`, async () => {
      await init({ module_or_path: readFileSync(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
      const archive = new DocxArchive(readFileSync(path));
      let model: DocxDocumentModel;
      try {
        model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
      } finally {
        archive.free();
      }
      const { layout } = blockingLayout(model);
      const tables = layout.pages.flatMap((page) => page.layers.body
        .filter((node) => node.kind === 'table').map((node) => ({
          source: node.source.path, page: page.pageIndex + 1, bounds: node.flowBounds,
        })));
      expect(tables).toHaveLength(model.body.filter((node) => node.type === 'table').length);
      expect(layout.diagnostics.some((entry) => entry.code === 'NON_CONVERGENCE')).toBe(false);
    }, 30_000);
  }
});
