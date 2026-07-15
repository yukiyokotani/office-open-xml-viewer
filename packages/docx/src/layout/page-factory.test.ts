import { describe, expect, it } from 'vitest';
import { createCanvasFontRoute } from '@silurus/ooxml-core';
import type { SectionLayoutContext } from '../layout-context.js';
import {
  accumulatePagePaintNode,
  accumulatePageSectionRegion,
  bodyFlowDomainId,
  createLayoutPageAccumulator,
  createLayoutPage,
  createParityBlankLayoutPage,
  finalizeLayoutPage,
} from './page-factory.js';
import { assertDocumentLayout } from './invariants.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutRect,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  TextBoxLayout,
} from './types.js';

const rect = (xPt: number, yPt: number, widthPt: number, heightPt: number): LayoutRect => ({
  xPt,
  yPt,
  widthPt,
  heightPt,
});

const source = (path: readonly number[]): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path,
});

function section(
  textDirection: string,
  columns: readonly Readonly<{ xPt: number; wPt: number }>[],
): SectionLayoutContext {
  return {
    geometry: {
      pageWidth: 612,
      pageHeight: 792,
      marginTop: 72,
      marginRight: 72,
      marginBottom: 72,
      marginLeft: 72,
      headerDistance: 36,
      footerDistance: 36,
    },
    columns,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection,
    verticalAlignment: 'top',
  };
}

function drawing(id: string, flowDomainId: string, bounds: LayoutRect): DrawingLayout {
  return {
    kind: 'drawing',
    id,
    source: source([Number(id.replace(/\D/g, '')) || 0]),
    flowDomainId,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: bounds.heightPt,
    ordinaryFlow: true,
    commands: [],
  };
}

function bookmarkParagraph(
  id: string,
  flowDomainId: string,
  bookmark: string,
  bounds: LayoutRect,
): ParagraphLayout {
  return {
    kind: 'paragraph',
    id,
    source: source([3]),
    flowDomainId,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: bounds.heightPt,
    ordinaryFlow: true,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 1 },
      bounds,
      baselinePt: bounds.yPt + 10,
      advancePt: bounds.heightPt,
      placements: [{
        kind: 'text',
        text: 'x',
        range: { start: 0, end: 1 },
        origin: { xPt: bounds.xPt, yPt: bounds.yPt + 10 },
        bounds,
        advancePt: bounds.widthPt,
        clusters: [{
          range: { start: 0, end: 1 },
          offset: { xPt: 0, yPt: 0 },
          advancePt: bounds.widthPt,
        }],
        paintOps: [],
        color: { kind: 'default' },
        fontRoute: createCanvasFontRoute('sans-serif', 'generic'),
        fontSizePt: 10,
        fontWeight: 400,
        fontStyle: 'normal',
        direction: 'ltr',
        decorations: [],
        bookmark,
      }],
    }],
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  };
}

function bookmarkTextBox(
  id: string,
  flowDomainId: string,
  paragraph: ParagraphLayout,
  bounds: LayoutRect,
): TextBoxLayout {
  return {
    kind: 'textbox', id, source: source([4]), flowDomainId,
    flowBounds: bounds, inkBounds: bounds, advancePt: 0, ordinaryFlow: false,
    paragraphs: [paragraph], writingMode: 'horizontal-tb',
    insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
  };
}

function bookmarkTable(
  id: string,
  flowDomainId: string,
  blocks: readonly (ParagraphLayout | TableLayout)[],
  bounds: LayoutRect,
): TableLayout {
  return {
    kind: 'table', id, source: source([5]), flowDomainId,
    flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt, ordinaryFlow: true,
    columnWidthsPt: [bounds.widthPt], borders: [],
    rows: [{
      kind: 'table-row', id: `${id}:row`, source: source([5, 0]), flowDomainId,
      flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt, ordinaryFlow: true,
      heightPt: bounds.heightPt, contentHeightPt: bounds.heightPt,
      cells: [{
        kind: 'table-cell', id: `${id}:cell`, source: source([5, 0, 0]), flowDomainId,
        flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt, ordinaryFlow: true,
        contentBounds: bounds, verticalMerge: 'none', vAlign: 'top',
        blocks: blocks.map((layout) => ({ layout, offsetPt: 0, advancePt: layout.advancePt })),
      }],
    }],
  };
}

describe('createLayoutPage', () => {
  it('accumulates page inputs without mutating prior flow state', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const initial = createLayoutPageAccumulator({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:body',
      section: bodySection,
    });
    const withRegion = accumulatePageSectionRegion(initial, {
      id: 'region:body',
      sectionOccurrenceId: 'section:body',
      section: bodySection,
      writingMode: 'horizontal-tb',
      blockStartPt: 72,
      blockEndPt: 720,
      columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
    });
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const node = drawing('drawing-1', domainId, rect(72, 100, 50, 10));
    const complete = accumulatePagePaintNode(withRegion, {
      layer: 'body',
      node,
    }, true);

    expect(initial.sectionRegions).toEqual([]);
    expect(initial.paint).toEqual([]);
    expect(withRegion.paint).toEqual([]);
    expect(complete.readingOrder).toEqual([node]);
    expect(finalizeLayoutPage(complete, {
      displayNumber: 1,
      format: 'decimal',
      sectionOccurrenceId: 'section:body',
    }).layers.body).toEqual([node]);
  });

  it('creates physical geometry and distinct body domains for continuous section regions', () => {
    const first = section('lrTb', [
      { xPt: 72, wPt: 220 },
      { xPt: 320, wPt: 220 },
    ]);
    const second = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const firstDomain = bodyFlowDomainId(0, 'region:first', 0);
    const secondDomain = bodyFlowDomainId(0, 'region:second', 0);
    const firstNode = drawing('drawing-1', firstDomain, rect(72, 100, 200, 20));
    const secondNode = drawing('drawing-2', secondDomain, rect(72, 360, 200, 20));

    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:first',
      section: first,
      sectionRegions: [
        {
          id: 'region:first',
          sectionOccurrenceId: 'section:first',
          section: first,
          writingMode: 'horizontal-tb',
          blockStartPt: 72,
          blockEndPt: 330,
          columns: [
            { inlineStartPt: 72, inlineExtentPt: 220 },
            { inlineStartPt: 320, inlineExtentPt: 220 },
          ],
        },
        {
          id: 'region:second',
          sectionOccurrenceId: 'section:second',
          section: second,
          writingMode: 'horizontal-tb',
          blockStartPt: 330,
          blockEndPt: 720,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        },
      ],
      paint: [
        { layer: 'body', node: firstNode },
        { layer: 'body', node: secondNode },
      ],
      readingOrder: [firstNode, secondNode],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:first',
      },
    });

    expect(page.geometry).toEqual({
      xPt: 0,
      yPt: 0,
      widthPt: 612,
      heightPt: 792,
      contentTopPt: 72,
      contentBottomPt: 720,
    });
    expect(page.sectionRegions).toHaveLength(2);
    expect(page.sectionRegions?.map((region) => region.flowDomainIds)).toEqual([
      [
        bodyFlowDomainId(0, 'region:first', 0),
        bodyFlowDomainId(0, 'region:first', 1),
      ],
      [bodyFlowDomainId(0, 'region:second', 0)],
    ]);
    expect(page.flowDomains.map((domain) => domain.bounds)).toEqual([
      rect(72, 72, 220, 258),
      rect(320, 72, 220, 258),
      rect(72, 330, 468, 390),
    ]);
    expect(page.sectionOccurrenceId).toBe('section:first');
    expect(page.pageNumber).toEqual({
      displayNumber: 1,
      format: 'decimal',
      sectionOccurrenceId: 'section:first',
    });
  });

  it('retains a logical-to-physical transform for vertical section regions', () => {
    const vertical = section('tbRl', [{ xPt: 72, wPt: 648 }]);

    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:vertical',
      section: vertical,
      sectionRegions: [{
        id: 'region:vertical',
        sectionOccurrenceId: 'section:vertical',
        section: vertical,
        writingMode: 'vertical-rl',
        blockStartPt: 72,
        blockEndPt: 540,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 648 }],
      }],
      paint: [],
      readingOrder: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:vertical',
      },
    });

    expect(page.sectionRegions?.[0]?.coordinateSpace).toEqual({
      writingMode: 'vertical-rl',
      logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 612, f: 0 },
    });
    expect(page.flowDomains[0]?.bounds).toEqual(rect(72, 72, 468, 648));
  });

  it('uses caller-resolved effective body edges for positive and negative margin policies', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const positive = createParityBlankLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792,
        contentTopPt: 96, contentBottomPt: 676,
      },
      sectionOccurrenceId: 'section:positive', section: bodySection,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:positive' },
    });
    const negative = createParityBlankLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792,
        contentTopPt: 36, contentBottomPt: 738,
      },
      sectionOccurrenceId: 'section:negative', section: bodySection,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:negative' },
    });

    expect(positive.geometry).toMatchObject({ contentTopPt: 96, contentBottomPt: 676 });
    expect(negative.geometry).toMatchObject({ contentTopPt: 36, contentBottomPt: 738 });
  });

  it('keeps resolved body edges independent of the vertical-lr coordinate transform', () => {
    const vertical = section('tbLr', [{ xPt: 40, wPt: 700 }]);
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792,
        contentTopPt: 88, contentBottomPt: 690,
      },
      sectionOccurrenceId: 'section:vertical-lr', section: vertical,
      sectionRegions: [{
        id: 'region:vertical-lr', sectionOccurrenceId: 'section:vertical-lr',
        section: vertical, writingMode: 'vertical-lr', blockStartPt: 54, blockEndPt: 564,
        columns: [{ inlineStartPt: 40, inlineExtentPt: 700 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:vertical-lr' },
    });

    expect(page.geometry).toMatchObject({ contentTopPt: 88, contentBottomPt: 690 });
    expect(page.sectionRegions?.[0]?.coordinateSpace?.logicalToPhysical)
      .toEqual({ a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 });
  });

  it('builds layer order, reading order, and clone-safe bookmark ownership', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const behind = drawing('drawing-1', domainId, rect(72, 80, 50, 10));
    const paragraph = bookmarkParagraph('paragraph-3', domainId, 'destination', rect(72, 100, 50, 12));

    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:body',
      section: bodySection,
      sectionRegions: [{
        id: 'region:body',
        sectionOccurrenceId: 'section:body',
        section: bodySection,
        writingMode: 'horizontal-tb',
        blockStartPt: 72,
        blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [
        { layer: 'behindText', node: behind },
        { layer: 'body', node: paragraph },
      ],
      readingOrder: [paragraph],
      pageNumber: {
        displayNumber: 7,
        format: 'lowerRoman',
        sectionOccurrenceId: 'section:body',
      },
    });

    expect(page.layers.behindText).toEqual([behind]);
    expect(page.layers.body).toEqual([paragraph]);
    expect(page.layers.paintOrder).toEqual([
      { layer: 'behindText', nodeId: 'drawing-1' },
      { layer: 'body', nodeId: 'paragraph-3' },
    ]);
    expect(page.readingOrder).toEqual(['paragraph-3']);
    expect(page.bookmarkStarts).toEqual([{
      name: 'destination',
      nodeId: 'paragraph-3',
      sectionOccurrenceId: 'section:body',
    }]);
    expect(structuredClone(page)).toEqual(page);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();
  });

  it('validates bookmark destinations against nested tables and text boxes in the retained graph', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const nestedParagraph = bookmarkParagraph(
      'paragraph:nested-table', domainId, 'nested-table-destination', rect(80, 110, 100, 12),
    );
    const nestedTable = bookmarkTable(
      'table:nested', domainId, [nestedParagraph], rect(78, 105, 120, 20),
    );
    const outerTable = bookmarkTable(
      'table:outer', domainId, [nestedTable], rect(72, 100, 150, 30),
    );
    const textBoxParagraph = bookmarkParagraph(
      'paragraph:textbox', domainId, 'textbox-destination', rect(250, 150, 100, 12),
    );
    const textBox = bookmarkTextBox(
      'textbox:outer', domainId, textBoxParagraph, rect(240, 140, 120, 30),
    );
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792,
        contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:body', section: bodySection,
      sectionRegions: [{
        id: 'region:body', sectionOccurrenceId: 'section:body', section: bodySection,
        writingMode: 'horizontal-tb', blockStartPt: 72, blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [{ layer: 'body', node: outerTable }, { layer: 'front', node: textBox }],
      readingOrder: [outerTable, textBox],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });

    expect(page.bookmarkStarts).toEqual([
      {
        name: 'nested-table-destination',
        nodeId: 'paragraph:nested-table',
        sectionOccurrenceId: 'section:body',
      },
      {
        name: 'textbox-destination',
        nodeId: 'paragraph:textbox',
        sectionOccurrenceId: 'section:body',
      },
    ]);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();
  });

  it('rejects a duplicate ID inside a nested table graph', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const duplicate = bookmarkParagraph(
      'table:nested', domainId, 'nested', rect(80, 110, 100, 12),
    );
    const nested = bookmarkTable(
      'table:nested', domainId, [duplicate], rect(78, 105, 120, 20),
    );
    const outer = bookmarkTable('table:outer', domainId, [nested], rect(72, 100, 150, 30));
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720 },
      sectionOccurrenceId: 'section:body', section: bodySection,
      sectionRegions: [{
        id: 'region:body', sectionOccurrenceId: 'section:body', section: bodySection,
        writingMode: 'horizontal-tb', blockStartPt: 72, blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [{ layer: 'body', node: outer }], readingOrder: [outer],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });

    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] }))
      .toThrow(/duplicate retained node id table:nested/);
  });

  it('rejects a duplicate ID inside a text box graph', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const duplicate = bookmarkParagraph(
      'textbox:outer', domainId, 'textbox', rect(250, 150, 100, 12),
    );
    const textBox = bookmarkTextBox(
      'textbox:outer', domainId, duplicate, rect(240, 140, 120, 30),
    );
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720 },
      sectionOccurrenceId: 'section:body', section: bodySection,
      sectionRegions: [{
        id: 'region:body', sectionOccurrenceId: 'section:body', section: bodySection,
        writingMode: 'horizontal-tb', blockStartPt: 72, blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [{ layer: 'front', node: textBox }], readingOrder: [textBox],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });

    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] }))
      .toThrow(/duplicate retained node id textbox:outer/);
  });

  it('lets invariants reject unknown retained section and bookmark ownership', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const paragraph = bookmarkParagraph('paragraph-3', domainId, 'destination', rect(72, 100, 50, 12));
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:body',
      section: bodySection,
      sectionRegions: [{
        id: 'region:body',
        sectionOccurrenceId: 'section:body',
        section: bodySection,
        writingMode: 'horizontal-tb',
        blockStartPt: 72,
        blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [{ layer: 'body', node: paragraph }],
      readingOrder: [paragraph],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:body',
      },
    });
    const unknownPageNumberOwner: DocumentLayout = {
      pages: [{
        ...page,
        pageNumber: { ...page.pageNumber!, sectionOccurrenceId: 'section:unknown' },
      }],
      diagnostics: [],
    };
    const unknownBookmarkNode: DocumentLayout = {
      pages: [{
        ...page,
        bookmarkStarts: [{
          ...page.bookmarkStarts![0]!,
          nodeId: 'paragraph:unknown',
        }],
      }],
      diagnostics: [],
    };

    expect(() => assertDocumentLayout(unknownPageNumberOwner)).toThrow(/page number section owner/);
    expect(() => assertDocumentLayout(unknownBookmarkNode)).toThrow(/bookmark node/);
  });
});

describe('createParityBlankLayoutPage', () => {
  it('rejects a negative physical page index at construction', () => {
    const outgoing = section('lrTb', [{ xPt: 72, wPt: 468 }]);

    expect(() => createParityBlankLayoutPage({
      pageIndex: -1,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:outgoing', section: outgoing,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:outgoing' },
    })).toThrow(RangeError);
  });

  it('requires ordered effective page edges within the physical page and permits equality', () => {
    const outgoing = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const createWithEdges = (contentTopPt: number, contentBottomPt: number) =>
      createParityBlankLayoutPage({
        pageIndex: 0,
        physicalPage: { widthPt: 612, heightPt: 792, contentTopPt, contentBottomPt },
        sectionOccurrenceId: 'section:outgoing', section: outgoing,
        pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:outgoing' },
      });

    expect(() => createWithEdges(-1, 720)).toThrow(RangeError);
    expect(() => createWithEdges(72, 793)).toThrow(RangeError);
    expect(() => createWithEdges(721, 720)).toThrow(RangeError);
    expect(() => createWithEdges(0, 0)).not.toThrow();
    expect(() => createWithEdges(792, 792)).not.toThrow();
  });

  it('owns the blank page with the outgoing section and emits no flow or paint', () => {
    const outgoing = section('lrTb', [{ xPt: 72, wPt: 468 }]);

    const page = createParityBlankLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:outgoing',
      section: outgoing,
      pageNumber: {
        displayNumber: 4,
        format: 'decimal',
        sectionOccurrenceId: 'section:outgoing',
      },
    });

    expect(page.parityBlank).toBe(true);
    expect(page.sectionOccurrenceId).toBe('section:outgoing');
    expect(page.sectionRegions).toEqual([]);
    expect(page.flowDomains).toEqual([]);
    expect(page.layers.paintOrder).toEqual([]);
    expect(page.readingOrder).toEqual([]);
    expect(page.bookmarkStarts).toEqual([]);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    const invalidBlank: DocumentLayout = {
      pages: [{
        ...page,
        flowDomains: [{
          id: 'unexpected',
          kind: 'body',
          bounds: rect(72, 72, 468, 648),
        }],
      }],
      diagnostics: [],
    };
    expect(() => assertDocumentLayout(invalidBlank)).toThrow(/parity blank/);
  });
});
