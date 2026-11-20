import { describe, expect, it } from 'vitest';
import { createCanvasFontRoute } from '@silurus/ooxml-core';
import type { SectionLayoutContext } from '../layout-context.js';
import {
  accumulatePagePaintNode,
  accumulatePageSectionRegion,
  bodyOccurrenceDestinationFor,
  bodyFlowDomainId,
  createLayoutPageAccumulator,
  createLayoutPage,
  createParityBlankLayoutPage,
  finalizeLayoutPage,
} from './page-factory.js';
import { assertDocumentLayout } from './invariants.js';
import { LayoutInvariantError } from './diagnostics.js';
import { createPageLayers } from './page-graph.js';
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
    columnSeparator: false,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection,
    verticalAlignment: 'top',
  };
}

function verticalSection(
  textDirection: string,
  columns: readonly Readonly<{ xPt: number; wPt: number }>[],
): SectionLayoutContext {
  const horizontal = section(textDirection, columns);
  return {
    ...horizontal,
    geometry: {
      ...horizontal.geometry,
      pageWidth: 792,
      pageHeight: 612,
    },
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
    bookmarkStarts: [bookmark],
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
    story: {
      story: 'textbox',
      flowBounds: paragraph.flowBounds,
      inkBounds: paragraph.inkBounds,
      blocks: [paragraph],
      advancePt: paragraph.advancePt,
      diagnostics: [],
    },
    transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    writingMode: 'horizontal-tb',
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
  it('fails closed for invalid page, identity, region, column, and one-page section facts', () => {
    const first = section('lrTb', [{ xPt: 72, wPt: 220 }]);
    const base = {
      pageIndex: 0,
      physicalPage: { widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720 },
      sectionOccurrenceId: 'section:first', section: first,
      sectionRegions: [{
        id: 'region:first', sectionOccurrenceId: 'section:first', section: first,
        writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 330,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 220 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:first' },
    };

    expect(() => createLayoutPage({ ...base, sectionOccurrenceId: '' })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{ ...base.sectionRegions[0]!, id: '' }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [
        base.sectionRegions[0]!,
        {
          ...base.sectionRegions[0]!, id: 'region:second',
          sectionOccurrenceId: 'section:second', blockStartPt: 300, blockEndPt: 500,
        },
      ],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{
        ...base.sectionRegions[0]!,
        columns: [
          { inlineStartPt: 72, inlineExtentPt: 220 },
          { inlineStartPt: 200, inlineExtentPt: 100 },
        ],
      }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{ ...base.sectionRegions[0]!, blockEndPt: 793 }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{
        ...base.sectionRegions[0]!, columns: [{ inlineStartPt: 600, inlineExtentPt: 20 }],
      }],
    })).toThrow(RangeError);
    const differentBox = {
      ...first,
      geometry: { ...first.geometry, pageWidth: 500 },
    };
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{ ...base.sectionRegions[0]!, section: differentBox }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [
        base.sectionRegions[0]!,
        {
          ...base.sectionRegions[0]!, id: 'region:second',
          sectionOccurrenceId: 'section:second', writingMode: 'vertical-rl',
          blockStartPt: 330, blockEndPt: 500,
        },
      ],
    })).toThrow(RangeError);
  });

  it.each([
    ['tbRl', 'vertical-rl'],
    ['tbRlV', 'vertical-rl'],
    ['tbLrV', 'vertical-lr'],
    ['btLr', 'vertical-rl'],
  ] as const)('accepts normalized %s section geometry for an upright physical page', (
    textDirection,
    writingMode,
  ) => {
    const normalized = verticalSection(textDirection, [{ xPt: 72, wPt: 648 }]);
    expect(() => createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:vertical',
      section: normalized,
      sectionRegions: [{
        id: 'region:vertical', sectionOccurrenceId: 'section:vertical',
        section: normalized, writingMode, blockStartPt: 72, blockEndPt: 540,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 648 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: {
        displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:vertical',
      },
    })).not.toThrow();
  });

  it.each([
    {
      textDirection: 'tb',
      writingMode: 'horizontal-tb' as const,
      sectionPage: { pageWidth: 612, pageHeight: 792 },
      column: { xPt: 72, wPt: 468 },
      blockEndPt: 720,
      logicalBounds: rect(72, 72, 468, 648),
      physicalBounds: rect(72, 72, 468, 648),
    },
    {
      textDirection: 'rl',
      writingMode: 'vertical-rl' as const,
      sectionPage: { pageWidth: 792, pageHeight: 612 },
      column: { xPt: 72, wPt: 648 },
      blockEndPt: 540,
      logicalBounds: rect(72, 72, 648, 468),
      physicalBounds: rect(72, 72, 468, 648),
    },
    {
      textDirection: 'lr',
      writingMode: 'vertical-lr' as const,
      sectionPage: { pageWidth: 792, pageHeight: 612 },
      column: { xPt: 72, wPt: 648 },
      blockEndPt: 540,
      logicalBounds: rect(72, 72, 648, 468),
      physicalBounds: rect(72, 72, 468, 648),
    },
  ])('builds normalized $writingMode geometry for Strict $textDirection', ({
    textDirection,
    writingMode,
    sectionPage,
    column,
    blockEndPt,
    logicalBounds,
    physicalBounds,
  }) => {
    const baseSection = section(textDirection, [column]);
    const normalized = {
      ...baseSection,
      geometry: { ...baseSection.geometry, ...sectionPage },
    };
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:strict',
      section: normalized,
      sectionRegions: [{
        id: 'region:strict', sectionOccurrenceId: 'section:strict',
        section: normalized, writingMode, blockStartPt: 72, blockEndPt,
        columns: [{ inlineStartPt: column.xPt, inlineExtentPt: column.wPt }],
      }],
      paint: [], readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:strict' },
    });

    expect(page.sectionRegions?.[0]?.coordinateSpace.writingMode).toBe(writingMode);
    expect(page.flowDomains[0]?.logicalBounds).toEqual(logicalBounds);
    expect(page.flowDomains[0]?.physicalBounds).toEqual(physicalBounds);
  });

  it('keeps lrTb and lrTbV horizontal and rejects unswapped vertical geometry', () => {
    for (const textDirection of ['lrTb', 'lrTbV']) {
      const horizontal = section(textDirection, [{ xPt: 72, wPt: 468 }]);
      expect(() => createLayoutPage({
        pageIndex: 0,
        physicalPage: {
          widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
        },
        sectionOccurrenceId: 'section:horizontal', section: horizontal,
        sectionRegions: [{
          id: 'region:horizontal', sectionOccurrenceId: 'section:horizontal',
          section: horizontal, writingMode: 'horizontal-tb',
          blockStartPt: 72, blockEndPt: 720,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        }],
        paint: [], readingOrder: [],
        pageNumber: {
          displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:horizontal',
        },
      })).not.toThrow();
    }

    const unswapped = section('tbRl', [{ xPt: 72, wPt: 648 }]);
    expect(() => createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:vertical', section: unswapped,
      sectionRegions: [{
        id: 'region:vertical', sectionOccurrenceId: 'section:vertical',
        section: unswapped, writingMode: 'vertical-rl',
        blockStartPt: 72, blockEndPt: 540,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 648 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: {
        displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:vertical',
      },
    })).toThrow(RangeError);
  });

  it('rejects direction, writing-mode, and logical-column disagreement', () => {
    const horizontal = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const base = {
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:body', section: horizontal,
      sectionRegions: [{
        id: 'region:body', sectionOccurrenceId: 'section:body', section: horizontal,
        writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    };

    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{ ...base.sectionRegions[0]!, writingMode: 'vertical-rl' }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{
        ...base.sectionRegions[0]!,
        columns: [{ inlineStartPt: 73, inlineExtentPt: 468 }],
      }],
    })).toThrow(RangeError);
    expect(() => createLayoutPage({
      ...base,
      sectionRegions: [{ ...base.sectionRegions[0]!, section: { ...horizontal, textDirection: '' } }],
    })).toThrow(RangeError);
  });

  it('rejects page-start section facts that contradict the first region', () => {
    const first = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const base = {
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:first', section: first,
      sectionRegions: [{
        id: 'region:first', sectionOccurrenceId: 'section:first', section: first,
        writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 720,
        columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
      }],
      paint: [], readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:first' },
    };
    const mismatches: SectionLayoutContext[] = [
      { ...first, geometry: { ...first.geometry, marginLeft: 73 } },
      { ...first, columns: [{ xPt: 73, wPt: 468 }] },
      { ...first, columnSeparator: true },
      { ...first, textDirection: 'lrTbV' },
      { ...first, grid: { ...first.grid, kind: 'lines' } },
      { ...first, verticalAlignment: 'center' },
      { ...first, lineNumbering: { start: 1, countBy: 1, restart: 'newPage' } },
    ];

    expect(() => createLayoutPage({
      ...base, sectionOccurrenceId: 'section:other',
    })).toThrow(RangeError);
    for (const mismatch of mismatches) {
      expect(() => createLayoutPage({ ...base, section: mismatch })).toThrow(RangeError);
    }
  });

  it('binds one retained occurrence translation in logical page space', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const region = {
      id: 'region:body', sectionOccurrenceId: 'section:body', section: bodySection,
      writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 720,
      columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
    };

    expect(bodyOccurrenceDestinationFor(2, region, 0, 180, rect(12, 30, 100, 20)))
      .toEqual({
        coordinateSpace: 'logical-page-points',
        flowDomainId: bodyFlowDomainId(2, 'region:body', 0),
        translation: { xPt: 60, yPt: 150 },
      });
  });

  it('rejects invalid occurrence destination inputs', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const region = {
      id: 'region:body', sectionOccurrenceId: 'section:body', section: bodySection,
      writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 720,
      columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
    };

    expect(() => bodyOccurrenceDestinationFor(-1, region, 0, 100, rect(0, 0, 10, 10)))
      .toThrow(RangeError);
    expect(() => bodyOccurrenceDestinationFor(0, region, 1, 100, rect(0, 0, 10, 10)))
      .toThrow(RangeError);
    expect(() => bodyOccurrenceDestinationFor(0, region, 0, Number.NaN, rect(0, 0, 10, 10)))
      .toThrow(RangeError);
    expect(() => bodyOccurrenceDestinationFor(0, region, 0, 100, rect(0, 0, -1, 10)))
      .toThrow(RangeError);
  });

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
    expect(page.flowDomains.map((domain) => domain.logicalBounds)).toEqual([
      rect(72, 72, 220, 258),
      rect(320, 72, 220, 258),
      rect(72, 330, 468, 390),
    ]);
    expect(page.flowDomains.map((domain) => domain.physicalBounds))
      .toEqual(page.flowDomains.map((domain) => domain.logicalBounds));
    expect(page.sectionOccurrenceId).toBe('section:first');
    expect(page.pageNumber).toEqual({
      displayNumber: 1,
      format: 'decimal',
      sectionOccurrenceId: 'section:first',
    });
  });

  it('treats a section boundary reconstructed with floating-point dust as shared', () => {
    // educational blind document: section 1 ends at 42.55 + 134.40000000000003
    // and the continuous two-column section 2 starts at 176.95000000000002.
    const first = section('lrTb', [{ xPt: 42.55, wPt: 510.2 }]);
    const second = section('lrTb', [
      { xPt: 42.55, wPt: 244.475 },
      { xPt: 308.275, wPt: 244.475 },
    ]);
    const build = (secondStartPt: number) => createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 612, heightPt: 792, contentTopPt: 42.55, contentBottomPt: 760 },
      sectionOccurrenceId: 'section:first',
      section: first,
      sectionRegions: [
        {
          id: 'region:first',
          sectionOccurrenceId: 'section:first',
          section: first,
          writingMode: 'horizontal-tb',
          blockStartPt: 42.55,
          blockEndPt: 176.95000000000005,
          columns: [{ inlineStartPt: 42.55, inlineExtentPt: 510.2 }],
        },
        {
          id: 'region:second',
          sectionOccurrenceId: 'section:second',
          section: second,
          writingMode: 'horizontal-tb',
          blockStartPt: secondStartPt,
          blockEndPt: 760,
          columns: [
            { inlineStartPt: 42.55, inlineExtentPt: 244.475 },
            { inlineStartPt: 308.275, inlineExtentPt: 244.475 },
          ],
        },
      ],
      paint: [],
      readingOrder: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:first' },
    });

    expect(build(176.95000000000002).flowDomains).toHaveLength(3);
    expect(() => build(176.9)).toThrow('Section flow domains on one page must be physically disjoint');
  });

  it('retains overlapping block bands when section regions own disjoint authored columns', () => {
    const columns = [
      { xPt: 72, wPt: 142 },
      { xPt: 235, wPt: 142 },
      { xPt: 398, wPt: 142 },
    ];
    const outgoing = section('lrTb', columns);
    const incoming = section('lrTb', columns);
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:outgoing',
      section: outgoing,
      sectionRegions: [
        {
          id: 'region:outgoing',
          sectionOccurrenceId: 'section:outgoing',
          section: outgoing,
          writingMode: 'horizontal-tb',
          blockStartPt: 240,
          blockEndPt: 720,
          columnIndexes: [0],
          columns: [{ inlineStartPt: 72, inlineExtentPt: 142 }],
        },
        {
          id: 'region:incoming',
          sectionOccurrenceId: 'section:incoming',
          section: incoming,
          writingMode: 'horizontal-tb',
          blockStartPt: 240,
          blockEndPt: 720,
          columnIndexes: [1, 2],
          columns: [
            { inlineStartPt: 235, inlineExtentPt: 142 },
            { inlineStartPt: 398, inlineExtentPt: 142 },
          ],
        },
      ],
      paint: [],
      readingOrder: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:outgoing',
      },
    });

    expect(page.sectionRegions.map((region) => region.columnIndexes)).toEqual([
      [0],
      [1, 2],
    ]);
    expect(page.flowDomains.map((domain) => domain.id)).toEqual([
      bodyFlowDomainId(0, 'region:outgoing', 0),
      bodyFlowDomainId(0, 'region:incoming', 1),
      bodyFlowDomainId(0, 'region:incoming', 2),
    ]);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    expect(() => createLayoutPage({
      ...page,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionRegions: [
        {
          id: 'region:outgoing',
          sectionOccurrenceId: 'section:outgoing',
          section: outgoing,
          writingMode: 'horizontal-tb',
          blockStartPt: 240,
          blockEndPt: 720,
          columnIndexes: [0],
          columns: [{ inlineStartPt: 72, inlineExtentPt: 142 }],
        },
        {
          id: 'region:incoming',
          sectionOccurrenceId: 'section:incoming',
          section: incoming,
          writingMode: 'horizontal-tb',
          blockStartPt: 240,
          blockEndPt: 720,
          columnIndexes: [0],
          columns: [{ inlineStartPt: 72, inlineExtentPt: 142 }],
        },
      ],
      paint: [],
      readingOrder: [],
    })).toThrow(RangeError);
  });

  it('requires retained column population direction to agree with sectPr bidi', () => {
    const rtlSection = {
      ...section('lrTb', [
        { xPt: 72, wPt: 224 },
        { xPt: 316, wPt: 224 },
      ]),
      sectionBidi: true,
    };

    expect(() => createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:rtl',
      section: rtlSection,
      sectionRegions: [{
        id: 'region:rtl',
        sectionOccurrenceId: 'section:rtl',
        section: rtlSection,
        writingMode: 'horizontal-tb',
        blockStartPt: 72,
        blockEndPt: 720,
        columnFlowDirection: 'ltr',
        columns: [
          { inlineStartPt: 72, inlineExtentPt: 224 },
          { inlineStartPt: 316, inlineExtentPt: 224 },
        ],
      }],
      paint: [],
      readingOrder: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:rtl',
      },
    })).toThrow(/column flow direction/i);
  });

  it('retains section column separator geometry on the completed page', () => {
    const separated = {
      ...section('lrTb', [
        { xPt: 72, wPt: 220 },
        { xPt: 320, wPt: 220 },
      ]),
      columnSeparator: true,
    };

    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: {
        widthPt: 612,
        heightPt: 792,
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:separated',
      section: separated,
      sectionRegions: [{
        id: 'region:separated',
        sectionOccurrenceId: 'section:separated',
        section: separated,
        writingMode: 'horizontal-tb',
        blockStartPt: 72,
        blockEndPt: 330,
        columns: [
          { inlineStartPt: 72, inlineExtentPt: 220 },
          { inlineStartPt: 320, inlineExtentPt: 220 },
        ],
      }],
      paint: [],
      readingOrder: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'section:separated',
      },
    });

    expect(page.columnSeparators).toEqual([{
      start: { xPt: 306, yPt: 72 },
      end: { xPt: 306, yPt: 330 },
    }]);
    expect(Object.isFrozen(page.columnSeparators)).toBe(true);
  });

  it('retains an empty continuous region only for out-of-flow occurrence ownership', () => {
    const first = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const second = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const firstDomain = bodyFlowDomainId(0, 'region:first', 0);
    const secondDomain = bodyFlowDomainId(0, 'region:second', 0);
    const floating = {
      ...drawing('floating', firstDomain, rect(72, 72, 100, 20)),
      ordinaryFlow: false,
    };
    const follower = drawing('follower', secondDomain, rect(72, 72, 100, 20));
    const input = {
      pageIndex: 0,
      physicalPage: {
        widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720,
      },
      sectionOccurrenceId: 'section:first',
      section: first,
      sectionRegions: [
        {
          id: 'region:first', sectionOccurrenceId: 'section:first', section: first,
          writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 72,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        },
        {
          id: 'region:second', sectionOccurrenceId: 'section:second', section: second,
          writingMode: 'horizontal-tb' as const, blockStartPt: 72, blockEndPt: 720,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        },
      ],
      paint: [
        { layer: 'body' as const, node: floating },
        { layer: 'body' as const, node: follower },
      ],
      readingOrder: [floating, follower],
      pageNumber: {
        displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:first',
      },
    };

    const page = createLayoutPage(input);

    expect(page.flowDomains.map((domain) => domain.logicalBounds.heightPt)).toEqual([0, 648]);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();
    expect(() => assertDocumentLayout({
      pages: [{
        ...page,
        layers: createPageLayers([
          { layer: 'body', node: { ...floating, ordinaryFlow: true } },
          { layer: 'body', node: follower },
        ]),
      }],
      diagnostics: [],
    })).toThrow(LayoutInvariantError);
  });

  it('retains a logical-to-physical transform for vertical section regions', () => {
    const vertical = verticalSection('tbRl', [{ xPt: 72, wPt: 648 }]);

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
      physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 612 },
    });
    expect(page.flowDomains[0]?.logicalBounds).toEqual(rect(72, 72, 648, 468));
    expect(page.flowDomains[0]?.physicalBounds).toEqual(rect(72, 72, 468, 648));
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
    const vertical = verticalSection('tbLrV', [{ xPt: 40, wPt: 700 }]);
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
    expect(page.sectionRegions?.[0]?.coordinateSpace.physicalToLogical)
      .toEqual({ a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 });
    expect(page.flowDomains[0]?.logicalBounds).toEqual(rect(40, 54, 700, 510));
    expect(page.flowDomains[0]?.physicalBounds).toEqual(rect(54, 40, 510, 700));
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
    expect(page.layers.roots).toEqual([
      { layer: 'behindText', node: behind, coordinateSpace: 'section-logical' },
      { layer: 'body', node: paragraph, coordinateSpace: 'section-logical' },
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

  it('derives bookmark metadata in retained paint order across layer storage', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const front = bookmarkParagraph(
      'paragraph:front', domainId, 'front-destination', rect(200, 100, 50, 12),
    );
    const body = bookmarkParagraph(
      'paragraph:body', domainId, 'body-destination', rect(72, 100, 50, 12),
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
      paint: [{ layer: 'front', node: front }, { layer: 'body', node: body }],
      readingOrder: [front, body],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });

    expect(page.bookmarkStarts?.map(({ name }) => name))
      .toEqual(['front-destination', 'body-destination']);
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

  it('rejects bookmark metadata assigned to the wrong retained section region', () => {
    const first = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const second = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const firstDomainId = bodyFlowDomainId(0, 'region:first', 0);
    const paragraph = bookmarkParagraph(
      'paragraph:first', firstDomainId, 'destination', rect(72, 100, 50, 12),
    );
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 612, heightPt: 792, contentTopPt: 72, contentBottomPt: 720 },
      sectionOccurrenceId: 'section:first', section: first,
      sectionRegions: [
        {
          id: 'region:first', sectionOccurrenceId: 'section:first', section: first,
          writingMode: 'horizontal-tb', blockStartPt: 72, blockEndPt: 300,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        },
        {
          id: 'region:second', sectionOccurrenceId: 'section:second', section: second,
          writingMode: 'horizontal-tb', blockStartPt: 300, blockEndPt: 720,
          columns: [{ inlineStartPt: 72, inlineExtentPt: 468 }],
        },
      ],
      paint: [{ layer: 'body', node: paragraph }], readingOrder: [paragraph],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:first' },
    });
    const wrongRegion: DocumentLayout = {
      pages: [{
        ...page,
        bookmarkStarts: [{ ...page.bookmarkStarts![0]!, sectionOccurrenceId: 'section:second' }],
      }],
      diagnostics: [],
    };

    expect(() => assertDocumentLayout(wrongRegion)).toThrow(/bookmark metadata/);
  });

  it('rejects bookmark metadata with no matching retained placement', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const paragraph = bookmarkParagraph(
      'paragraph:body', domainId, 'retained-destination', rect(72, 100, 50, 12),
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
      paint: [{ layer: 'body', node: paragraph }], readingOrder: [paragraph],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });
    const inventedPlacement: DocumentLayout = {
      pages: [{
        ...page,
        bookmarkStarts: [{
          name: 'invented-destination',
          nodeId: paragraph.id,
          sectionOccurrenceId: 'section:body',
        }],
      }],
      diagnostics: [],
    };

    expect(() => assertDocumentLayout(inventedPlacement)).toThrow(/bookmark metadata/);
  });

  it('rejects an ownerless derived bookmark while preserving canonical emptiness', () => {
    const bodySection = section('lrTb', [{ xPt: 72, wPt: 468 }]);
    const domainId = bodyFlowDomainId(0, 'region:body', 0);
    const paragraph = bookmarkParagraph(
      'paragraph:body', domainId, 'destination', rect(72, 100, 50, 12),
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
      paint: [{ layer: 'body', node: paragraph }], readingOrder: [paragraph],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:body' },
    });
    const ownerless: DocumentLayout = {
      pages: [{
        ...page,
        bookmarkStarts: [{
          name: 'destination', nodeId: paragraph.id, sectionOccurrenceId: '',
        }],
      }],
      diagnostics: [],
    };
    try {
      assertDocumentLayout(ownerless);
      throw new Error('expected ownerless bookmark metadata to fail');
    } catch (error) {
      expect(error).toBeInstanceOf(LayoutInvariantError);
      expect((error as LayoutInvariantError).code).toBe('INVALID_REFERENCE');
    }

    expect(() => assertDocumentLayout({
      pages: [{
        ...page,
        bookmarkStarts: [],
        layers: createPageLayers([]),
        readingOrder: [],
      }],
      diagnostics: [],
    })).not.toThrow();
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
    expect(page.layers.roots).toEqual([]);
    expect(page.readingOrder).toEqual([]);
    expect(page.bookmarkStarts).toEqual([]);
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    const invalidBlank: DocumentLayout = {
      pages: [{
        ...page,
        flowDomains: [{
          id: 'unexpected',
          kind: 'body',
          logicalBounds: rect(72, 72, 468, 648),
          physicalBounds: rect(72, 72, 468, 648),
        }],
      }],
      diagnostics: [],
    };
    expect(() => assertDocumentLayout(invalidBlank)).toThrow(/parity blank/);
  });
});
