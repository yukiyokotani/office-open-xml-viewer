import { describe, expect, it } from 'vitest';
import { layoutFlowBlocks } from './flow.js';
import { assertDocumentLayout, layoutFingerprint } from './invariants.js';
import type {
  BlockLayoutAlgorithms,
  DocumentLayout,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutServices,
  PageSectionRegion,
  PagePaintNode,
  SourceRef,
  TableEdgeInputs,
  TableLayoutInput,
} from './types.js';
import type { SectionLayoutContext } from '../layout-context.js';
import { createCanvasFontRoute } from '@silurus/ooxml-core';
import { canonicalLogicalToPhysical, mapAffineRect } from './affine.js';

const source = (index: number): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path: [index],
});

const rect = (xPt: number, yPt: number, widthPt: number, heightPt: number): LayoutRect => ({
  xPt,
  yPt,
  widthPt,
  heightPt,
});

const noTableBorders: TableEdgeInputs = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function tableInput(index: number): TableLayoutInput {
  return {
    kind: 'table', id: `table-input-${index}`, source: source(index),
    flowDomainId: 'body', ordinaryFlow: true,
    alignment: 'left', indentPt: 0, bidiVisual: false,
    columnWidthsPt: [], borders: noTableBorders, rows: [],
  };
}

function drawing(
  id: string,
  flowBounds: LayoutRect,
  options: Partial<Pick<DrawingLayout, 'inkBounds' | 'clipBounds' | 'ordinaryFlow' | 'flowDomainId'>> = {},
): DrawingLayout {
  return {
    kind: 'drawing',
    id,
    source: source(Number(id.replace(/\D/g, '')) || 0),
    flowBounds,
    inkBounds: options.inkBounds ?? flowBounds,
    ...(options.clipBounds ? { clipBounds: options.clipBounds } : {}),
    advancePt: flowBounds.heightPt,
    ordinaryFlow: options.ordinaryFlow ?? true,
    flowDomainId: options.flowDomainId ?? 'body',
    commands: [],
  };
}

const bodyDomain: FlowDomain = {
  id: 'body',
  kind: 'body',
  bounds: rect(72, 72, 468, 648),
};

function serviceStubs(): LayoutServices {
  return {
    text: {
      fingerprint: 'text',
      localMetrics: {},
      resolve: () => ({
        requestedFamily: 'sans-serif', resolvedFamily: 'sans-serif',
        route: createCanvasFontRoute('sans-serif', 'generic'),
        source: 'generic', weight: 400, style: 'normal', diagnostics: [], genericFamily: 'sans-serif',
      }),
      shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], graphemeBoundaries: [0], diagnostics: [] }),
    },
    images: {
      fingerprint: 'images',
      resolve: () => ({ widthPt: 0, heightPt: 0, mimeType: 'application/octet-stream' }),
    },
    math: {
      fingerprint: 'math',
      resolve: () => ({ resourceKey: 'math', widthEm: 0, ascentEm: 0, descentEm: 0, diagnostics: [] }),
    },
  };
}

function documentWith(
  nodes: readonly PagePaintNode[],
  diagnostics: DocumentLayout['diagnostics'] = [],
): DocumentLayout {
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        ...rect(0, 0, 612, 792),
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      flowDomains: [bodyDomain],
      section: {} as SectionLayoutContext,
      sectionOccurrenceId: 'section:0',
      parityBlank: false,
      bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0',
        coordinateSpace: {
          writingMode: 'horizontal-tb',
          logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        },
        blockStartPt: 72, blockEndPt: 720, flowDomainIds: ['body'],
        section: {} as SectionLayoutContext,
      }],
      layers: {
        paintOrder: nodes.map((node) => ({
          layer: 'body' as const,
          nodeId: node.id,
          coordinateSpace: 'logical-body-points' as const,
          logicalBlock: {
            blockStartPt: node.flowBounds.yPt,
            blockExtentPt: node.flowBounds.heightPt,
          },
        })),
        background: [],
        behindText: [],
        header: [],
        body: nodes,
        notes: [],
        front: [],
        footer: [],
      },
      readingOrder: nodes.map((node) => node.id),
    }],
    diagnostics,
  };
}

describe('assertDocumentLayout', () => {
  it('rejects overlapping ordinary flow allocations', () => {
    const layout = documentWith([
      drawing('n1', rect(72, 100, 200, 30)),
      drawing('n2', rect(72, 120, 200, 30)),
    ]);

    expect(() => assertDocumentLayout(layout)).toThrow(/FLOW_OVERLAP/);
  });

  it('rejects ordinary flow that enters the bottom margin', () => {
    const layout = documentWith([drawing('n1', rect(72, 710, 200, 20))]);

    expect(() => assertDocumentLayout(layout)).toThrow(/FLOW_DOMAIN_INVASION/);
  });

  it('allows floating overlap, negative-spacing ink, and clipped overhang', () => {
    const ordinary = drawing('n1', rect(72, 100, 200, 30), {
      inkBounds: rect(72, 92, 200, 38),
    });
    const floating = drawing('n2', rect(72, 110, 200, 30), {
      ordinaryFlow: false,
    });
    const clipped = drawing('n3', rect(72, 200, 200, 30), {
      inkBounds: rect(60, 190, 240, 600),
      clipBounds: rect(72, 200, 200, 30),
    });

    expect(() => assertDocumentLayout(documentWith([ordinary, floating, clipped]))).not.toThrow();
  });

  it('validates ordinary flow only against siblings in the same story container', () => {
    const body = drawing('body-1', rect(72, 690, 200, 20));
    const footer = {
      ...drawing('footer-1', rect(72, 738, 200, 20), { flowDomainId: 'footer:default' }),
      source: { story: 'footer' as const, storyInstance: 'default', path: [0] },
    };
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        flowDomains: [
          bodyDomain,
          { id: 'footer:default', kind: 'footer', bounds: rect(72, 730, 468, 40) },
        ],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [
            { layer: 'body', nodeId: body.id, coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 690, blockExtentPt: 20 } },
            { layer: 'footer', nodeId: footer.id, coordinateSpace: 'physical-page-points' },
          ],
          body: [body],
          footer: [footer],
        },
        readingOrder: [body.id, footer.id],
      }],
    };

    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('rejects overlap within one domain but permits the same geometry in independent cells', () => {
    const first = drawing('cell-1', rect(100, 100, 80, 20), { flowDomainId: 'cell:1' });
    const second = drawing('cell-2', rect(100, 100, 80, 20), { flowDomainId: 'cell:2' });
    const base = documentWith([first, second]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        flowDomains: [
          { id: 'cell:1', kind: 'tableCell', bounds: rect(90, 90, 100, 40) },
          { id: 'cell:2', kind: 'tableCell', bounds: rect(90, 90, 100, 40) },
        ],
        sectionRegions: [],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: base.pages[0]!.layers.paintOrder.map((entry) => ({
            ...entry,
            coordinateSpace: 'physical-page-points' as const,
          })),
        },
      }],
    };

    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('rejects missing domains and invalid paint or reading-order references', () => {
    const node = drawing('n1', rect(72, 100, 200, 30), { flowDomainId: 'missing' });
    expect(() => assertDocumentLayout(documentWith([node]))).toThrow(/INVALID_REFERENCE/);

    const base = documentWith([drawing('n1', rect(72, 100, 200, 30))]);
    const badPaint: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        layers: { ...base.pages[0]!.layers, paintOrder: [{ layer: 'body', nodeId: 'unknown', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 100, blockExtentPt: 30 } }] },
      }],
    };
    expect(() => assertDocumentLayout(badPaint)).toThrow(/INVALID_REFERENCE/);

    const badReading: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, readingOrder: ['unknown'] }],
    };
    expect(() => assertDocumentLayout(badReading)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects duplicate node IDs and duplicate paint entries', () => {
    expect(() => assertDocumentLayout(documentWith([
      drawing('n1', rect(72, 100, 200, 20)),
      drawing('n1', rect(72, 130, 200, 20)),
    ]))).toThrow(/INVALID_REFERENCE/);

    const base = documentWith([drawing('n1', rect(72, 100, 200, 20))]);
    const duplicatePaint: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [
            { layer: 'body', nodeId: 'n1', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 100, blockExtentPt: 20 } },
            { layer: 'body', nodeId: 'n1', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 100, blockExtentPt: 20 } },
          ],
        },
      }],
    };
    expect(() => assertDocumentLayout(duplicatePaint)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects non-finite retained geometry', () => {
    const layout = documentWith([drawing('n1', rect(Number.NaN, 100, 200, 30))]);

    expect(() => assertDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it('requires non-negative sequential page identity', () => {
    const base = documentWith([]);
    const negative: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, pageIndex: -1 }],
    };
    const skipped: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, pageIndex: 2 }],
    };
    const duplicate: DocumentLayout = {
      ...base,
      pages: [base.pages[0]!, { ...base.pages[0]! }],
    };

    expect(() => assertDocumentLayout(negative)).toThrow(/page index/);
    expect(() => assertDocumentLayout(skipped)).toThrow(/page index/);
    expect(() => assertDocumentLayout(duplicate)).toThrow(/page index/);
  });

  it('requires ordered effective page edges within the physical page and permits equality', () => {
    const base = documentWith([]);
    const withEdges = (contentTopPt: number, contentBottomPt: number): DocumentLayout => ({
      ...base,
      pages: [{
        ...base.pages[0]!,
        geometry: { ...base.pages[0]!.geometry, contentTopPt, contentBottomPt },
      }],
    });

    expect(() => assertDocumentLayout(withEdges(-1, 720))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(72, 793))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(721, 720))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(0, 0))).not.toThrow();
    expect(() => assertDocumentLayout(withEdges(792, 792))).not.toThrow();
  });

  it('requires every body flow domain to belong to exactly one page-local section region', () => {
    const base = documentWith([drawing('n1', rect(72, 100, 200, 30))]);
    const layout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        sectionRegions: [{
          id: 'section-region:0',
          sectionOccurrenceId: 'section:0',
          coordinateSpace: {
            writingMode: 'horizontal-tb' as const,
            logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          },
          blockStartPt: 72,
          blockEndPt: 720,
          flowDomainIds: [],
          section: base.pages[0]!.section,
        }],
      }],
    } as DocumentLayout;

    expect(() => assertDocumentLayout(layout)).toThrow(/section region ownership/);
  });

  it.each(['vertical-rl', 'vertical-lr'] as const)(
    'maps %s logical bounds once before physical domain containment',
    (writingMode) => {
      const matrix = canonicalLogicalToPhysical(writingMode, 641);
      const logical = rect(91, 307, 47, 23);
      const mapped = mapAffineRect(matrix, logical);
      const node = drawing(`mapped-${writingMode}`, logical);
      const base = documentWith([node]);
      const layout: DocumentLayout = {
        ...base,
        pages: [{
          ...base.pages[0]!,
          geometry: { ...base.pages[0]!.geometry, widthPt: 641, heightPt: 733 },
          flowDomains: [{ id: 'body', kind: 'body', bounds: mapped }],
          sectionRegions: [{
            ...base.pages[0]!.sectionRegions[0]!,
            coordinateSpace: { writingMode, logicalToPhysical: matrix },
            blockStartPt: 0,
            blockEndPt: 641,
          }],
        }],
      };

      expect(mapped).not.toEqual(logical);
      expect(() => assertDocumentLayout(layout)).not.toThrow();
      const overflow = structuredClone(layout);
      const overflowNode = overflow.pages[0]!.layers.body[0] as {
        flowBounds: { heightPt: number };
        inkBounds: { heightPt: number };
      };
      overflowNode.flowBounds.heightPt += 1;
      overflowNode.inkBounds.heightPt += 1;
      expect(() => assertDocumentLayout(overflow)).toThrow(/FLOW_DOMAIN_INVASION/);
    },
  );

  it('rejects missing, non-finite, wrong-direction, and wrong-page-width matrices', () => {
    const logical = rect(91, 307, 47, 23);
    const matrix = canonicalLogicalToPhysical('vertical-rl', 641);
    const node = drawing('matrix-node', logical);
    const base = documentWith([node]);
    const valid: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        geometry: { ...base.pages[0]!.geometry, widthPt: 641, heightPt: 733 },
        flowDomains: [{ id: 'body', kind: 'body', bounds: mapAffineRect(matrix, logical) }],
        sectionRegions: [{
          ...base.pages[0]!.sectionRegions[0]!,
          coordinateSpace: { writingMode: 'vertical-rl', logicalToPhysical: matrix },
          blockStartPt: 0, blockEndPt: 641,
        }],
      }],
    };
    expect(() => assertDocumentLayout(structuredClone(valid))).not.toThrow();
    expect(layoutFingerprint(structuredClone(valid))).toBe(layoutFingerprint(valid));

    const missing = structuredClone(valid) as unknown as { pages: Array<{ sectionRegions: Array<{ coordinateSpace?: unknown }> }> };
    delete missing.pages[0]!.sectionRegions[0]!.coordinateSpace;
    expect(() => assertDocumentLayout(missing as unknown as DocumentLayout)).toThrow(/coordinate space/);
    for (const bad of [
      { ...matrix, e: Number.NaN },
      canonicalLogicalToPhysical('vertical-lr', 641),
      canonicalLogicalToPhysical('vertical-rl', 640),
    ]) {
      const invalid = structuredClone(valid);
      (invalid.pages[0]!.sectionRegions[0]!.coordinateSpace as {
        logicalToPhysical: typeof bad;
      }).logicalToPhysical = bad;
      expect(() => assertDocumentLayout(invalid)).toThrow(/INVALID_GEOMETRY/);
    }
  });

  it('validates mixed horizontal and vertical regions independently', () => {
    const horizontal = drawing('mixed-horizontal', rect(20, 30, 40, 20), { flowDomainId: 'horizontal' });
    const vertical = drawing('mixed-vertical', rect(100, 300, 50, 20), { flowDomainId: 'vertical' });
    const verticalMatrix = canonicalLogicalToPhysical('vertical-rl', 641);
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        geometry: { ...base.pages[0]!.geometry, widthPt: 641, heightPt: 733 },
        flowDomains: [
          { id: 'horizontal', kind: 'body', bounds: horizontal.flowBounds },
          { id: 'vertical', kind: 'body', bounds: mapAffineRect(verticalMatrix, vertical.flowBounds) },
        ],
        sectionRegions: [
          { ...base.pages[0]!.sectionRegions[0]!, id: 'horizontal', flowDomainIds: ['horizontal'], blockStartPt: 0, blockEndPt: 733 },
          { ...base.pages[0]!.sectionRegions[0]!, id: 'vertical', flowDomainIds: ['vertical'], blockStartPt: 0, blockEndPt: 641, coordinateSpace: { writingMode: 'vertical-rl', logicalToPhysical: verticalMatrix } },
        ],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [
            { layer: 'body', nodeId: horizontal.id, coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 30, blockExtentPt: 20 } },
            { layer: 'body', nodeId: vertical.id, coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 300, blockExtentPt: 20 } },
          ],
          body: [horizontal, vertical],
        },
        readingOrder: [horizontal.id, vertical.id],
      }],
    };
    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('requires an explicit upright table block footprint and matching float coordinate space', () => {
    const matrix = canonicalLogicalToPhysical('vertical-rl', 641);
    const bounds = rect(371, 90, 50, 80);
    const table = {
      kind: 'table' as const, id: 'upright-table', source: source(9), flowDomainId: 'body',
      flowBounds: bounds, inkBounds: bounds, advancePt: 80, ordinaryFlow: true,
      columnWidthsPt: [50], rows: [], borders: [],
      paintReadyFloatingTables: {
        kind: 'resolved' as const,
        coordinateSpace: 'upright-physical-page-points' as const,
        unresolved: [], placements: [],
      },
    };
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        geometry: { ...base.pages[0]!.geometry, widthPt: 641, heightPt: 733 },
        flowDomains: [{ id: 'body', kind: 'body', bounds: rect(300, 60, 200, 300) }],
        sectionRegions: [{
          ...base.pages[0]!.sectionRegions[0]!,
          coordinateSpace: { writingMode: 'vertical-rl', logicalToPhysical: matrix },
          blockStartPt: 0, blockEndPt: 641,
        }],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [{
            layer: 'body', nodeId: table.id,
            coordinateSpace: 'upright-physical-page-points',
            logicalBlock: { blockStartPt: 220, blockExtentPt: 50 },
          }],
          body: [table],
        },
        readingOrder: [table.id],
      }],
    };
    expect(() => assertDocumentLayout(layout)).not.toThrow();
    expect(table.advancePt).toBe(80);
    expect(layout.pages[0]!.layers.paintOrder[0]!.logicalBlock?.blockExtentPt).toBe(50);

    const missingFootprint = structuredClone(layout);
    delete (missingFootprint.pages[0]!.layers.paintOrder[0] as { logicalBlock?: unknown }).logicalBlock;
    expect(() => assertDocumentLayout(missingFootprint)).toThrow(/logical block footprint/);
    const wrongFloatSpace = structuredClone(layout);
    const wrongTable = wrongFloatSpace.pages[0]!.layers.body[0] as unknown as {
      paintReadyFloatingTables: {
        coordinateSpace: 'logical-page-points' | 'upright-physical-page-points';
      };
    };
    wrongTable.paintReadyFloatingTables.coordinateSpace = 'logical-page-points';
    expect(() => assertDocumentLayout(wrongFloatSpace)).toThrow(/mismatched floating-table/);
    const doubleMapped = structuredClone(layout);
    (doubleMapped.pages[0]!.layers.paintOrder[0] as {
      coordinateSpace: 'logical-body-points' | 'upright-physical-page-points';
    }).coordinateSpace = 'logical-body-points';
    (doubleMapped.pages[0]!.layers.body[0] as unknown as {
      paintReadyFloatingTables: {
        coordinateSpace: 'logical-page-points' | 'upright-physical-page-points';
      };
    }).paintReadyFloatingTables.coordinateSpace
      = 'logical-page-points';
    expect(() => assertDocumentLayout(doubleMapped)).toThrow(/FLOW_DOMAIN_INVASION/);

    const unsupported = structuredClone(layout);
    const unsupportedRegions = unsupported.pages[0]!.sectionRegions as PageSectionRegion[];
    unsupportedRegions[0] = {
      ...unsupported.pages[0]!.sectionRegions[0]!,
      coordinateSpace: {
        writingMode: 'vertical-lr',
        logicalToPhysical: canonicalLogicalToPhysical('vertical-lr', 641),
      },
    };
    expect(() => assertDocumentLayout(unsupported)).toThrow(/UNSUPPORTED_FEATURE/);
  });

  it('validates every paint-ready floating-table destination edge', () => {
    const rootBounds = rect(300, 60, 200, 300);
    const childBounds = rect(340, 120, 40, 30);
    const child = {
      kind: 'table' as const, id: 'nested-table', source: source(11), flowDomainId: 'body',
      flowBounds: childBounds, inkBounds: childBounds, advancePt: 30, ordinaryFlow: false,
      columnWidthsPt: [40], rows: [], borders: [],
    };
    const sourcePlacement = {
      kind: 'floating-table-placement' as const,
      occurrenceId: 'float:11', ownership: 'source' as const,
      physicalPageIndex: 0, displayPageNumber: 7,
      hostCellId: 'root:cell', sourceBlockIndex: 0, anchorBlockIndex: 0,
      tableId: child.id, overlap: 'never' as const, positioning: {} as never,
      anchorBounds: childBounds, child,
    };
    const placement = {
      kind: 'resolved-floating-table-placement' as const,
      occurrenceId: sourcePlacement.occurrenceId,
      xPt: childBounds.xPt, yPt: childBounds.yPt,
      bounds: childBounds, exclusionBounds: rect(338, 118, 44, 34),
      overlap: 'never' as const, child, source: sourcePlacement,
    };
    const root = {
      kind: 'table' as const, id: 'root-table', source: source(10), flowDomainId: 'body',
      flowBounds: rect(320, 90, 100, 80), inkBounds: rect(320, 90, 100, 80),
      advancePt: 80, ordinaryFlow: true, columnWidthsPt: [100], borders: [],
      rows: [{
        kind: 'table-row' as const, id: 'root:row', source: source(10), flowDomainId: 'body',
        flowBounds: rect(320, 90, 100, 80), inkBounds: rect(320, 90, 100, 80),
        advancePt: 80, ordinaryFlow: true, heightPt: 80, contentHeightPt: 80,
        cells: [{
          kind: 'table-cell' as const, id: 'root:cell', source: source(10), flowDomainId: 'body',
          flowBounds: rect(320, 90, 100, 80), inkBounds: rect(320, 90, 100, 80),
          contentBounds: rect(322, 92, 96, 76), advancePt: 80, ordinaryFlow: true,
          verticalMerge: 'none' as const, vAlign: 'top' as const, blocks: [],
        }],
      }],
      paintReadyFloatingTables: {
        kind: 'resolved' as const,
        coordinateSpace: 'upright-physical-page-points' as const,
        unresolved: [], placements: [placement],
      },
    };
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!, pageIndex: 0,
        geometry: { ...base.pages[0]!.geometry, widthPt: 641, heightPt: 733 },
        pageNumber: { displayNumber: 7, format: 'decimal', sectionOccurrenceId: 'section:0' },
        flowDomains: [{ id: 'body', kind: 'body', bounds: rootBounds }],
        sectionRegions: [{
          ...base.pages[0]!.sectionRegions[0]!,
          coordinateSpace: {
            writingMode: 'vertical-rl',
            logicalToPhysical: canonicalLogicalToPhysical('vertical-rl', 641),
          },
          blockStartPt: 0, blockEndPt: 641,
        }],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [{
            layer: 'body', nodeId: root.id,
            coordinateSpace: 'upright-physical-page-points',
            logicalBlock: { blockStartPt: 220, blockExtentPt: 100 },
          }],
          body: [root],
        },
        readingOrder: [root.id],
      }],
    };
    expect(() => assertDocumentLayout(layout)).not.toThrow();

    const mutate = (change: (copy: any) => void, error: RegExp): void => {
      const copy = structuredClone(layout);
      change(copy);
      expect(() => assertDocumentLayout(copy)).toThrow(error);
    };
    mutate((copy) => { delete copy.pages[0].layers.body[0].paintReadyFloatingTables; }, /paint-ready floating-table ownership/);
    mutate((copy) => { delete copy.pages[0].layers.body[0].paintReadyFloatingTables.coordinateSpace; }, /coordinate space/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].source.physicalPageIndex = 2; }, /destination ownership/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].source.displayPageNumber = 6; }, /destination ownership/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].child.flowDomainId = 'other'; }, /destination ownership/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].occurrenceId = 'other'; }, /destination ownership/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].bounds.xPt = 100; }, /destination ownership|destination domain/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].exclusionBounds.xPt = 100; }, /destination domain/);
    mutate((copy) => { copy.pages[0].layers.body[0].paintReadyFloatingTables.placements[0].source.hostCellId = 'missing'; }, /destination ownership/);
    mutate((copy) => {
      copy.pages[0].layers.paintOrder[0].coordinateSpace = 'logical-body-points';
    }, /mismatched floating-table coordinate space/);
  });
});

describe('layoutFingerprint', () => {
  it('normalizes geometry and excludes diagnostic prose while retaining diagnostic identity', () => {
    const first = documentWith(
      [drawing('n1', rect(72.0000001, 100, 200, 30))],
      [{ code: 'UNSUPPORTED_FEATURE', severity: 'warning', message: 'first prose' }],
    );
    const second = documentWith(
      [drawing('n1', rect(72.0000002, 100, 200, 30))],
      [{ code: 'UNSUPPORTED_FEATURE', severity: 'warning', message: 'different prose' }],
    );
    const changedCode = documentWith(
      [drawing('n1', rect(72.0000002, 100, 200, 30))],
      [{ code: 'NON_CONVERGENCE', severity: 'warning', message: 'different prose' }],
    );

    expect(layoutFingerprint(first)).toBe(layoutFingerprint(second));
    expect(layoutFingerprint(first)).not.toBe(layoutFingerprint(changedCode));
  });
});

describe('layoutFlowBlocks', () => {
  it('dispatches paragraph and table blocks through one injected coordinator', () => {
    const calls: string[] = [];
    const algorithms: BlockLayoutAlgorithms = {
      layoutParagraph(input, placement) {
        calls.push(`paragraph:${input.source.path.join('.')}:${placement.cursor.yPt}`);
        const layout = {
          ...drawing('p1', rect(10, placement.cursor.yPt, 100, 12)),
          kind: 'paragraph' as const,
          spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
          lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 12 } };
      },
      layoutTable(input, placement) {
        calls.push(`table:${input.source.path.join('.')}:${placement.cursor.yPt}`);
        const layout = {
          ...drawing('t2', rect(10, placement.cursor.yPt, 100, 18)),
          kind: 'table' as const,
          columnWidthsPt: [], rows: [], borders: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 18 } };
      },
    };
    const services = serviceStubs();

    const result = layoutFlowBlocks({
      source: source(0),
      container: { id: 'body', kind: 'body', bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
      blocks: [
        { kind: 'paragraph', source: source(1) },
        tableInput(2),
      ],
    }, services, algorithms);

    expect(calls).toEqual(['paragraph:1:20', 'table:2:32']);
    expect(result.blocks.map((block) => block.id)).toEqual(['p1', 't2']);
    expect(result.advancePt).toBe(30);
    expect(result.flowBounds).toEqual(rect(10, 20, 100, 30));
    expect(result.nextCursor).toEqual({ xPt: 10, yPt: 50 });
  });

  it('rejects a block assigned outside its enclosing flow domain', () => {
    const algorithms: BlockLayoutAlgorithms = {
      layoutParagraph(_input, placement) {
        const layout = {
          ...drawing('p1', rect(10, placement.cursor.yPt, 100, 12), { flowDomainId: 'other' }),
          kind: 'paragraph' as const,
          spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
          lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 12 } };
      },
      layoutTable() {
        throw new Error('not used');
      },
    };
    const services = serviceStubs();

    expect(() => layoutFlowBlocks({
      source: source(0),
      container: { id: 'cell:1', kind: 'tableCell', bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
      blocks: [{ kind: 'paragraph', source: source(1) }],
    }, services, algorithms)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects invalid containers and initial cursors before dispatch', () => {
    const unused: BlockLayoutAlgorithms = {
      layoutParagraph() { throw new Error('not used'); },
      layoutTable() { throw new Error('not used'); },
    };
    const services = serviceStubs();
    const base = {
      source: source(0),
      blocks: [],
      container: { id: 'body', kind: 'body' as const, bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
    };

    expect(() => layoutFlowBlocks({
      ...base,
      container: { ...base.container, bounds: rect(10, 20, Number.NaN, 200) },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFlowBlocks({
      ...base,
      cursor: { xPt: 10, yPt: 19 },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFlowBlocks({
      ...base,
      cursor: { xPt: 111, yPt: 20 },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
  });
});
