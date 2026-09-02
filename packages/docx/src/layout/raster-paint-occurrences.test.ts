import { describe, expect, it } from 'vitest';
import { buildPageLayers } from './page-layers.js';
import { rasterPaintOccurrencesForPage } from './text-index.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutPage,
  LayoutRect,
  Matrix2DData,
  PageLayers,
  ParagraphLayout,
  ParagraphPlacement,
  SourceRef,
  TableLayout,
  TextBoxLayout,
} from './types.js';

const identity = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

const verticalRlToPhysical = Object.freeze({
  a: 0, b: 1, c: -1, d: 0, e: 200, f: 0,
}) satisfies Matrix2DData;

const physicalToVerticalRl = Object.freeze({
  a: 0, b: -1, c: 1, d: 0, e: 0, f: 200,
}) satisfies Matrix2DData;

function rect(xPt = 0, yPt = 0, widthPt = 100, heightPt = 20): LayoutRect {
  return Object.freeze({ xPt, yPt, widthPt, heightPt });
}

function source(path: readonly number[]): SourceRef {
  return Object.freeze({
    story: 'body',
    storyInstance: 'body',
    path: Object.freeze([...path]),
  });
}

function paragraph(
  id: string,
  placements: readonly ParagraphPlacement[],
  options: Readonly<{
    drawings?: readonly DrawingLayout[];
    resources?: ParagraphLayout['resources'];
    textBoxes?: readonly TextBoxLayout[];
    clipBounds?: LayoutRect;
  }> = {},
): ParagraphLayout {
  const bounds = rect();
  return Object.freeze({
    kind: 'paragraph',
    id,
    source: source([0]),
    flowDomainId: 'body:domain',
    flowBounds: bounds,
    inkBounds: bounds,
    ...(options.clipBounds ? { clipBounds: options.clipBounds } : {}),
    advancePt: 20,
    ordinaryFlow: true,
    spacing: Object.freeze({ beforePt: 0, afterPt: 0 }),
    contextualSpacing: false,
    lines: Object.freeze([Object.freeze({
      range: Object.freeze({ start: 0, end: placements.length }),
      bounds,
      baselinePt: 10,
      advancePt: 20,
      placements: Object.freeze([...placements]),
    })]),
    borders: Object.freeze([]),
    resources: Object.freeze([...(options.resources ?? [])]),
    drawings: Object.freeze([...(options.drawings ?? [])]),
    textBoxes: Object.freeze([...(options.textBoxes ?? [])]),
    events: Object.freeze([]),
    exclusions: Object.freeze([]),
  });
}

function tableWithParagraph(child: ParagraphLayout): TableLayout {
  const tableBounds = rect(10, 20, 100, 40);
  return Object.freeze({
    kind: 'table',
    id: 'table',
    source: source([1]),
    flowDomainId: 'body:domain',
    flowBounds: tableBounds,
    inkBounds: tableBounds,
    clipBounds: rect(10, 20, 2, 2),
    advancePt: 40,
    ordinaryFlow: true,
    columnWidthsPt: Object.freeze([100]),
    borders: Object.freeze([]),
    rows: Object.freeze([Object.freeze({
      kind: 'table-row',
      id: 'table:row',
      source: source([1, 0]),
      flowDomainId: 'body:domain',
      flowBounds: tableBounds,
      inkBounds: tableBounds,
      advancePt: 40,
      ordinaryFlow: true,
      heightPt: 40,
      contentHeightPt: 20,
      cells: Object.freeze([Object.freeze({
        kind: 'table-cell',
        id: 'table:cell',
        source: source([1, 0, 0]),
        flowDomainId: 'body:domain',
        flowBounds: tableBounds,
        inkBounds: tableBounds,
        clipBounds: rect(11, 21, 1, 1),
        advancePt: 40,
        ordinaryFlow: true,
        contentBounds: rect(15, 20, 90, 40),
        verticalMerge: 'none',
        vAlign: 'top',
        blocks: Object.freeze([Object.freeze({
          layout: child,
          offsetPt: 7,
          advancePt: child.advancePt,
        })]),
      })]),
    })]),
  });
}

function page(
  layers: PageLayers,
  rootIds: readonly string[],
  coordinateSpace: Readonly<{
    writingMode: 'horizontal-tb' | 'vertical-rl';
    logicalToPhysical: Matrix2DData;
    physicalToLogical: Matrix2DData;
  }> = {
    writingMode: 'horizontal-tb',
    logicalToPhysical: identity,
    physicalToLogical: identity,
  },
): LayoutPage {
  const geometry = Object.freeze({
    ...rect(0, 0, 200, 300),
    contentTopPt: 0,
    contentBottomPt: 300,
  });
  return Object.freeze({
    pageIndex: 0,
    geometry,
    flowDomains: Object.freeze([Object.freeze({
      id: 'body:domain',
      kind: 'body' as const,
      sectionRegionId: 'region',
      logicalBounds: geometry,
      physicalBounds: geometry,
    })]),
    section: {} as LayoutPage['section'],
    sectionOccurrenceId: 'section',
    parityBlank: false,
    bookmarkStarts: Object.freeze([]),
    pageNumber: Object.freeze({
      displayNumber: 1,
      format: 'decimal',
      sectionOccurrenceId: 'section',
    }),
    sectionRegions: Object.freeze([Object.freeze({
      id: 'region',
      sectionOccurrenceId: 'section',
      coordinateSpace,
      blockStartPt: 0,
      blockEndPt: 300,
      columnFlowDirection: 'ltr' as const,
      columnIndexes: Object.freeze([0]),
      flowDomainIds: Object.freeze(['body:domain']),
      section: {} as LayoutPage['section'],
    })]),
    columnSeparators: Object.freeze([]),
    pageBorder: null,
    layers,
    readingOrder: Object.freeze([...rootIds]),
  });
}

function documentLayout(layoutPage: LayoutPage): DocumentLayout {
  return Object.freeze({
    pages: Object.freeze([layoutPage]),
    diagnostics: Object.freeze([]),
  });
}

function imageFillPlan(
  rectValue: Readonly<{ x: number; y: number; w: number; h: number }>,
  rotationDeg = 0,
) {
  return Object.freeze({
    rect: Object.freeze(rectValue),
    geometry: Object.freeze({
      kind: 'preset' as const,
      name: 'rect',
      adjustments: Object.freeze([]),
    }),
    fill: null,
    stroke: null,
    transform: Object.freeze({ rotationDeg, flipH: true, flipV: false }),
  });
}

describe('rasterPaintOccurrencesForPage', () => {
  it('uses the retained grouped-child command size through anchored table frames', () => {
    const resourceKey = 'image:grouped-anchor';
    const finalChildFrame = rect(20, 10, 8, 9);
    const drawing: DrawingLayout = Object.freeze({
      kind: 'drawing',
      id: 'grouped-anchor',
      source: source([0, 1]),
      flowDomainId: 'body:domain',
      flowBounds: finalChildFrame,
      inkBounds: finalChildFrame,
      advancePt: 0,
      ordinaryFlow: false,
      commands: Object.freeze([Object.freeze({
        kind: 'resource',
        resourceKey,
        resourceKind: 'image',
        rect: finalChildFrame,
      })]),
      anchorLayer: Object.freeze({
        occurrenceId: 'anchor:grouped',
        behindDoc: false,
        relativeHeight: 1,
        sourceOrder: 1,
        horizontalOwnership: 'host',
        verticalOwnership: 'host',
      }),
    });
    const owner = paragraph('owner', [], {
      drawings: [drawing],
      // This is the pre-group resource size. Decode demand must instead use
      // the completed child command above (8 x 9 pt).
      resources: [Object.freeze({
        kind: 'image',
        resourceKey,
        intrinsicSize: Object.freeze({ widthPt: 4, heightPt: 3 }),
      })],
      clipBounds: rect(0, 0, 1, 1),
    });
    const root = tableWithParagraph(owner);
    const layers = buildPageLayers([{ layer: 'body', node: root }]);
    const anchorEntry = layers.paintOrder.find((entry) => entry.kind === 'drawing');

    expect(anchorEntry?.kind === 'drawing' ? anchorEntry.frames : []).toContainEqual({
      kind: 'transform',
      transform: { a: 1, b: 0, c: 0, d: 1, e: 15, f: 27 },
    });
    expect(rasterPaintOccurrencesForPage(
      documentLayout(page(layers, [root.id])),
      0,
    )).toEqual([{
      resourceKey,
      resourceKind: 'image',
      widthPt: 8,
      heightPt: 9,
    }]);
  });

  it('uses the actual DrawingML image-fill destination including fillRect', () => {
    const drawing: DrawingLayout = Object.freeze({
      kind: 'drawing',
      id: 'shape-image-fills',
      source: source([0]),
      flowDomainId: 'body:domain',
      flowBounds: rect(0, 0, 100, 60),
      inkBounds: rect(0, 0, 100, 60),
      advancePt: 0,
      ordinaryFlow: false,
      commands: Object.freeze([
        Object.freeze({
          kind: 'drawingml-image-fill',
          resourceKey: 'image:cropped-fill',
          fillRect: Object.freeze({ l: 0.25, t: 0.125, r: -0.125, b: 0.25 }),
          plan: imageFillPlan({ x: 10, y: 5, w: 80, h: 48 }, 37),
        }),
        Object.freeze({
          kind: 'drawingml-image-fill',
          resourceKey: 'image:whole-fill',
          plan: imageFillPlan({ x: -10, y: 2, w: 30, h: 40 }),
        }),
      ]),
    });
    const layers = buildPageLayers([{ layer: 'body', node: drawing }]);

    expect(rasterPaintOccurrencesForPage(
      documentLayout(page(layers, [drawing.id])),
      0,
    )).toEqual([
      {
        resourceKey: 'image:cropped-fill',
        resourceKind: 'image',
        widthPt: 70,
        heightPt: 30,
      },
      {
        resourceKey: 'image:whole-fill',
        resourceKind: 'image',
        widthPt: 30,
        heightPt: 40,
      },
    ]);
  });

  it('accounts for the local counter-rotation of upright inline resources', () => {
    const bounds = rect(10, 20, 12, 5);
    const body = paragraph('vertical-body', [
      Object.freeze({
        kind: 'resource',
        range: Object.freeze({ start: 0, end: 1 }),
        resourceKey: 'image:flow-relative',
        resourceKind: 'image',
        bounds,
        advancePt: 12,
      }),
      Object.freeze({
        kind: 'resource',
        range: Object.freeze({ start: 1, end: 2 }),
        resourceKey: 'image:upright',
        resourceKind: 'image',
        orientation: 'upright-physical',
        bounds,
        advancePt: 12,
      }),
    ]);
    const layers = buildPageLayers([{ layer: 'body', node: body }]);

    expect(rasterPaintOccurrencesForPage(
      documentLayout(page(layers, [body.id], {
        writingMode: 'vertical-rl',
        logicalToPhysical: verticalRlToPhysical,
        physicalToLogical: physicalToVerticalRl,
      })),
      0,
    )).toEqual([
      {
        resourceKey: 'image:flow-relative',
        resourceKind: 'image',
        widthPt: 12,
        heightPt: 5,
      },
      {
        resourceKey: 'image:upright',
        resourceKind: 'image',
        widthPt: 5,
        heightPt: 12,
      },
    ]);
  });

  it.each(['vert', 'vert270'] as const)(
    'keeps an inline image in a nested %s text-box table at its physical size',
    (verticalMode) => {
      const resourceKey = `image:textbox:${verticalMode}`;
      // Vertical text-box acquisition stores the 24 x 36 physical image in a
      // 36 x 24 logical frame. Paint counter-rotates and swaps that frame after
      // entering the text box's own quarter-turn transform.
      const logicalBounds = rect(4, 6, 36, 24);
      const imageParagraph = paragraph(`textbox-image:${verticalMode}`, [Object.freeze({
        kind: 'resource',
        range: Object.freeze({ start: 0, end: 1 }),
        resourceKey,
        resourceKind: 'image',
        bounds: logicalBounds,
        advancePt: 36,
      })]);
      const nestedTable = tableWithParagraph(imageParagraph);
      const textBoxId = `vertical-textbox:${verticalMode}`;
      const textBoxBounds = rect(0, 0, 120, 80);
      const textBox: TextBoxLayout = Object.freeze({
        kind: 'textbox',
        id: textBoxId,
        source: source([0, 1, 0]),
        flowDomainId: 'body:domain:textbox',
        flowBounds: textBoxBounds,
        inkBounds: textBoxBounds,
        advancePt: 0,
        ordinaryFlow: false,
        story: Object.freeze({
          story: 'textbox',
          flowBounds: nestedTable.flowBounds,
          inkBounds: nestedTable.inkBounds,
          blocks: Object.freeze([nestedTable]),
          advancePt: nestedTable.advancePt,
          diagnostics: Object.freeze([]),
        }),
        transform: verticalMode === 'vert270'
          ? Object.freeze({ a: 0, b: -1, c: 1, d: 0, e: 60, f: 40 })
          : Object.freeze({ a: 0, b: 1, c: -1, d: 0, e: 60, f: 40 }),
        writingMode: verticalMode === 'vert270' ? 'vertical-lr' : 'vertical-rl',
        verticalMode,
        insets: Object.freeze({ topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 }),
      });
      const drawing: DrawingLayout = Object.freeze({
        kind: 'drawing',
        id: `textbox-owner:${verticalMode}`,
        source: source([0, 1]),
        flowDomainId: 'body:domain',
        flowBounds: textBoxBounds,
        inkBounds: textBoxBounds,
        advancePt: 0,
        ordinaryFlow: false,
        commands: Object.freeze([Object.freeze({ kind: 'noop' })]),
        textBoxIds: Object.freeze([textBoxId]),
      });
      const owner = paragraph(`textbox-owner-paragraph:${verticalMode}`, [], {
        drawings: [drawing],
        textBoxes: [textBox],
      });
      const layers = buildPageLayers([{ layer: 'body', node: owner }]);

      expect(rasterPaintOccurrencesForPage(
        documentLayout(page(layers, [owner.id])),
        0,
      )).toEqual([{
        resourceKey,
        resourceKind: 'image',
        widthPt: 24,
        heightPt: 36,
      }]);
    },
  );
});
