import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { hitTestDocxElementContext } from './element-context.js';
import { buildPageLayers } from './layout/page-layers.js';
import { createPaintResourceRegistry } from './layout/paint-resources.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutRect,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  TextBoxLayout,
} from './layout/types.js';

const bounds = Object.freeze({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 60 }) satisfies LayoutRect;

function source(path: readonly number[]): SourceRef {
  return { story: 'body', storyInstance: 'body', path };
}

function drawing(
  id: string,
  sourceOrder: number,
  resourceKind: 'chart' | 'image',
  resourceKey: string,
  relativeHeight: number,
): DrawingLayout {
  return {
    kind: 'drawing',
    id,
    source: source([0, sourceOrder]),
    flowDomainId: 'body',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 0,
    ordinaryFlow: false,
    commands: [{ kind: 'resource', resourceKey, resourceKind, rect: bounds }],
    anchorLayer: {
      occurrenceId: `anchor:${id}`,
      behindDoc: false,
      relativeHeight,
      sourceOrder,
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    },
  };
}

function paragraph(drawings: readonly DrawingLayout[]): ParagraphLayout {
  return {
    kind: 'paragraph',
    id: 'paragraph',
    source: source([0]),
    flowDomainId: 'body',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 20,
    ordinaryFlow: true,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [],
    borders: [],
    resources: [],
    drawings,
    textBoxes: [],
    events: [],
    exclusions: [],
  };
}

function inlineImageParagraph(
  id: string,
  paragraphSource: SourceRef,
  resourceKey: string,
  resourceBounds: LayoutRect,
  sourceRunIndex: number,
  clipBounds?: LayoutRect,
): ParagraphLayout {
  return {
    ...paragraph([]),
    id,
    source: paragraphSource,
    flowBounds: resourceBounds,
    inkBounds: resourceBounds,
    ...(clipBounds ? { clipBounds } : {}),
    lines: [{
      range: { start: 0, end: 1 },
      bounds: resourceBounds,
      baselinePt: resourceBounds.yPt + resourceBounds.heightPt,
      advancePt: resourceBounds.heightPt,
      placements: [{
        kind: 'resource',
        range: { start: 0, end: 1 },
        sourceRunIndex,
        resourceKey,
        resourceKind: 'image',
        bounds: resourceBounds,
        advancePt: resourceBounds.widthPt,
      }],
    }],
  };
}

function textBox(
  id: string,
  child: ParagraphLayout,
  transform = { a: 1, b: 0, c: 0, d: 1, e: 5, f: 7 },
  clipBounds?: LayoutRect,
): TextBoxLayout {
  return {
    kind: 'textbox',
    id,
    source: { story: 'textbox', storyInstance: id, path: [] },
    flowDomainId: id,
    flowBounds: child.flowBounds,
    inkBounds: child.inkBounds,
    ...(clipBounds ? { clipBounds } : {}),
    advancePt: child.advancePt,
    ordinaryFlow: false,
    story: {
      story: 'textbox',
      flowBounds: child.flowBounds,
      inkBounds: child.inkBounds,
      blocks: [child],
      advancePt: child.advancePt,
      diagnostics: [],
    },
    transform,
    writingMode: 'horizontal-tb',
    insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
  };
}

function layoutFor(drawings: readonly DrawingLayout[]): DocumentLayout {
  const root = paragraph(drawings);
  const layers = buildPageLayers([{ layer: 'body', node: root }]);
  return {
    pages: [{
      pageIndex: 0,
      geometry: { xPt: 0, yPt: 0, widthPt: 600, heightPt: 800, contentTopPt: 0, contentBottomPt: 800 },
      flowDomains: [],
      section: {} as never,
      sectionOccurrenceId: 'section',
      parityBlank: false,
      bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section' },
      sectionRegions: [],
      columnSeparators: [],
      pageBorder: null,
      layers,
      readingOrder: [root.id],
    }],
    diagnostics: [],
  };
}

const chart = {
  chartType: 'bar',
  title: 'Quarterly revenue',
  categories: ['Q1', 'Q2'],
  series: [{ name: 'Revenue', values: [10, 20] }],
} as unknown as ChartModel;

describe('DOCX element context', () => {
  it('uses retained page paint order and returns bounded chart context', () => {
    const image = drawing('image', 0, 'image', 'image:1', 1);
    const topChart = drawing('chart', 1, 'chart', 'chart:1', 2);
    const registry = createPaintResourceRegistry([
      {
        kind: 'image', resourceKey: 'image:1', partPath: 'word/media/image1.png',
        mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
      },
      { kind: 'chart', resourceKey: 'chart:1', intrinsicSize: { widthPt: 100, heightPt: 60 }, model: chart },
    ]);

    const context = hitTestDocxElementContext(
      layoutFor([image, topChart]), 0, { xPt: 20, yPt: 30 }, registry,
    );

    expect(context).toMatchObject({
      format: 'docx',
      kind: 'element',
      pageIndex: 0,
      elementType: 'chart',
      source: { story: 'body', storyInstance: 'body', path: [0, 1] },
      seriesCount: 1,
      text: expect.stringContaining('Quarterly revenue'),
    });
  });

  it('does no work for points outside retained drawing ink', () => {
    const image = drawing('image', 0, 'image', 'image:1', 1);
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: 'image:1', partPath: 'word/media/image1.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
    }]);
    expect(hitTestDocxElementContext(
      layoutFor([image]), 0, { xPt: 500, yPt: 500 }, registry,
    )).toBeNull();
  });

  it('hit-tests an inline image retained as a line resource placement', () => {
    const layout = layoutFor([]);
    const root = layout.pages[0].layers.roots[0].node as ParagraphLayout;
    const inlineBounds = { xPt: 30, yPt: 40, widthPt: 120, heightPt: 80 };
    Object.assign(root, {
      lines: [{
        range: { start: 0, end: 1 },
        bounds: inlineBounds,
        baselinePt: 120,
        advancePt: 80,
        placements: [{
          kind: 'resource',
          range: { start: 0, end: 1 },
          sourceRunIndex: 3,
          resourceKey: 'image:inline',
          resourceKind: 'image',
          bounds: inlineBounds,
          advancePt: 120,
        }],
      }],
    });
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: 'image:inline', partPath: 'word/media/inline.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 120, heightPt: 80 },
    }]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 90, yPt: 80 }, registry,
    )).toMatchObject({
      format: 'docx',
      kind: 'element',
      elementType: 'image',
      bounds: inlineBounds,
      source: { story: 'body', storyInstance: 'body', path: [0, 3] },
      mimeType: 'image/png',
    });
  });

  it('hit-tests inline drawings above the line resources they are painted over', () => {
    const retained = drawing('retained-inline', 0, 'image', 'image:retained', 0);
    const { anchorLayer: _anchorLayer, ...inlineDrawing } = retained;
    const layout = layoutFor([inlineDrawing]);
    const root = layout.pages[0].layers.roots[0].node as ParagraphLayout;
    Object.assign(root, {
      lines: [{
        range: { start: 0, end: 2 },
        bounds,
        baselinePt: 80,
        advancePt: 60,
        placements: [{
          kind: 'resource',
          range: { start: 1, end: 2 },
          sourceRunIndex: 1,
          resourceKey: 'image:placement',
          resourceKind: 'image',
          bounds,
          advancePt: 100,
        }],
      }],
    });
    const registry = createPaintResourceRegistry([
      {
        kind: 'image', resourceKey: 'image:retained', partPath: 'word/media/retained.png',
        mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
      },
      {
        kind: 'image', resourceKey: 'image:placement', partPath: 'word/media/placement.png',
        mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
      },
    ]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 20, yPt: 30 }, registry,
    )).toMatchObject({
      elementType: 'image',
      source: { path: [0, 0] },
    });
  });

  it('keeps inline resources in an owned text box above their enclosing front drawing', () => {
    const parent = {
      ...drawing('parent', 0, 'image', 'image:parent', 1),
      textBoxIds: ['owned-box'],
    };
    const nestedBounds = { xPt: 0, yPt: 0, widthPt: 40, heightPt: 40 };
    const nested = inlineImageParagraph(
      'nested-paragraph',
      { story: 'textbox', storyInstance: 'owned-box', path: [0] },
      'image:nested',
      nestedBounds,
      2,
    );
    const owned = textBox(
      'owned-box', nested, undefined,
      { xPt: 0, yPt: 0, widthPt: 40, heightPt: 20 },
    );
    const layout = layoutFor([parent]);
    const root = layout.pages[0].layers.roots[0].node as ParagraphLayout;
    Object.assign(root, { textBoxes: [owned] });
    // Rebuild layers so the retained drawing entry owns the text box exactly as paint does.
    Object.assign(layout.pages[0], {
      layers: buildPageLayers([{ layer: 'body', node: root }]),
    });
    const registry = createPaintResourceRegistry([
      {
        kind: 'image', resourceKey: 'image:parent', partPath: 'word/media/parent.png',
        mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
      },
      {
        kind: 'image', resourceKey: 'image:nested', partPath: 'word/media/nested.png',
        mimeType: 'image/png', intrinsicSize: { widthPt: 40, heightPt: 40 },
      },
    ]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 20, yPt: 20 }, registry,
    )).toMatchObject({
      elementType: 'image',
      source: { story: 'textbox', storyInstance: 'owned-box', path: [0, 2] },
    });
    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 20, yPt: 40 }, registry,
    )).toMatchObject({
      elementType: 'image',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
    });
  });

  it('does not hit the clipped-away portion of an inline resource', () => {
    const resourceBounds = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 };
    const clipBounds = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 50 };
    const clipped = inlineImageParagraph(
      'clipped-paragraph', source([0]), 'image:clipped', resourceBounds, 1, clipBounds,
    );
    const layout = layoutFor([]);
    Object.assign(layout.pages[0], {
      layers: buildPageLayers([{ layer: 'body', node: clipped }]),
      readingOrder: [clipped.id],
    });
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: 'image:clipped', partPath: 'word/media/clipped.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 100 },
    }]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 25, yPt: 25 }, registry,
    )).toMatchObject({ elementType: 'image' });
    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 25, yPt: 75 }, registry,
    )).toBeNull();
  });

  it('inherits an exact table-cell clip when projecting an inline resource', () => {
    const resourceBounds = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 };
    const child = inlineImageParagraph(
      'cell-paragraph', source([0, 0, 0, 0]), 'image:cell', resourceBounds, 1,
    );
    const table: TableLayout = {
      kind: 'table',
      id: 'table',
      source: source([0]),
      flowDomainId: 'body',
      flowBounds: resourceBounds,
      inkBounds: resourceBounds,
      advancePt: 50,
      ordinaryFlow: true,
      columnWidthsPt: [100],
      rows: [{
        kind: 'table-row',
        id: 'row',
        source: source([0, 0]),
        flowDomainId: 'body',
        flowBounds: resourceBounds,
        inkBounds: resourceBounds,
        advancePt: 50,
        ordinaryFlow: true,
        heightPt: 50,
        contentHeightPt: 100,
        cells: [{
          kind: 'table-cell',
          id: 'cell',
          source: source([0, 0, 0]),
          flowDomainId: 'body',
          flowBounds: resourceBounds,
          inkBounds: resourceBounds,
          clipBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 50 },
          advancePt: 50,
          ordinaryFlow: true,
          contentBounds: resourceBounds,
          verticalMerge: 'none',
          vAlign: 'top',
          blocks: [{ layout: child, offsetPt: 0, advancePt: 100 }],
        }],
      }],
      borders: [],
    };
    const layout = layoutFor([]);
    Object.assign(layout.pages[0], {
      layers: buildPageLayers([{ layer: 'body', node: table }]),
      readingOrder: [table.id],
    });
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: 'image:cell', partPath: 'word/media/cell.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 100 },
    }]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 25, yPt: 25 }, registry,
    )).toMatchObject({ elementType: 'image' });
    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 25, yPt: 75 }, registry,
    )).toBeNull();
  });

  it('does not inspect retained text placement details while hit-testing elements', () => {
    const image = drawing('image', 0, 'image', 'image:1', 1);
    const layout = layoutFor([image]);
    const root = layout.pages[0].layers.roots[0].node as ParagraphLayout;
    const textPlacement = new Proxy({ kind: 'text' } as never, {
      get(target, property, receiver) {
        if (property !== 'kind') throw new Error('element hit-testing must not inspect text placement details');
        return Reflect.get(target, property, receiver);
      },
    });
    Object.assign(root, {
      lines: [{
        range: { start: 0, end: 1 },
        bounds,
        baselinePt: 30,
        advancePt: 20,
        placements: [textPlacement],
      }],
    });
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: 'image:1', partPath: 'word/media/image1.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 60 },
    }]);

    expect(hitTestDocxElementContext(
      layout, 0, { xPt: 20, yPt: 30 }, registry,
    )).toMatchObject({ kind: 'element', elementType: 'image' });
  });
});
