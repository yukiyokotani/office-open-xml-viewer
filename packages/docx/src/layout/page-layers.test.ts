import { describe, expect, it } from 'vitest';
import { buildPageLayers } from './page-layers.js';
import type {
  DrawingPaintCommand,
  DrawingLayout,
  LayoutRect,
  PagePaintDrawingEntry,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  TextBoxLayout,
} from './types.js';

const bounds = Object.freeze({
  xPt: 0,
  yPt: 0,
  widthPt: 100,
  heightPt: 20,
}) satisfies LayoutRect;

function source(
  story: SourceRef['story'],
  path: readonly number[],
): SourceRef {
  return Object.freeze({
    story,
    storyInstance: `${story}:story`,
    path: Object.freeze([...path]),
  });
}

function drawing(
  id: string,
  behindDoc: boolean,
  relativeHeight: number,
  sourceOrder: number,
  story: SourceRef['story'] = 'body',
  textBoxIds: readonly string[] = [],
): DrawingLayout {
  return Object.freeze({
    kind: 'drawing',
    id,
    source: source(story, [sourceOrder]),
    flowDomainId: `${story}:domain`,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 0,
    ordinaryFlow: false,
    commands: Object.freeze([]),
    anchorLayer: Object.freeze({
      occurrenceId: `anchor:${id}`,
      behindDoc,
      relativeHeight,
      sourceOrder,
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    }),
    ...(textBoxIds.length > 0 ? { textBoxIds: Object.freeze([...textBoxIds]) } : {}),
  });
}

function paragraph(
  id: string,
  story: SourceRef['story'],
  drawings: readonly DrawingLayout[] = [],
  textBoxes: readonly TextBoxLayout[] = [],
  clipBounds?: LayoutRect,
): ParagraphLayout {
  return Object.freeze({
    kind: 'paragraph',
    id,
    source: source(story, [Number(id.replace(/\D/g, '')) || 0]),
    flowDomainId: `${story}:domain`,
    flowBounds: bounds,
    inkBounds: bounds,
    ...(clipBounds ? { clipBounds } : {}),
    advancePt: 20,
    ordinaryFlow: story === 'body',
    spacing: Object.freeze({ beforePt: 0, afterPt: 0 }),
    contextualSpacing: false,
    lines: Object.freeze([]),
    borders: Object.freeze([]),
    resources: Object.freeze([]),
    drawings: Object.freeze([...drawings]),
    textBoxes: Object.freeze([...textBoxes]),
    events: Object.freeze([]),
    exclusions: Object.freeze([]),
  });
}

function textBox(
  id: string,
  blocks: readonly ParagraphLayout[],
): TextBoxLayout {
  return Object.freeze({
    kind: 'textbox',
    id,
    source: source('textbox', [0]),
    flowDomainId: 'textbox:domain',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 20,
    ordinaryFlow: false,
    story: Object.freeze({
      story: 'textbox',
      flowBounds: bounds,
      inkBounds: bounds,
      blocks: Object.freeze([...blocks]),
      advancePt: 20,
      diagnostics: Object.freeze([]),
    }),
    transform: Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 5, f: 6 }),
    writingMode: 'horizontal-tb',
    insets: Object.freeze({ topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 }),
  });
}

function tableWithNestedParagraph(node: ParagraphLayout): TableLayout {
  const tableClip = Object.freeze({ xPt: 1, yPt: 2, widthPt: 90, heightPt: 80 });
  const cellClip = Object.freeze({ xPt: 3, yPt: 4, widthPt: 70, heightPt: 60 });
  return Object.freeze({
    kind: 'table',
    id: 'table',
    source: source('body', [20]),
    flowDomainId: 'body:domain',
    flowBounds: bounds,
    inkBounds: bounds,
    clipBounds: tableClip,
    advancePt: 40,
    ordinaryFlow: true,
    columnWidthsPt: Object.freeze([100]),
    rows: Object.freeze([Object.freeze({
      kind: 'table-row',
      id: 'row',
      source: source('body', [20, 0]),
      flowDomainId: 'body:domain',
      flowBounds: bounds,
      inkBounds: bounds,
      advancePt: 40,
      ordinaryFlow: true,
      heightPt: 40,
      contentHeightPt: 40,
      cells: Object.freeze([Object.freeze({
        kind: 'table-cell',
        id: 'cell',
        source: source('body', [20, 0, 0]),
        flowDomainId: 'body:domain',
        flowBounds: Object.freeze({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 40 }),
        inkBounds: bounds,
        clipBounds: cellClip,
        advancePt: 40,
        ordinaryFlow: true,
        contentBounds: Object.freeze({ xPt: 15, yPt: 20, widthPt: 90, heightPt: 40 }),
        verticalMerge: 'none',
        vAlign: 'top',
        blocks: Object.freeze([Object.freeze({
          layout: node,
          offsetPt: 7,
          advancePt: 20,
        })]),
      })]),
    })]),
    borders: Object.freeze([]),
  });
}

function drawingEntries(
  layers: ReturnType<typeof buildPageLayers>,
): readonly PagePaintDrawingEntry[] {
  return layers.paintOrder.filter(
    (entry): entry is PagePaintDrawingEntry => entry.kind === 'drawing',
  );
}

describe('buildPageLayers', () => {
  it('materializes ECMA-376 anchor order before paint with stable document-order ties', () => {
    const layers = buildPageLayers([
      {
        layer: 'body',
        node: paragraph('p1', 'body', [
          drawing('behind-high', true, 20, 4),
          drawing('front-high', false, 20, 4),
        ]),
      },
      {
        layer: 'body',
        node: paragraph('p2', 'body', [
          drawing('behind-tie-1', true, 10, 2),
          drawing('front-tie-1', false, 10, 2),
        ]),
      },
      {
        layer: 'body',
        node: paragraph('p3', 'body', [
          drawing('behind-tie-2', true, 10, 2),
          drawing('front-tie-2', false, 10, 2),
        ]),
      },
    ]);

    expect(layers.paintOrder.map((entry) => `${entry.kind}:${entry.node.id}`)).toEqual([
      'drawing:behind-tie-1',
      'drawing:behind-tie-2',
      'drawing:behind-high',
      'node:p1',
      'node:p2',
      'node:p3',
      'drawing:front-tie-1',
      'drawing:front-tie-2',
      'drawing:front-high',
    ]);
    expect(layers.roots.map((entry) => entry.node.id)).toEqual(['p1', 'p2', 'p3']);
    expect(layers.body.map((node) => node.id)).toEqual(['p1', 'p2', 'p3']);
  });

  it('retains structured-clone-safe transform and clip frames for a nested table anchor', () => {
    const paragraphClip = Object.freeze({ xPt: 8, yPt: 9, widthPt: 50, heightPt: 30 });
    const anchored = drawing('nested-front', false, 10, 1);
    const table = tableWithNestedParagraph(
      paragraph('p1', 'body', [anchored], [], paragraphClip),
    );

    const layers = buildPageLayers([{ layer: 'body', node: table }]);
    const entry = drawingEntries(layers)[0]!;

    expect(entry.node).toBe(anchored);
    expect(entry.ownerNodeId).toBe('p1');
    expect(entry.layoutTranslationPt).toEqual({ xPt: 15, yPt: 27 });
    expect(entry.frames).toEqual([
      { kind: 'clip', clip: table.clipBounds },
      { kind: 'clip', clip: table.rows[0]!.cells[0]!.clipBounds },
      { kind: 'transform', transform: { a: 1, b: 0, c: 0, d: 1, e: 15, f: 27 } },
      { kind: 'clip', clip: paragraphClip },
    ]);
    expect(structuredClone(layers).paintOrder).toEqual(layers.paintOrder);
    expect(Object.isFrozen(layers.paintOrder)).toBe(true);
    expect(Object.isFrozen(entry.frames)).toBe(true);
  });

  it('keeps text-box-owned anchors atomic with their owner instead of double-registering them', () => {
    const nested = drawing('textbox-nested', false, 100, 3, 'textbox');
    const owned = textBox('owned-textbox', [paragraph('tp1', 'textbox', [nested])]);
    const owner = drawing('owner', false, 10, 1, 'body', [owned.id]);

    const layers = buildPageLayers([
      { layer: 'body', node: paragraph('p1', 'body', [owner], [owned]) },
    ]);

    expect(drawingEntries(layers).map((entry) => entry.node.id)).toEqual(['owner']);
    expect(drawingEntries(layers)[0]?.textBoxes).toEqual([owned]);
  });

  it('keeps each page story as a distinct stacking context in retained compatibility order', () => {
    const roots = [
      ['header', 'header'] as const,
      ['body', 'body'] as const,
      ['notes', 'footnote'] as const,
      ['footer', 'footer'] as const,
    ].map(([layer, story], index) => ({
      layer,
      node: paragraph(
        `p${index + 1}`,
        story,
        [drawing(`${story}-front`, false, 10, index, story)],
      ),
    }));

    const layers = buildPageLayers(roots);

    expect(layers.paintOrder.map((entry) => ({
      kind: entry.kind,
      id: entry.node.id,
      sourceLayer: entry.sourceLayer,
      layer: entry.layer,
    }))).toEqual([
      { kind: 'node', id: 'p1', sourceLayer: 'header', layer: 'header' },
      { kind: 'drawing', id: 'header-front', sourceLayer: 'header', layer: 'front' },
      { kind: 'node', id: 'p2', sourceLayer: 'body', layer: 'body' },
      { kind: 'drawing', id: 'body-front', sourceLayer: 'body', layer: 'front' },
      { kind: 'node', id: 'p3', sourceLayer: 'notes', layer: 'notes' },
      { kind: 'drawing', id: 'footnote-front', sourceLayer: 'notes', layer: 'front' },
      { kind: 'node', id: 'p4', sourceLayer: 'footer', layer: 'footer' },
      { kind: 'drawing', id: 'footer-front', sourceLayer: 'footer', layer: 'front' },
    ]);
  });

  it('retains the element-backed target requirement from a proven vertical glyph operation', () => {
    const ordinary = paragraph('p1', 'body');
    const vertical = Object.freeze({
      ...ordinary,
      lines: Object.freeze([Object.freeze({
        placements: Object.freeze([Object.freeze({
          kind: 'text',
          paintOps: Object.freeze([Object.freeze({ verticalFeature: true })]),
        })]),
      })]) as unknown as ParagraphLayout['lines'],
    });

    expect(buildPageLayers([{ layer: 'body', node: ordinary }]).capabilities)
      .toEqual({ requiresElementBackedVerticalGlyphPaint: false, resourceKeys: [] });
    expect(buildPageLayers([{ layer: 'body', node: vertical }]).capabilities)
      .toEqual({ requiresElementBackedVerticalGlyphPaint: true, resourceKeys: [] });
  });

  it('retains exactly the resource keys reachable from the immutable page graph', () => {
    const inline = Object.freeze({
      ...paragraph('p1', 'body'),
      resources: Object.freeze([
        Object.freeze({
          kind: 'image',
          resourceKey: 'image:inline',
          intrinsicSize: Object.freeze({ widthPt: 10, heightPt: 10 }),
        }),
      ]),
      drawings: Object.freeze([Object.freeze({
        ...drawing('drawing', false, 1, 1),
        commands: Object.freeze([
          Object.freeze({
            kind: 'resource',
            resourceKey: 'chart:drawing',
            resourceKind: 'chart',
            rect: bounds,
          }),
          Object.freeze({
            kind: 'drawingml-image-fill',
            resourceKey: 'image:shape-fill',
            plan: Object.freeze({}),
          }) as unknown as DrawingPaintCommand,
        ]),
      })]),
    }) as ParagraphLayout;

    expect(buildPageLayers([{ layer: 'body', node: inline }]).capabilities.resourceKeys)
      .toEqual(['image:inline', 'chart:drawing', 'image:shape-fill']);
  });
});
