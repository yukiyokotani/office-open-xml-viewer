import { resolveDrawingMLGeometry, type DrawingMLShapePaintPlan as GeometryFixtureInput } from '@silurus/ooxml-core';
import { createPaintResourceRegistry } from './paint-resources.js';
import { acquireNativeReadingImagePlan } from './native-reading-image-frame.js';
import { describe, expect, it } from 'vitest';
import type { DocumentLayout, DrawingLayout } from './types.js';
import { assertDocumentLayout } from './invariants.js';
import { buildPageLayers } from './page-graph.js';

function layoutWith(command: DrawingLayout['commands'][number]): DocumentLayout {
  const bounds = { xPt: 72, yPt: 72, widthPt: 100, heightPt: 50 };
  const section: DocumentLayout['pages'][number]['section'] = {
    geometry: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
    columns: [{ xPt: 72, wPt: 468 }],
    columnSeparator: false,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection: 'lrTb', verticalAlignment: 'top',
  };
  const drawing: DrawingLayout = {
    kind: 'drawing',
    id: 'shape-1',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 0,
    ordinaryFlow: false,
    commands: [command],
  };
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        xPt: 0, yPt: 0, widthPt: 612, heightPt: 792,
        contentTopPt: 72, contentBottomPt: 720,
      },
      flowDomains: [{
        id: 'body', kind: 'body',
        logicalBounds: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
        physicalBounds: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
      }],
      section,
      sectionOccurrenceId: 'section:0',
      parityBlank: false,
      bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      columnSeparators: [],
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0',
        coordinateSpace: {
          writingMode: 'horizontal-tb',
          logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        },
        blockStartPt: 72, blockEndPt: 720,
        columnFlowDirection: 'ltr', columnIndexes: [0],
        flowDomainIds: ['body'], section,
      }],
      pageBorder: null,
      layers: buildPageLayers([
        { layer: 'front', node: drawing, coordinateSpace: 'section-logical' },
      ]),
      readingOrder: [drawing.id],
    }],
    diagnostics: [],
  };
}

const validShape = {
  kind: 'drawingml-shape' as const,
  plan: retainedGeometryFixture({
    rect: { x: 72, y: 72, w: 100, h: 50 },
    geometry: {
      kind: 'preset' as const,
      name: 'rect',
      adjustments: [null, 1000, 2000],
    },
    fill: { fillType: 'solid' as const, color: 'FFFFFF' },
    stroke: { color: '000000', width: 1 },
    transform: { rotationDeg: 0, flipH: false, flipV: false },
  }),
};

describe('drawing-command invariants', () => {
  it('accepts explicit DrawingML shape geometry without a legacy command rect', () => {
    expect(() => assertDocumentLayout(layoutWith(validShape))).not.toThrow();
  });

  it('rejects invalid shape extents and non-finite adjustments', () => {
    const negative = structuredClone(validShape);
    negative.plan.rect.w = -1;
    expect(() => assertDocumentLayout(layoutWith(negative))).toThrow(/INVALID_GEOMETRY/);

    const nonFinite = structuredClone(validShape);
    nonFinite.plan.geometry.adjustments[1] = Number.NaN;
    expect(() => assertDocumentLayout(layoutWith(nonFinite))).toThrow(/INVALID_GEOMETRY/);
  });

  it('validates retained VML textPath source geometry and accepts explicit no-op', () => {
    expect(() => assertDocumentLayout(layoutWith({ kind: 'noop' }))).not.toThrow();
    const valid = {
      kind: 'watermark-text' as const,
      rect: { xPt: 72, yPt: 72, widthPt: 100, heightPt: 50 },
      text: 'DRAFT', fill: { fillType: 'solid' as const, color: '808080' },
      opacity: .5, rotationDeg: 315, fitShape: true, fontSizePt: 12,
      sourceBounds: { xPt: -1, yPt: -8, widthPt: 31, heightPt: 10 },
      spans: [{
        text: 'DRAFT', advancePt: 30,
        fontRoute: { familyList: 'Arial', scope: 'native' as const, fingerprint: 'arial' },
        fontWeight: 400, fontStyle: 'normal' as const,
      }],
    };
    expect(() => assertDocumentLayout(layoutWith(valid))).not.toThrow();

    const degenerate = structuredClone(valid);
    degenerate.sourceBounds.widthPt = 0;
    expect(() => assertDocumentLayout(layoutWith(degenerate))).toThrow(/INVALID_GEOMETRY/);
  });
});

describe('retained reading image command invariants', () => {
  it('validates cloned plans in drawing-local upright frames and rejects mismatched orientation/dimensions', () => {
    const key = 'image:invented-upright';
    const descriptor = createPaintResourceRegistry([{ kind: 'image', resourceKey: key, partPath: 'invented.png', mimeType: 'image/png', intrinsicSize: { widthPt: 100, heightPt: 50 } }]).resolve(key, 'image');
    const plan = acquireNativeReadingImagePlan(descriptor, 100, 50);
    const layout = layoutWith({ kind: 'resource', resourceKind: 'image', resourceKey: key,
      rect: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 50 }, nativeImagePlan: plan });
    const node = layout.pages[0].layers.front[0] as DrawingLayout;
    Object.assign(node, { orientation: 'upright-physical', transform: { a: 0, b: 1, c: -1, d: 0, e: 122, f: 72 },
      flowBounds: { xPt: 72, yPt: 72, widthPt: 50, heightPt: 100 }, inkBounds: { xPt: 72, yPt: 72, widthPt: 50, heightPt: 100 } });
    expect(() => assertDocumentLayout(structuredClone(layout))).not.toThrow();
    const wrongDimensions = structuredClone(layout), changed = wrongDimensions.pages[0].layers.front[0] as DrawingLayout;
    const command = changed.commands[0];
    if (command.kind !== 'resource') throw new Error('invalid fixture');
    Object.assign(command, { rect: { ...command.rect, widthPt: 101 } });
    expect(() => assertDocumentLayout(wrongDimensions)).toThrow(/projection mismatch/);
    const wrongOrientation = structuredClone(layout), oriented = wrongOrientation.pages[0].layers.front[0] as DrawingLayout;
    Object.assign(oriented.commands[0], { orientation: 'upright-physical' });
    expect(() => assertDocumentLayout(wrongOrientation)).toThrow(/unacquired image orientation/);
  });
});

function retainedGeometryFixture<T extends GeometryFixtureInput>(plan: T) {
  return { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) };
}
