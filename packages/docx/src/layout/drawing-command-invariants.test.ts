import { describe, expect, it } from 'vitest';
import type { DocumentLayout, DrawingLayout } from './types.js';
import { assertDocumentLayout } from './invariants.js';

function layoutWith(command: DrawingLayout['commands'][number]): DocumentLayout {
  const bounds = { xPt: 72, yPt: 72, widthPt: 100, heightPt: 50 };
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
        bounds: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
      }],
      section: {} as DocumentLayout['pages'][number]['section'],
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0',
        coordinateSpace: { writingMode: 'horizontal-tb', logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
        blockStartPt: 72, blockEndPt: 720, flowDomainIds: ['body'],
        section: {} as DocumentLayout['pages'][number]['section'],
      }],
      layers: {
        paintOrder: [{ layer: 'front', nodeId: drawing.id, coordinateSpace: 'physical-page-points' }],
        background: [], behindText: [], header: [], body: [], notes: [],
        front: [drawing], footer: [],
      },
      readingOrder: [drawing.id],
    }],
    diagnostics: [],
  };
}

const validShape = {
  kind: 'drawingml-shape' as const,
  plan: {
    rect: { x: 72, y: 72, w: 100, h: 50 },
    geometry: {
      kind: 'preset' as const,
      name: 'rect',
      adjustments: [null, 1000, 2000],
    },
    fill: { fillType: 'solid' as const, color: 'FFFFFF' },
    stroke: { color: '000000', width: 1 },
    transform: { rotationDeg: 0, flipH: false, flipV: false },
  },
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
