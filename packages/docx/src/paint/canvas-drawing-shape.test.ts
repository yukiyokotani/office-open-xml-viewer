import { resolveDrawingMLGeometry, type DrawingMLShapePaintPlan as GeometryFixtureInput } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import type { DrawingLayout, TextBoxLayout } from '../layout/types.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import { paintDrawingWithOwnedTextBoxes } from './canvas-text.js';
import type { CanvasPaintContext } from './types.js';

function shapeDrawing(): DrawingLayout {
  const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
  return {
    kind: 'drawing',
    id: 'shape-1',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 50,
    ordinaryFlow: false,
    commands: [{
      kind: 'drawingml-shape',
      plan: retainedGeometryFixture({
        rect: { x: 10, y: 20, w: 100, h: 50 },
        geometry: { kind: 'preset', name: 'rect', adjustments: [] },
        fill: { fillType: 'solid', color: 'FF0000' },
        stroke: { color: '000000', width: 2 },
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      }),
    }],
  };
}

function recordingContext(): { context: CanvasPaintContext; operations: string[] } {
  const operations: string[] = [];
  const ctx = {
    fillStyle: '', strokeStyle: '', lineWidth: 1, lineCap: 'butt',
    globalAlpha: 1, font: '', textAlign: 'left', textBaseline: 'alphabetic',
    save: () => operations.push('save'),
    restore: () => operations.push('restore'),
    beginPath: () => operations.push('beginPath'),
    moveTo: () => operations.push('moveTo'),
    lineTo: () => operations.push('lineTo'),
    closePath: () => operations.push('closePath'),
    clip: () => operations.push('clip'),
    rect: () => operations.push('rect'),
    fill: () => operations.push('fill'),
    stroke: () => operations.push('stroke'),
    setLineDash: () => operations.push('setLineDash'),
    translate: () => operations.push('translate'),
    rotate: () => operations.push('rotate'),
    scale: () => operations.push('scale'),
    createLinearGradient: () => {
      operations.push('createLinearGradient');
      return { addColorStop: () => operations.push('addColorStop') };
    },
    fillText: () => operations.push('fillText'),
    measureText: () => { throw new Error('retained drawing paint must not measure'); },
  } as unknown as CanvasRenderingContext2D;
  return {
    operations,
    context: {
      ctx,
      scale: 4,
      dpr: 2,
      resources: { paint: () => { throw new Error('shape must not use the resource painter'); } },
    },
  };
}

describe('retained DrawingML shape painting', () => {
  it('rejects an unresolved DrawingML command before beginning paint', () => {
    const { context, operations } = recordingContext();
    const resolved = shapeDrawing();
    const command = resolved.commands[0];
    if (command.kind !== 'drawingml-shape') throw new Error('fixture lost shape');
    const { resolvedGeometry: _geometry, ...plan } = command.plan;
    const unresolved = { ...resolved, commands: [{ kind: 'drawingml-shape', plan }] } as unknown as DrawingLayout;
    expect(() => paintDrawingLayout(unresolved, context)).toThrow(/Missing retained/);
    expect(operations).toEqual([]);
  });

  it('clips a retained image resource to the DrawingML geometry and applies fillRect overscan', () => {
    const { context, operations } = recordingContext();
    const painted: unknown[] = [];
    const resourceContext: CanvasPaintContext = {
      ...context,
      resources: {
        paint: (resourceKey, kind, bounds) => painted.push(resourceKey, kind, bounds),
      },
    };
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
    const drawing: DrawingLayout = {
      kind: 'drawing', id: 'image-fill',
      source: { story: 'header', storyInstance: 'default', path: [0] },
      flowDomainId: 'header', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false,
      commands: [{
        kind: 'drawingml-image-fill', resourceKey: 'image:overlay',
        fillRect: { l: 0.1, t: 0, r: -0.05, b: 0.2 },
        plan: retainedGeometryFixture({
          rect: { x: 10, y: 20, w: 100, h: 50 },
          geometry: { kind: 'preset', name: 'rect', adjustments: [] },
          fill: null, stroke: { color: '000000', width: 1 },
          transform: { rotationDeg: 0, flipH: false, flipV: false },
        }),
      }],
    };

    paintDrawingLayout(drawing, resourceContext);

    expect(painted).toEqual([
      'image:overlay', 'image',
      { xPt: 20, yPt: 20, widthPt: 95, heightPt: 40 },
    ]);
    expect(operations).toContain('clip');
    expect(operations).toContain('stroke');
    expect(operations).not.toContain('fill');
  });

  it('counter-rotates one upright physical frame around a shape and its owned text box', () => {
    const { context, operations } = recordingContext();
    const localBounds = { xPt: -40, yPt: -15, widthPt: 80, heightPt: 30 };
    const drawing: DrawingLayout = {
      kind: 'drawing', id: 'vertical-right-arrow',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body',
      flowBounds: { xPt: 50, yPt: 80, widthPt: 30, heightPt: 80 },
      inkBounds: { xPt: 50, yPt: 80, widthPt: 30, heightPt: 80 },
      advancePt: 0, ordinaryFlow: false,
      orientation: 'upright-physical',
      transform: { a: 0, b: -1, c: 1, d: 0, e: 65, f: 120 },
      commands: [{
        kind: 'drawingml-shape',
        plan: retainedGeometryFixture({
          rect: { x: -40, y: -15, w: 80, h: 30 },
          geometry: { kind: 'preset', name: 'rightArrow', adjustments: [] },
          fill: { fillType: 'solid', color: 'FF0000' },
          stroke: null,
          transform: { rotationDeg: 0, flipH: false, flipV: false },
        }),
      }],
      textBoxIds: ['vertical-right-arrow-text'],
    };
    const textBox: TextBoxLayout = {
      kind: 'textbox', id: 'vertical-right-arrow-text',
      source: { story: 'textbox', storyInstance: 'shape', path: [0] },
      flowDomainId: 'body:textbox', flowBounds: localBounds, inkBounds: localBounds,
      advancePt: 0, ordinaryFlow: false,
      story: {
        story: 'textbox', flowBounds: localBounds, inkBounds: localBounds,
        advancePt: 10, diagnostics: [],
        blocks: [{
          kind: 'paragraph', id: 'vertical-right-arrow-text:p',
          source: { story: 'textbox', storyInstance: 'shape', path: [0] },
          flowDomainId: 'body:textbox',
          flowBounds: localBounds, inkBounds: localBounds,
          advancePt: 10, ordinaryFlow: true,
          spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
          lines: [{
            range: { start: 0, end: 1 },
            bounds: { xPt: -30, yPt: -5, widthPt: 10, heightPt: 10 },
            baselinePt: 3, advancePt: 10,
            placements: [{
              kind: 'text', text: 'R', range: { start: 0, end: 1 },
              origin: { xPt: -30, yPt: 3 },
              bounds: { xPt: -30, yPt: -5, widthPt: 10, heightPt: 10 },
              advancePt: 10,
              clusters: [{
                range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10,
              }],
              paintOps: [{
                text: 'R', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 },
                letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto',
                writingMode: 'horizontal-tb',
              }],
              color: { kind: 'explicit', color: '#000000' },
              fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
              fontSizePt: 10, fontWeight: 400, fontStyle: 'normal',
              direction: 'ltr', decorations: [],
            }],
          }],
          borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
        }],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    };

    paintDrawingWithOwnedTextBoxes(drawing, [textBox], context);

    expect(operations.slice(0, 3)).toEqual(['save', 'translate', 'rotate']);
    expect(operations).toContain('fillText');
    expect(operations.at(-1)).toBe('restore');
    expect(operations.indexOf('beginPath')).toBeLessThan(operations.indexOf('fillText'));
  });

  it('counter-rotates an upright physical image inside a vertical section frame', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 40 };
    let paintedBounds: typeof bounds | undefined;
    const resourceContext: CanvasPaintContext = {
      ...context,
      resources: {
        paint(_resourceKey, _kind, retainedBounds) {
          operations.push('paintResource');
          paintedBounds = retainedBounds;
        },
      },
    };
    const drawing = {
      kind: 'drawing', id: 'upright-image',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false,
      commands: [{
        kind: 'resource', resourceKind: 'image', resourceKey: 'image:body:0',
        rect: bounds, orientation: 'upright-physical',
      }],
    } as DrawingLayout;

    paintDrawingLayout(drawing, resourceContext);

    expect(operations).toEqual(['save', 'translate', 'rotate', 'paintResource', 'restore']);
    expect(paintedBounds).toEqual({ xPt: -20, yPt: -50, widthPt: 40, heightPt: 100 });
  });

  it('dispatches explicit shape plans to the shared point-space painter', () => {
    const { context, operations } = recordingContext();

    paintDrawingLayout(shapeDrawing(), context);

    expect(operations).toEqual(expect.arrayContaining([
      'save', 'beginPath', 'rect', 'fill', 'setLineDash', 'stroke', 'restore',
    ]));
  });

  it('paints a shaped watermark command without measuring', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
    const drawing: DrawingLayout = {
      kind: 'drawing', id: 'watermark',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false,
      commands: [{
        kind: 'watermark-text', rect: bounds, text: 'DRAFT',
        fill: { fillType: 'solid', color: '808080' },
        opacity: .4, rotationDeg: 315, fitShape: true, fontSizePt: 36,
        sourceBounds: { xPt: -5, yPt: -80, widthPt: 250, heightPt: 100 },
        spans: [{
          text: 'DRAFT', advancePt: 250,
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
          fontWeight: 700, fontStyle: 'italic',
        }],
      }],
    };

    paintDrawingLayout(drawing, context);

    expect(operations).toEqual([
      'save', 'translate', 'rotate', 'scale', 'translate', 'fillText', 'restore',
    ]);
  });

  it('retains gradient fill semantics and skips no-fill text without measuring', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
    const command = {
      kind: 'watermark-text' as const, rect: bounds, text: 'DRAFT',
      opacity: 1, rotationDeg: 0, fitShape: false, fontSizePt: 12,
      sourceBounds: { xPt: 0, yPt: -8, widthPt: 30, heightPt: 10 },
      spans: [{
        text: 'DRAFT', advancePt: 30,
        fontRoute: { familyList: 'Arial', scope: 'native' as const, fingerprint: 'arial' },
        fontWeight: 400, fontStyle: 'normal' as const,
      }],
    };
    const layout = (fill: null | { fillType: 'gradient'; angle: number; gradType: string; stops: { position: number; color: string }[] }): DrawingLayout => ({
      kind: 'drawing', id: 'watermark-fill', source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds, advancePt: 0, ordinaryFlow: false,
      commands: [{ ...command, fill }],
    });
    paintDrawingLayout(layout(null), context);
    expect(operations).toEqual([]);

    paintDrawingLayout(layout({
      fillType: 'gradient', angle: 0, gradType: 'linear',
      stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
    }), context);
    expect(operations).toEqual([
      'createLinearGradient', 'addColorStop', 'addColorStop',
      'save', 'translate', 'translate', 'fillText', 'restore',
    ]);
  });

  it('ignores explicit no-op drawing commands', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 };
    paintDrawingLayout({
      kind: 'drawing', id: 'noop', source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false, commands: [{ kind: 'noop' }],
    }, context);
    expect(operations).toEqual([]);
  });
});

function retainedGeometryFixture<T extends GeometryFixtureInput>(plan: T) {
  return { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) };
}
