import { describe, expect, it } from 'vitest';
import type { ShapeRun } from '../types.js';
import {
  vmlTextPathAcquisitionInput,
  type InternalShapeRun,
} from '../parser-model.js';
import type { TextLayoutService } from './text.js';
import { planShapeDrawing } from './shape-drawing-plan.js';

function shape(overrides: Partial<ShapeRun> = {}): ShapeRun {
  return {
    widthPt: 100, heightPt: 50, anchorXPt: 10, anchorYPt: 20,
    anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0,
    subpaths: [], presetGeometry: 'rect',
    adjValues: Array.from({ length: 12 }, (_, index) => index * 1000),
    fill: {
      fillType: 'gradient', angle: 45, gradType: 'linear',
      stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '0000FF' }],
    },
    stroke: '123456', strokeWidth: 2, strokeDash: 'dash', strokeCap: 'round',
    strokeCustomDash: [{ dash: 1.25, space: .75 }],
    strokeJoin: 'miter', strokeMiterLimit: 8, strokeAlignment: 'in',
    strokeCompound: 'dbl',
    headEnd: { type: 'arrow', w: 'sm', len: 'sm' },
    tailEnd: { type: 'triangle', w: 'med', len: 'lg' },
    rotation: 30, flipH: true, flipV: false,
    textBlocks: [{ text: 'must not enter shape command', fontSizePt: 10, alignment: 'left' }],
    ...overrides,
  };
}

function textPathPlan(
  input: ShapeRun,
  bounds: Parameters<typeof planShapeDrawing>[1],
  text: TextLayoutService,
) {
  return planShapeDrawing(input, bounds, text, vmlTextPathAcquisitionInput(input));
}

describe('ShapeRun drawing command planner', () => {
  it('projects only clone-safe DrawingML paint facts without raw ShapeRun or shape resources', () => {
    const input = shape();
    const result = planShapeDrawing(input, { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 });
    if (result.status !== 'planned') throw new Error('expected a planned shape');
    if (result.command.kind !== 'drawingml-shape') throw new Error('expected DrawingML shape command');

    input.adjValues![0] = 999999;
    input.fill = null;

    expect(result.command).toMatchObject({
      kind: 'drawingml-shape',
      plan: {
        rect: { x: 10, y: 20, w: 100, h: 50 },
        geometry: { kind: 'preset', name: 'rect', adjustments: expect.any(Array) },
        fill: { fillType: 'gradient' },
        stroke: {
          color: '123456', width: 2, dashStyle: 'dash', lineCap: 'round',
          customDash: [{ dash: 1.25, space: .75 }],
          lineJoin: 'miter', miterLimit: 8, alignment: 'in', cmpd: 'dbl',
          headEnd: { type: 'arrow', w: 'sm', len: 'sm' },
          tailEnd: { type: 'triangle', w: 'med', len: 'lg' },
        },
        transform: { rotationDeg: 30, flipH: true, flipV: false },
      },
    });
    expect(result.command.plan.geometry.kind === 'preset'
      && result.command.plan.geometry.adjustments).toHaveLength(12);
    expect(JSON.stringify(result.command)).not.toContain('textBlocks');
    expect(JSON.stringify(result.command)).not.toContain('resourceKind');
    expect(structuredClone(result.command)).toEqual(result.command);
    expect(Object.isFrozen(result.command.plan)).toBe(true);
  });

  it('retains custom curve and arc commands as plain normalized geometry', () => {
    const result = planShapeDrawing(shape({
      presetGeometry: null,
      subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: .2, y1: 0, x2: .8, y2: 1, x: 1, y: 1 },
        { cmd: 'arcTo', wr: .2, hr: .3, stAng: 0, swAng: 180 },
      ]],
    }), { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 });
    if (result.status !== 'planned') throw new Error('expected a planned shape');
    if (result.command.kind !== 'drawingml-shape') throw new Error('expected DrawingML shape command');

    expect(result.command.plan.geometry).toMatchObject({
      kind: 'custom',
      subpaths: [[
        { cmd: 'moveTo' }, { cmd: 'cubicBezTo' }, { cmd: 'arcTo' },
      ]],
    });
    expect(result.command.plan.geometry).not.toHaveProperty('paint');
  });

  it('carries per-path fill and stroke flags aligned with the subpaths', () => {
    const subpaths: ShapeRun['subpaths'] = [
      [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1, y: 1 }],
      [{ cmd: 'moveTo', x: 1, y: 0 }, { cmd: 'lineTo', x: 0, y: 1 }],
    ];
    const plan = (subpathPaint: ShapeRun['subpathPaint']) => {
      const result = planShapeDrawing(shape({ presetGeometry: null, subpaths, subpathPaint }),
        { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 });
      if (result.status !== 'planned' || result.command.kind !== 'drawingml-shape') {
        throw new Error('expected a DrawingML shape command');
      }
      return result.command.plan.geometry;
    };
    expect(plan([{ fill: 'none' }, { stroke: false }])).toMatchObject({
      kind: 'custom', paint: [{ fill: 'none' }, { stroke: false }],
    });
    expect(plan([{ fill: 'none' }])).not.toHaveProperty('paint');
  });

  it('plans a stretched blip fill as a resource clipped by the shape geometry', () => {
    const result = planShapeDrawing(shape({
      fill: {
        fillType: 'image', imagePath: 'word/media/overlay.png', mimeType: 'image/png',
        fillRect: { l: 0, t: 0, r: -0.07574, b: 0 }, alpha: 0.5,
      },
    }), { xPt: 1, yPt: 2, widthPt: 100, heightPt: 40 }, undefined, undefined, 'image:key');

    expect(result).toMatchObject({
      status: 'planned',
      command: {
        kind: 'drawingml-image-fill',
        resourceKey: 'image:key',
        fillRect: { l: 0, t: 0, r: -0.07574, b: 0 },
        plan: {
          rect: { x: 1, y: 2, w: 100, h: 40 },
          fill: null,
          geometry: { kind: 'preset', name: 'rect' },
        },
      },
    });
    expect(structuredClone(result.command)).toEqual(result.command);
  });

  it('diagnoses tiled blip fills instead of silently stretching them', () => {
    const result = planShapeDrawing(shape({
      fill: {
        fillType: 'image', imagePath: 'word/media/tile.png', mimeType: 'image/png',
        tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
      },
    }), { xPt: 0, yPt: 0, widthPt: 100, heightPt: 40 }, undefined, undefined, 'image:tile');

    expect(result).toEqual({
      status: 'unsupported',
      command: { kind: 'noop' },
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        message: 'Tiled DrawingML shape image fills are not rendered',
      }],
    });
  });

  it('retains authored textPath shaping and fitshape geometry as a clone-safe command', () => {
    const requests: Parameters<TextLayoutService['shape']>[0][] = [];
    const text = {
      shape: (request: Parameters<TextLayoutService['shape']>[0]) => {
        requests.push(request);
        return ({
        advancePt: 250, ascentPt: 80, descentPt: 20,
        inkBounds: { xMinPt: -5, xMaxPt: 245, ascentPt: 70, descentPt: 10 },
        spans: [{
          text: 'DRAFT', start: 0, end: 5, script: 'ascii', breakBefore: true,
          advancePt: 250, ascentPt: 80, descentPt: 20,
          inkBounds: { xMinPt: -5, xMaxPt: 245, ascentPt: 70, descentPt: 10 },
          font: { family: 'Arial', route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 700, style: 'italic', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }],
        graphemeBoundaries: [0, 1, 2, 3, 4, 5], diagnostics: [],
      });
      },
    } as unknown as TextLayoutService;
    const result = textPathPlan(shape({
      textPath: {
        string: 'DRAFT', fontFamily: 'Arial', bold: true, italic: true,
        textPathOk: true, on: true, fitShape: true, fitPath: false,
        trim: false, xScale: false, fontSizePt: 36,
      },
      fill: { fillType: 'solid', color: '808080' }, fillOpacity: .4, rotation: 315,
    } as InternalShapeRun), { xPt: 1, yPt: 2, widthPt: 300, heightPt: 120 }, text);

    expect(result).toMatchObject({
      status: 'planned',
      command: {
        kind: 'watermark-text',
        rect: { xPt: 1, yPt: 2, widthPt: 300, heightPt: 120 },
        text: 'DRAFT', fill: { fillType: 'solid', color: '808080' },
        opacity: .4, rotationDeg: 315, fitShape: true,
        fontSizePt: 36,
        sourceBounds: { xPt: 0, yPt: -80, widthPt: 250, heightPt: 100 },
        spans: [{ text: 'DRAFT', advancePt: 250, fontWeight: 700, fontStyle: 'italic' }],
      },
    });
    expect(requests).toMatchObject([{ fontSizePt: 36, text: 'DRAFT' }]);
    expect(structuredClone(result)).toEqual(result);
  });

  it('honours parser display controls while keeping absent private facts compatible', () => {
    const text = {
      shape: () => ({
        advancePt: 10, ascentPt: 8, descentPt: 2,
        spans: [{
          text: 'X', start: 0, end: 1, script: 'ascii', breakBefore: true,
          advancePt: 10, ascentPt: 8, descentPt: 2,
          font: { route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 400, style: 'normal', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }], graphemeBoundaries: [0, 1], diagnostics: [],
      }),
    } as unknown as TextLayoutService;
    const bounds = { xPt: 1, yPt: 2, widthPt: 30, heightPt: 12 };

    const pathDisabled = textPathPlan(shape({
      textPath: { string: 'X', textPathOk: false, on: true, fitShape: true },
    } as InternalShapeRun), bounds, text);
    const textDisabled = textPathPlan(shape({
      textPath: { string: 'X', textPathOk: true, on: false, fitShape: true },
    } as InternalShapeRun), bounds, text);
    const manualPublic = textPathPlan(shape({ textPath: { string: 'X' } }), bounds, text);

    expect(pathDisabled.command.kind).toBe('drawingml-shape');
    expect(textDisabled.command.kind).toBe('drawingml-shape');
    expect(manualPublic.command.kind).toBe('watermark-text');
  });

  it('uses tight ink bounds for trim and typographic bounds otherwise', () => {
    const text = {
      shape: () => ({
        advancePt: 20, ascentPt: 12, descentPt: 4,
        inkBounds: { xMinPt: -2, xMaxPt: 23, ascentPt: 7, descentPt: 1 },
        spans: [{
          text: 'A', start: 0, end: 1, script: 'ascii', breakBefore: true,
          advancePt: 20, ascentPt: 12, descentPt: 4,
          inkBounds: { xMinPt: -2, xMaxPt: 23, ascentPt: 7, descentPt: 1 },
          font: { route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 400, style: 'normal', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }], graphemeBoundaries: [0, 1], diagnostics: [],
      }),
    } as unknown as TextLayoutService;
    const make = (trim: boolean) => textPathPlan(shape({
      textPath: {
        string: 'A', textPathOk: true, on: true, fitShape: true,
        fitPath: false, xScale: false, trim, fontSizePt: 18,
      },
    } as InternalShapeRun), { xPt: 0, yPt: 0, widthPt: 20, heightPt: 20 }, text);

    expect(make(false).command).toMatchObject({
      kind: 'watermark-text',
      sourceBounds: { xPt: 0, yPt: -12, widthPt: 20, heightPt: 16 },
    });
    expect(make(true).command).toMatchObject({
      kind: 'watermark-text',
      sourceBounds: { xPt: -2, yPt: -7, widthPt: 25, heightPt: 8 },
    });
  });

  it('paints zero-advance ink only when trim selects the tight ink box', () => {
    const text = {
      shape: () => ({
        advancePt: 0, ascentPt: 12, descentPt: 4,
        inkBounds: { xMinPt: -2, xMaxPt: 3, ascentPt: 7, descentPt: 1 },
        spans: [{
          text: '\u0301', start: 0, end: 1, script: 'highAnsi', breakBefore: true,
          advancePt: 0, ascentPt: 12, descentPt: 4,
          inkBounds: { xMinPt: -2, xMaxPt: 3, ascentPt: 7, descentPt: 1 },
          font: { route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 400, style: 'normal', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }], graphemeBoundaries: [0, 1], diagnostics: [],
      }),
    } as unknown as TextLayoutService;
    const make = (trim: boolean) => textPathPlan(shape({
      textPath: {
        string: '\u0301', textPathOk: true, on: true, fitShape: true,
        fitPath: false, xScale: false, trim, fontSizePt: 18,
      },
    } as InternalShapeRun), { xPt: 0, yPt: 0, widthPt: 20, heightPt: 20 }, text);

    expect(make(true).command).toMatchObject({
      kind: 'watermark-text',
      sourceBounds: { xPt: -2, yPt: -7, widthPt: 5, heightPt: 8 },
    });
    expect(make(false)).toEqual({
      status: 'unsupported',
      command: { kind: 'noop' },
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        message: 'VML textPath produced empty glyph metrics',
      }],
    });
  });

  it('turns whitespace-only content into a paint no-op without shaping', () => {
    const text = { shape: () => { throw new Error('whitespace must not be shaped'); } } as unknown as TextLayoutService;
    const result = textPathPlan(shape({
      textPath: {
        string: ' \t\n', textPathOk: true, on: true, fitShape: true,
        fitPath: false, trim: false, xScale: false,
      },
    } as InternalShapeRun), { xPt: 0, yPt: 0, widthPt: 20, heightPt: 20 }, text);

    expect(result.command).toEqual({ kind: 'noop' });
  });

  it.each(['fitPath', 'xScale'] as const)('isolates unsupported active %s semantics as a diagnostic noop', (control) => {
    const text = { shape: () => { throw new Error('unsupported mode must be rejected before shaping'); } } as unknown as TextLayoutService;
    const result = textPathPlan(shape({
      textPath: {
        string: 'DRAFT', textPathOk: true, on: true, fitShape: false,
        fitPath: false, trim: false, xScale: false, [control]: true,
      } as NonNullable<InternalShapeRun['textPath']>,
    }), { xPt: 0, yPt: 0, widthPt: 20, heightPt: 20 }, text);

    expect(result).toEqual({
      status: 'unsupported',
      command: { kind: 'noop' },
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        message: `VML textPath ${control}=true is not rendered`,
      }],
    });
  });

  it.each([
    [Number.NaN, /must contain finite numbers/],
    [-1, /must be finite and non-negative/],
  ] as const)('keeps invalid textPath font size %s fatal', (fontSizePt, expected) => {
    const text = { shape: () => { throw new Error('invalid input must fail before shaping'); } } as unknown as TextLayoutService;
    expect(() => textPathPlan(shape({
      textPath: {
        string: 'DRAFT', textPathOk: true, on: true, fitShape: true,
        fitPath: false, trim: false, xScale: false, fontSizePt,
      },
    } as InternalShapeRun), { xPt: 0, yPt: 0, widthPt: 20, heightPt: 20 }, text))
      .toThrow(expected);
  });

  it('isolates fitShape=false without an authored font size before shaping', () => {
    const text = {
      shape: () => { throw new Error('unsupported input must not be shaped'); },
    } as unknown as TextLayoutService;

    const result = textPathPlan(shape({
      textPath: {
        string: 'DRAFT', fontFamily: 'Arial', bold: false, italic: false,
        textPathOk: true, on: true, fitShape: false, fitPath: false,
        trim: false, xScale: false,
      },
    } as InternalShapeRun), { xPt: 1, yPt: 2, widthPt: 300, heightPt: 120 }, text);

    expect(result).toEqual({
      status: 'unsupported',
      command: { kind: 'noop' },
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        message: 'VML textPath fitShape=false requires an authored font-size',
      }],
    });
  });

  it('isolates trim when the measurement adapter cannot provide ink bounds', () => {
    const text = {
      shape: () => ({
        advancePt: 20, ascentPt: 8, descentPt: 2,
        spans: [{
          text: 'DRAFT', start: 0, end: 5, script: 'ascii', breakBefore: true,
          advancePt: 20, ascentPt: 8, descentPt: 2,
          font: { route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 400, style: 'normal', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }], graphemeBoundaries: [0, 5], diagnostics: [],
      }),
    } as unknown as TextLayoutService;

    const result = textPathPlan(shape({
      textPath: {
        string: 'DRAFT', fontFamily: 'Arial', bold: false, italic: false,
        textPathOk: true, on: true, fitShape: true, fitPath: false,
        trim: true, xScale: false, fontSizePt: 12,
      },
    } as InternalShapeRun), { xPt: 1, yPt: 2, widthPt: 300, heightPt: 120 }, text);

    expect(result).toEqual({
      status: 'unsupported',
      command: { kind: 'noop' },
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        message: 'VML textPath trim=true requires glyph ink bounds',
      }],
    });
  });

  it('keeps non-finite shaped text metrics fatal', () => {
    const text = {
      shape: () => ({
        advancePt: Number.POSITIVE_INFINITY, ascentPt: 8, descentPt: 2,
        spans: [{
          text: 'DRAFT', start: 0, end: 5, script: 'ascii', breakBefore: true,
          advancePt: Number.POSITIVE_INFINITY, ascentPt: 8, descentPt: 2,
          font: { route: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' }, weight: 400, style: 'normal', diagnostics: [] },
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
        }], graphemeBoundaries: [0, 5], diagnostics: [],
      }),
    } as unknown as TextLayoutService;

    expect(() => textPathPlan(shape({
      textPath: {
        string: 'DRAFT', fontFamily: 'Arial', bold: false, italic: false,
        textPathOk: true, on: true, fitShape: true, fitPath: false,
        trim: false, xScale: false, fontSizePt: 12,
      },
    } as InternalShapeRun), { xPt: 1, yPt: 2, widthPt: 300, heightPt: 120 }, text))
      .toThrow('Shape textPath acquisition produced non-finite metrics');
  });
});
