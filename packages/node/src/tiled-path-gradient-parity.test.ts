import { loadPreviousPainters, type PreviousPainters } from '../../../tests/helpers/previous-painters';
import { beforeAll, describe, expect, it } from 'vitest';
import { paintDrawingMLShape, resolveDrawingMLGeometry, type GradientFill } from '@silurus/ooxml-core';
import { paintDrawingLayout } from '../../docx/src/paint/canvas-drawing';
import type { DrawingLayout } from '../../docx/src/layout/types';
import { renderSlide } from '../../pptx/src/renderer';
import { renderViewport } from '../../xlsx/src/renderer';
import type { Slide } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles;

// Compare every decoded pixel with the pinned production renderer running on
// this host. This protects paint frames and format wiring without treating one
// platform's native Skia output as a portable Office-fidelity reference.
// Plus has aspect-sensitive concave arms, gear6/gear9 cross the raster support
// boundary at tile aspect ratios, and star5 exercises a supported outline.
const presets = ['plus', 'gear6', 'gear9', 'star5'] as const;

describe.skipIf(!skia)('tiled path gradients retain main pixels in every format', () => {
  let previous: PreviousPainters;
  beforeAll(async () => { previous = await loadPreviousPainters(); }, 60000);
  it.each(['rect', 'shape'] as const)('%s retains main tiled connector decoration bytes', path => {
    const pixels = (paint: PreviousPainters['paintDrawingMLShape']) => {
      const c = new (skia as NonNullable<typeof skia>).Canvas(400, 400);
      const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
      paint(ctx, {
        rect: { x: 200, y: 200, w: 50, h: 1 },
        geometry: { kind: 'preset', name: 'line', adjustments: [] }, fill: null,
        stroke: { color: '000000', width: 20, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' },
          fill: { fillType: 'gradient', gradType: 'radial', path, angle: 0,
            tileRect: { r: .5 }, flip: 'xy', fillToRect: { l: .1, r: .5, t: .4, b: .1 },
            stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
              { position: 1, color: 'FFFFFF' }] } },
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
      return ctx.getImageData(0, 0, 400, 400).data;
    };
    expect(pixels(paintDrawingMLShape)).toEqual(pixels(previous.paintDrawingMLShape));
  }, 30000);

  it.each(['docx', 'pptx', 'xlsx'] as const)('%s retains rect/shape tile bytes', async format => {
    const current = { paintDrawingLayout, renderSlide, renderViewport };
    for (const path of ['rect', 'shape'] as const) for (const preset of presets) {
      const pixels = async (painters: Pick<PreviousPainters, 'paintDrawingLayout' | 'renderSlide' | 'renderViewport'>) => {
        const c = new (skia as NonNullable<typeof skia>).Canvas(200, 120);
        const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
        const fill: GradientFill = {
          fillType: 'gradient', gradType: 'radial', path, angle: 0,
          tileRect: { r: .5 }, fillToRect: { l: .1, r: .5, t: .4, b: .1 },
          stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
            { position: 1, color: 'FFFFFF' }],
        };
        if (format === 'docx') {
          const bounds = { xPt: 0, yPt: 0, widthPt: 200, heightPt: 120 };
          const plan = {
            rect: { x: 0, y: 0, w: 200, h: 120 },
            geometry: { kind: 'preset' as const, name: preset, adjustments: [] },
            fill, stroke: null, transform: { rotationDeg: 0, flipH: false, flipV: false },
          };
          const drawing: DrawingLayout = {
            kind: 'drawing', id: 'tiled-shape', source: { story: 'body', storyInstance: 'body', path: [0] },
            flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds, advancePt: 120, ordinaryFlow: false,
            commands: [{ kind: 'drawingml-shape', plan: {
              ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1),
            } }],
          };
          painters.paintDrawingLayout(drawing, { ctx, scale: 1, dpr: 1,
            resources: { paint: () => { throw new Error('shape must not paint a retained resource'); } } });
        } else if (format === 'pptx') {
          await painters.renderSlide(c as unknown as HTMLCanvasElement, {
            index: 0, slideNumber: 1, background: null, elements: [{
              type: 'shape', x: 0, y: 0, width: 200 * 9525, height: 120 * 9525,
              rotation: 0, flipH: false, flipV: false, geometry: preset,
              fill, stroke: null, textBody: null, custGeom: null, shadow: null,
            }],
          } as Slide, 200 * 9525, 120 * 9525, { width: 200, dpr: 1 });
        } else {
          painters.renderViewport(ctx, {
            name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
            freezeRows: 0, freezeCols: 0, defaultColWidth: 8.43, defaultRowHeight: 15,
            mergeCells: [], conditionalFormats: [], images: [], charts: [],
            defaultFontFamily: 'Calibri', defaultFontSize: 11,
            shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0,
              toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
              nativeExtCx: 200 * 9525, nativeExtCy: 120 * 9525,
              shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
                fill, geom: { type: 'preset', name: preset, adj: [] } }],
            }],
          } as Worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 });
        }
        return ctx.getImageData(0, 0, 200, 120).data;
      };
      expect(await pixels(current), `${format}/${path}/${preset}`).toEqual(await pixels(previous));
    }
  }, 30000);
});
