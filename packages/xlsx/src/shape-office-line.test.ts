import { describe, expect, it } from 'vitest';
import { PT_TO_PX, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { bindXlsxOfficeFontRoutes, drawShapeText } from './renderer.js';
import { xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import { shapeOfficeNaturalLineRatio } from './shape-office-line.js';
import type { ShapeText, Worksheet } from './types.js';

function shape(anchor: 't' | 'ctr' | 'b', spaceLine?: { type: 'pct'; val: number }): ShapeText {
  return {
    anchor, wrap: 'square', autoFit: 'none',
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    paragraphs: [{ align: 'l', spaceLine, runs: [{
      type: 'text', text: '個人予算表', fontFace: 'Meiryo UI', fontFaceEa: 'Meiryo UI',
      bold: true, italic: false, size: 25,
    }] }],
  };
}

function context(): { ctx: CanvasRenderingContext2D; draws: Array<{ y: number; font: string }> } {
  let font = '11px sans-serif';
  const draws: Array<{ y: number; font: string }> = [];
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    measureText() { return { width: 140, actualBoundingBoxAscent: 20 }; },
    fillText(_text: string, _x: number, y: number) { draws.push({ y, font }); },
    fillStyle: '#000', textBaseline: 'alphabetic',
  } as unknown as CanvasRenderingContext2D;
  return { ctx, draws };
}

const localRoute: OfficeFontFallbackRoute = {
  requestedFamily: 'Meiryo UI', family: '__exact_meiryo_ui_bold', source: 'local',
  resourceIdentity: 'office-local:local("MeiryoUI-Bold")',
  weight: 700, style: 'normal', metric: { family: '__exact_meiryo_ui_bold', synthesized: false },
};

describe('XLSX single-run DrawingML natural line', () => {
  it('projects distinct Office line classes from OpenType metadata', () => {
    const eastAsian = shape('t').paragraphs[0].runs[0];
    if (eastAsian.type !== 'text') throw new Error('test fixture must contain text');
    expect(shapeOfficeNaturalLineRatio(eastAsian, localRoute)).toBeCloseTo(1.65, 2);
    const latin = { ...eastAsian, text: 'Hgj', fontFace: 'Arial', fontFaceEa: undefined };
    const arialRoute: OfficeFontFallbackRoute = {
      ...localRoute, requestedFamily: 'Arial', family: '__exact_arial_bold',
      metric: { family: '__exact_arial_bold', synthesized: false },
    };
    // Excel PDF controls measured 1.157em at 25pt and 1.149em at 12pt.
    expect(shapeOfficeNaturalLineRatio(latin, arialRoute)).toBeCloseTo(1.15, 2);
  });

  it('uses the Excel-measured line box only for a positively loaded exact tuple', () => {
    // Independent Excel PDF controls: one 25pt Meiryo UI Bold line in a 98px
    // zero-inset box had a 1.646em natural line (12pt control: 1.649em).
    // The static OpenType projection is 1.651em; PDF raster rounding explains
    // the <0.01em difference. Top/bottom anchors expose the whole line box.
    for (const [anchor, expected] of [
      ['t', 25 * PT_TO_PX * 1.65 / 2],
      ['b', 98 - 25 * PT_TO_PX * 1.65 / 2],
    ] as const) {
      const { ctx, draws } = context();
      bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, { 'meiryo ui:700:normal': localRoute });
      drawShapeText(ctx, shape(anchor), 230, 98, 1);
      expect(draws).toHaveLength(1);
      expect(draws[0].y).toBeCloseTo(expected, 1);
      expect(draws[0].font).toContain('__exact_meiryo_ui_bold');
    }
  });

  it('keeps the ordinary line for absent routes and explicit line spacing', () => {
    const absent = context();
    drawShapeText(absent.ctx, shape('t'), 230, 98, 1);
    expect(absent.draws[0].y).toBeCloseTo(25 * PT_TO_PX * 1.2 / 2, 5);
    const explicit = context();
    bindXlsxOfficeFontRoutes(explicit.ctx, {} as Worksheet, { 'meiryo ui:700:normal': localRoute });
    drawShapeText(explicit.ctx, shape('t', { type: 'pct', val: 100000 }), 230, 98, 1);
    expect(explicit.draws[0].y).toBeCloseTo(absent.draws[0].y, 5);
    expect(explicit.draws[0].font).not.toContain('__exact_meiryo_ui_bold');
  });

  it('discovers catalogued shape tuples separately from cell fonts', () => {
    const ws = { rows: [{ cells: [{ value: { type: 'text', text: 'x', runs: [{
      text: 'x', font: { name: 'Calibri', bold: true, italic: false },
    }] } }] }], shapeGroups: [{ shapes: [
      { text: shape('t') },
      { text: { ...shape('t'), paragraphs: [{ ...shape('t').paragraphs[0], runs: [{
        type: 'text', text: 'x', fontFace: 'Calibri', bold: true, italic: false, size: 25,
      }] }] } },
      { text: { ...shape('t'), paragraphs: [{ ...shape('t').paragraphs[0], spaceLine: { type: 'pct', val: 100000 } }] } },
    ] }] } as unknown as Worksheet;
    expect(xlsxWorksheetOfficeFontRequests(ws)).toEqual([
      { family: 'Calibri', weight: 700, style: 'normal' },
      { family: 'Meiryo UI', weight: 700, style: 'normal' },
    ]);
  });
});
