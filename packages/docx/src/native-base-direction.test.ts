import { expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, independentTextShapeRequest, replaceTextShapeRequest,
  sliceTextShapeRequest, type GlyphMeasurer, type TextLayoutService, type TextShapeRequest } from './layout/text.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocxDocumentModel } from './types.js';

const request: TextShapeRequest = { text: 'ab', fontSizePt: 12,
  fonts: { ascii: 'Fixture', highAnsi: 'Fixture', eastAsia: 'Fixture', complexScript: 'Fixture' } };
function directed(service: TextLayoutService, baseDirection: 'ltr' | 'rtl', input = request) {
  if (!service.shapeDirected) throw new Error('directed acquisition unavailable');
  return service.shapeDirected({ ...input, baseDirection });
}
const metric = (advancePt: number) => ({ advancePt, ascentPt: 8, descentPt: 2 });

it('keeps directed aggregate, contextual prefix and ordinary cache authority separate', () => {
  let ordinaryCalls = 0;
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
    fingerprint: 'direction-sensitive',
    measure: r => { ordinaryCalls++; return metric(r.text.length * 5); },
    measureDirected: r => ({ ...metric(r.text.length * (r.baseDirection === 'ltr' ? 10 : 20)),
      nativeBaseDirection: r.baseDirection }),
  } });
  expect(directed(service, 'ltr').advancePt).toBe(20);
  expect(directed(service, 'rtl').advancePt).toBe(40);
  expect(directed(service, 'ltr').clusters?.[0]?.advancePt).toBe(10);
  expect(directed(service, 'rtl').clusters?.[0]?.advancePt).toBe(20);
  expect(directed(service, 'rtl').nativeBaseDirection).toBe('rtl');
  expect(service.shape(request).advancePt).toBe(10);
  expect(ordinaryCalls).toBeGreaterThan(0);
});

it('declines a legacy stateful provider before changing its ordinary call history', () => {
  let calls = 0;
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
    fingerprint: 'stateful-legacy', measure: r => metric(++calls * r.text.length),
  } });
  expect(service.shapeDirected).toBeUndefined();
  expect(() => service.shape({ ...request, baseDirection: 'rtl' })).toThrow(/directed/i);
  expect(calls).toBe(0);
  expect(service.shape({ ...request, clusterGeometry: false }).advancePt).toBe(2);
  expect(calls).toBe(1);
});

it('declines an auxiliary provider without direction capability before measuring', () => {
  let calls = 0;
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
    fingerprint: 'old-raster-provider', measure: () => metric(++calls),
    measureDirected: r => ({ ...metric(++calls), nativeBaseDirection: r.baseDirection }),
    sameGlyphRaster: () => true,
  } });
  expect(service.shapeDirected).toBeUndefined();
  expect(() => service.shape({ ...request, baseDirection: 'ltr' })).toThrow(/directed/i);
  expect(calls).toBe(0);
});

it('rejects a broken opted-in receipt without retrying the ordinary provider', () => {
  let ordinaryCalls = 0;
  const measurer: GlyphMeasurer = { fingerprint: 'broken-direction-receipt',
    measure: () => { ordinaryCalls++; return metric(1); },
    measureDirected: () => metric(30),
  };
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer });
  expect(() => directed(service, 'rtl')).toThrow(/receipt/i);
  expect(ordinaryCalls).toBe(0);
});

it('keeps the same native base in cached raster proof requests', () => {
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
    fingerprint: 'directed-raster-proof', measure: () => metric(10),
    measureDirected: r => ({ ...metric(10), nativeBaseDirection: r.baseDirection }),
    supportsDirectedAuxiliaries: true,
    sameGlyphRaster: r => r.baseDirection === 'rtl',
  } });
  const proof = { text: 'ab', weight: 400, style: 'normal' as const,
    route: { familyList: 'Fixture', scope: 'native' as const, fingerprint: 'fixture' }, faceFamily: 'comparison' };
  expect(service.sameGlyphRaster?.({ ...proof, baseDirection: 'rtl' })).toBe(true);
  expect(service.sameGlyphRaster?.({ ...proof, baseDirection: 'ltr' })).toBe(false);
  expect(service.sameGlyphRaster?.({ ...proof, baseDirection: 'rtl' })).toBe(true);
});

it('preserves an acquired unit base for slices but not independent or replaced text', () => {
  const service = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
    fingerprint: 'slice-direction', measure: r => metric(r.text.length * 5),
    measureDirected: r => ({ ...metric(r.text.length * 20), nativeBaseDirection: r.baseDirection }),
  } });
  const original = { ...request, baseDirection: 'rtl' as const };
  const slice = sliceTextShapeRequest(original, 0, 1);
  expect(service.shape(slice).advancePt).toBe(20);
  expect(service.shape(slice).nativeBaseDirection).toBe('rtl');
  expect(service.shape(independentTextShapeRequest(original, 'a')).advancePt).toBe(5);
  expect(service.shape(replaceTextShapeRequest(original, 'a')).advancePt).toBe(5);
});

it('applies explicit Canvas direction and restores all measurement state after failure', () => {
  const context = { font: 'old-font', direction: 'inherit', textAlign: 'start', letterSpacing: '3px', fontKerning: 'none',
    measureText(this: { direction: string; textAlign: string }, text: string) {
      if (text === 'boom') throw new Error('backend failure');
      const width = text.length * (this.direction === 'rtl' ? 20 : 10);
      const left = this.direction === 'rtl' && this.textAlign === 'start' ? width : 0;
      return { width, actualBoundingBoxLeft: left, actualBoundingBoxRight: width - left,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
  const model = { body: [], section: {}, headers: {}, footers: {} } as unknown as DocxDocumentModel;
  const service = createLayoutServices(model, { measureContext: context }).text;
  const acquired = directed(service, 'rtl');
  expect(acquired.advancePt).toBe(40);
  expect(acquired.inkBounds?.xMinPt).toBe(0);
  expect(acquired.inkBounds?.xMaxPt).toBe(40);
  expect(() => directed(service, 'rtl', { ...request, text: 'boom', kerning: true })).toThrow('backend failure');
  expect([context.font, context.direction, context.textAlign, context.letterSpacing, context.fontKerning])
    .toEqual(['old-font', 'inherit', 'start', '3px', 'none']);
});
