import { afterAll, describe, expect, it, vi } from 'vitest';
import { ownsGlyphs, sameGlyphRaster } from './glyph-raster-proof.js';

// The coverage probe face answers zero advance and zero ink for the glyphs it
// owns; everything else comes from a fallback face with real metrics.
const OWNED = new Set(['国', 'あ']);
let lastContext: { direction: string; textAlign: string } | undefined;
const visibleRasterUnits: number[] = [];

class StubCanvas {
  constructor(public width: number, public height: number) {}
  getContext() {
    const context = {
      font: '', direction: 'inherit', textAlign: 'start', pixel: 0,
      clearRect() {},
      fillText(text: string, x: number) {
        const left = this.direction === 'rtl' && this.textAlign === 'start' ? x - text.length * 25 : x;
        visibleRasterUnits.push(Math.max(0, text.length - Math.ceil(Math.max(0, -left) / 25)));
        this.pixel = this.direction === 'ltr' && this.font.includes('comparison') ? 1 : 2;
      },
      getImageData() { return { data: new Uint8ClampedArray([0, 0, 0, this.pixel]) }; },
      measureText(text: string) {
        const owned = OWNED.has(text) || (text === 'A' && this.direction === 'rtl');
        return {
          width: owned ? 0 : text.length * 25,
          actualBoundingBoxLeft: 0,
          actualBoundingBoxRight: owned ? 0 : 20,
          actualBoundingBoxAscent: owned ? 0 : 70,
          actualBoundingBoxDescent: 0,
        };
      },
    };
    lastContext = context;
    return context;
  }
}

vi.stubGlobal('OffscreenCanvas', StubCanvas);
afterAll(() => vi.unstubAllGlobals());

const request = (text: string) => ({
  text, coverageProbeFamily: 'probe', weight: 700, style: 'normal' as const,
});

describe('ownsGlyphs (LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY)', () => {
  it('owns text whose every grapheme is the probe face', () => {
    expect(ownsGlyphs(request('国あ'))).toBe(true);
  });

  it('declines whitespace the face does not own', () => {
    expect(ownsGlyphs(request('国 あ'))).toBe(false);
  });

  it('declines empty text', () => {
    expect(ownsGlyphs(request(''))).toBe(false);
  });
});

it('keys native ownership and raster proofs by their explicit shaping base and restores it', () => {
  expect(ownsGlyphs({ ...request('A'), baseDirection: 'rtl' })).toBe(true);
  expect(ownsGlyphs({ ...request('A'), baseDirection: 'ltr' })).toBe(false);
  expect(ownsGlyphs({ ...request('A'), baseDirection: 'rtl' })).toBe(true);
  visibleRasterUnits.length = 0;
  const proof = { text: 'BBBB', weight: 400, style: 'normal' as const,
    route: { familyList: 'Fixture', scope: 'native' as const, fingerprint: 'fixture' }, faceFamily: 'comparison' };
  expect(sameGlyphRaster({ ...proof, baseDirection: 'rtl' })).toBe(true);
  expect(sameGlyphRaster({ ...proof, baseDirection: 'ltr' })).toBe(false);
  expect(visibleRasterUnits).toHaveLength(4);
  expect(visibleRasterUnits.every(count => count === 4)).toBe(true);
  expect([lastContext?.direction, lastContext?.textAlign]).toEqual(['inherit', 'start']);
});
