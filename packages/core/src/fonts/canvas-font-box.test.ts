import { describe, expect, it } from 'vitest';
import { measureResolvedCanvasFontBoxRatio } from './canvas-font-box.js';

function contextFor(
  measure: (font: string, text: string) => Partial<TextMetrics>,
): Pick<CanvasRenderingContext2D, 'font' | 'measureText'> {
  let font = '12px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText(text: string) {
      return measure(font, text) as TextMetrics;
    },
  };
}

describe('measureResolvedCanvasFontBoxRatio', () => {
  it('returns the selected face box only when it differs from a missing-family control', () => {
    const context = contextFor((font) => font.includes('Available East Asian')
      ? {
          width: 100,
          fontBoundingBoxAscent: 106,
          fontBoundingBoxDescent: 44,
          actualBoundingBoxAscent: 80,
          actualBoundingBoxDescent: 21,
        }
      : {
          width: 92,
          fontBoundingBoxAscent: 90,
          fontBoundingBoxDescent: 25,
          actualBoundingBoxAscent: 78,
          actualBoundingBoxDescent: 20,
        });

    expect(measureResolvedCanvasFontBoxRatio(
      context,
      'Available East Asian',
      { text: '国', emPx: 100 },
    )).toBe(1.5);
    expect(context.font).toBe('12px serif');
  });

  it('returns null when the authored family resolves to the same fallback face', () => {
    const context = contextFor(() => ({
      width: 92,
      fontBoundingBoxAscent: 90,
      fontBoundingBoxDescent: 25,
      actualBoundingBoxAscent: 78,
      actualBoundingBoxDescent: 20,
    }));

    expect(measureResolvedCanvasFontBoxRatio(
      context,
      'Unavailable East Asian',
      { text: '国', emPx: 100 },
    )).toBeNull();
  });

  it('remains fail-closed if another document registers the control family', () => {
    const context = contextFor((font) => font.includes('__ooxml_missing_font_control_6f33c9b4__')
      ? {
          width: 75,
          fontBoundingBoxAscent: 130,
          fontBoundingBoxDescent: 40,
          actualBoundingBoxAscent: 73,
          actualBoundingBoxDescent: 19,
        }
      : {
          width: 92,
          fontBoundingBoxAscent: 90,
          fontBoundingBoxDescent: 25,
          actualBoundingBoxAscent: 78,
          actualBoundingBoxDescent: 20,
        });

    expect(measureResolvedCanvasFontBoxRatio(
      context,
      'Unavailable East Asian',
      { text: '国', emPx: 100 },
    )).toBeNull();
  });

  it('does not mistake a primary face box for coverage supplied by fallback', () => {
    const context = contextFor((font) => font.includes('Latin Only Face')
      ? {
          width: 92,
          fontBoundingBoxAscent: 120,
          fontBoundingBoxDescent: 40,
          actualBoundingBoxAscent: 78,
          actualBoundingBoxDescent: 20,
        }
      : {
          width: 92,
          fontBoundingBoxAscent: 90,
          fontBoundingBoxDescent: 25,
          actualBoundingBoxAscent: 78,
          actualBoundingBoxDescent: 20,
        });

    expect(measureResolvedCanvasFontBoxRatio(
      context,
      'Latin Only Face',
      { text: '国', emPx: 100 },
    )).toBeNull();
  });

  it('rejects unusable font boxes without leaking the probe font', () => {
    const context = contextFor((font) => font.includes('Broken Face')
      ? { width: 100, fontBoundingBoxAscent: Number.NaN, fontBoundingBoxDescent: 0 }
      : { width: 92, fontBoundingBoxAscent: 90, fontBoundingBoxDescent: 25 });

    expect(measureResolvedCanvasFontBoxRatio(
      context,
      'Broken Face',
      { text: '国', emPx: 100 },
    )).toBeNull();
    expect(context.font).toBe('12px serif');
  });
});
