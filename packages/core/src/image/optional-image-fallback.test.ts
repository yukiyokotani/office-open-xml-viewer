import { describe, expect, it, vi } from 'vitest';
import {
  OptionalImageCodecUnavailableError,
  isOptionalImageCodecUnavailableError,
  paintOptionalImagePlaceholder,
} from './optional-image-fallback.js';

describe('optional image codec fallback', () => {
  it('distinguishes a missing optional TIFF codec from a TIFF decode failure', () => {
    const error = new OptionalImageCodecUnavailableError('tiff');

    expect(isOptionalImageCodecUnavailableError(error, 'tiff')).toBe(true);
    expect(isOptionalImageCodecUnavailableError({
      code: 'ooxml-optional-image-codec-unavailable',
      codec: 'tiff',
    }, 'tiff')).toBe(true);
    expect(isOptionalImageCodecUnavailableError({
      code: 'ooxml-tiff-decode',
      codec: 'tiff',
    }, 'tiff')).toBe(false);
  });

  it('paints a constant-work placeholder inside the authored image bounds', () => {
    const fillText = vi.fn();
    const save = vi.fn();
    const restore = vi.fn();
    const ctx = {
      save,
      restore,
      fillText,
      fillStyle: '',
      font: '',
      textAlign: 'left',
      textBaseline: 'alphabetic',
    } as unknown as CanvasRenderingContext2D;

    paintOptionalImagePlaceholder(ctx, 'tiff', { x: 10, y: 20, width: 80, height: 40 });

    expect(save).toHaveBeenCalledOnce();
    expect(fillText).toHaveBeenCalledWith('TIFF image unavailable', 50, 40, 80);
    expect(restore).toHaveBeenCalledOnce();
  });
});
