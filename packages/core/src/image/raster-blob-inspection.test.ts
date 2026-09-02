import { describe, expect, it } from 'vitest';
import { isDecodeTargetResizableRasterFormat } from './raster-blob-inspection';

describe('isDecodeTargetResizableRasterFormat', () => {
  it.each(['png', 'jpeg', 'gif', 'bmp', 'webp'] as const)(
    'admits browser resize for %s',
    (format) => {
      expect(isDecodeTargetResizableRasterFormat(format)).toBe(true);
    },
  );

  it('admits TIFF only when its optional decoder is available', () => {
    expect(isDecodeTargetResizableRasterFormat('tiff')).toBe(false);
    expect(isDecodeTargetResizableRasterFormat('tiff', true)).toBe(true);
  });

  it.each(['wmf', 'emf', null] as const)(
    'does not treat %s as a source raster grid',
    (format) => {
      expect(isDecodeTargetResizableRasterFormat(format, true)).toBe(false);
    },
  );
});
