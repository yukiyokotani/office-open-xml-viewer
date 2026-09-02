import { expect, test } from '@playwright/test';
import { deflateSync } from 'node:zlib';

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type: string, payload: Uint8Array): Buffer {
  const typeBytes = Buffer.from(type, 'ascii');
  const out = Buffer.alloc(12 + payload.byteLength);
  out.writeUInt32BE(payload.byteLength, 0);
  typeBytes.copy(out, 4);
  Buffer.from(payload).copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + payload.byteLength)), 8 + payload.byteLength);
  return out;
}

/** A valid, highly compressible one-bit indexed PNG. Its ZIP-like byte size is
 * tiny while its declared raster grid exactly matches the reported poster. */
function posterPng(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 1; // bit depth
  ihdr[9] = 3; // indexed colour
  const scanlines = Buffer.alloc((Math.ceil(width / 8) + 1) * height);
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('PLTE', Buffer.from([17, 34, 51, 255, 255, 255])),
    chunk('IDAT', deflateSync(scanlines, { level: 9 })),
    chunk('IEND', new Uint8Array()),
  ]);
}

test('decodes the reported 109,571,670-pixel poster at display resolution', async ({ page }) => {
  const sourceWidth = 12_090;
  const sourceHeight = 9_063;
  const targetWidth = 960;
  const targetHeight = 720;
  const png = posterPng(sourceWidth, sourceHeight);
  expect(png.byteLength).toBeLessThan(100_000);
  await page.goto('/iframe.html?id=internal-raster-decode-smoke--harness&viewMode=story');
  await expect(page.locator('[data-raster-decode-ready="true"]')).toHaveCount(1);

  const result = await page.evaluate(async (base64) => {
    const decodeRasterOrMetafile = (globalThis as unknown as {
      __ooxmlDecodeRasterOrMetafile: (
        data: Blob,
        options: { targetWidthPx: number; targetHeightPx: number },
      ) => Promise<ImageBitmap | null>;
    }).__ooxmlDecodeRasterOrMetafile;
    const binary = atob(base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const bitmap = await decodeRasterOrMetafile(
      new Blob([bytes], { type: 'image/png' }),
      { targetWidthPx: 960, targetHeightPx: 720 },
    );
    if (!bitmap) return null;
    const canvas = document.createElement('canvas');
    canvas.width = 1;
    canvas.height = 1;
    const context = canvas.getContext('2d');
    context?.drawImage(bitmap, 0, 0, 1, 1);
    const pixel = context ? [...context.getImageData(0, 0, 1, 1).data] : [];
    const dimensions = { width: bitmap.width, height: bitmap.height, pixel };
    bitmap.close();
    return dimensions;
  }, png.toString('base64'));

  expect(result).not.toBeNull();
  const scale = Math.max(targetWidth / sourceWidth, targetHeight / sourceHeight);
  const retainedWidth = Math.ceil(sourceWidth * scale);
  const retainedHeight = Math.ceil(sourceHeight * retainedWidth / sourceWidth);
  expect(result?.width).toBe(retainedWidth);
  // Engines may round the aspect-derived axis to the nearest pixel or upward.
  expect(result?.height).toBeGreaterThanOrEqual(targetHeight);
  expect(result?.height).toBeLessThanOrEqual(retainedHeight);
  expect(result?.pixel).toEqual([17, 34, 51, 255]);
});
