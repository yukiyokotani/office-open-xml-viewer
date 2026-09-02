import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { getPosterBitmap } from './renderer.js';
import {
  acquireBitmapCacheLease,
  dropBitmapCacheByPath,
  getCachedBitmapByPath,
  MAX_DECODED_IMAGE_BYTES,
} from '@silurus/ooxml-core';
import type { MediaElement } from './types';

/**
 * RB1 (poster path): a `<p:pic>` media element's poster image is attacker-
 * controllable bytes (`posterPath` / `posterMimeType` come from the
 * `<a:blip>` in shape.rs). `getPosterBitmap` used to hand the raw poster blob
 * straight to `createImageBitmap`, which sizes its decoded RGBA surface from the
 * image HEADER — so a tiny PNG declaring 60000×60000 forces a ~14 GB allocation
 * (a decompression bomb) that OOMs the tab, bypassing the RB1 guard that already
 * protects picture blips.
 *
 * The fix routes the poster through the shared raster decoder and decoded-
 * surface owner. These tests assert the bomb is rejected BEFORE
 * `createImageBitmap`, normal posters decode, and pictures/posters share one
 * presentation-level live-byte budget.
 */

/** Big-endian u32 into a byte array at `o`. */
function putBeU32(b: Uint8Array, o: number, v: number): void {
  b[o] = (v >>> 24) & 0xff;
  b[o + 1] = (v >>> 16) & 0xff;
  b[o + 2] = (v >>> 8) & 0xff;
  b[o + 3] = v & 0xff;
}

/** A PNG header (8-byte sig + IHDR) declaring `w × h` with almost no payload. */
function pngHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(26);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  putBeU32(b, 8, 13);
  b.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  putBeU32(b, 16, w);
  putBeU32(b, 20, h);
  return b;
}

const SENTINEL = { width: 1, height: 1, close: () => {} } as unknown as ImageBitmap;

function mediaEl(posterMimeType = 'image/png'): MediaElement {
  return {
    type: 'media',
    x: 0,
    y: 0,
    width: 100,
    height: 100,
    mediaKind: 'video',
    posterPath: 'ppt/media/image1.png',
    posterMimeType,
    mediaPath: 'ppt/media/media1.mp4',
    mimeType: 'video/mp4',
  } as unknown as MediaElement;
}

describe('getPosterBitmap — RB1 poster decode-bomb guard', () => {
  let createImageBitmapSpy: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    createImageBitmapSpy = vi.fn(async (_blob: Blob) => SENTINEL);
    vi.stubGlobal('createImageBitmap', createImageBitmapSpy);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('rejects a 60000×60000 PNG poster bomb WITHOUT calling createImageBitmap', async () => {
    const bomb = pngHeader(60000, 60000); // ~14 GB decoded — tiny on the wire
    const fetchMedia = vi.fn(
      async (_path: string) => new Blob([bomb as BlobPart], { type: 'image/png' }),
    );

    await expect(getPosterBitmap(mediaEl(), fetchMedia)).rejects.toThrow();
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
  });

  it('decodes a normal in-budget poster (guard does not block legitimate images)', async () => {
    const ok = pngHeader(1920, 1080);
    const fetchMedia = vi.fn(
      async (_path: string) => new Blob([ok as BlobPart], { type: 'image/png' }),
    );

    const bmp = await getPosterBitmap(mediaEl(), fetchMedia);
    expect(bmp).toBe(SENTINEL);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
  });

  it('downsamples the reported 109,571,670-pixel poster class to its display target', async () => {
    const large = pngHeader(12_090, 9_063);
    expect(12_090 * 9_063).toBe(109_571_670);
    const fetchMedia = vi.fn(
      async (_path: string) => new Blob([large as BlobPart], { type: 'image/png' }),
    );
    const resized = { width: 1280, height: 960, close() {} } as unknown as ImageBitmap;
    createImageBitmapSpy.mockResolvedValueOnce(resized);

    await expect(getPosterBitmap(
      mediaEl(),
      fetchMedia,
      undefined,
      undefined,
      { targetWidthPx: 1280, targetHeightPx: 960 },
    )).resolves.toBe(resized);
    expect(createImageBitmapSpy).toHaveBeenCalledWith(expect.any(Blob), {
      // Preserve aspect ratio while covering both requested axes. This source
      // is fractionally wider than 4:3, so 960 target rows require 1281 columns.
      resizeWidth: 1281,
      resizeQuality: 'high',
    });
  });

  it('leaves an unrecognized (non-raster) poster header to decode normally (fail-open)', async () => {
    // e.g. an SVG poster: not a recognized raster ⇒ not blocked by the sniff.
    const svg = '<svg xmlns="http://www.w3.org/2000/svg"/>';
    const fetchMedia = vi.fn(
      async (_path: string) => new Blob([svg], { type: 'image/svg+xml' }),
    );

    const bmp = await getPosterBitmap(mediaEl('image/svg+xml'), fetchMedia);
    expect(bmp).toBe(SENTINEL);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
  });

  it('shares the live decoded-byte ceiling with ordinary presentation images', async () => {
    const width = 4096;
    const baseHeight = MAX_DECODED_IMAGE_BYTES / 2 / width / 4;
    const posterHeight = baseHeight + 1;
    const closeBase = vi.fn();
    const closePoster = vi.fn();
    createImageBitmapSpy
      .mockResolvedValueOnce({ width, height: baseHeight, close: closeBase } as unknown as ImageBitmap)
      .mockResolvedValueOnce({ width, height: posterHeight, close: closePoster } as unknown as ImageBitmap);
    const fetchImage = vi.fn(async () => new Blob([new Uint8Array([1])]));
    const fetchMedia = vi.fn(async () => new Blob([new Uint8Array([2])]));
    const release = acquireBitmapCacheLease(fetchImage);

    await getCachedBitmapByPath('ppt/media/picture.bin', 'application/octet-stream', fetchImage);
    await expect(getPosterBitmap(mediaEl('application/octet-stream'), fetchMedia, fetchImage))
      .rejects.toMatchObject({
        code: 'ooxml-decoded-image-limit',
        metric: 'active-decoded-bytes',
      });
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchMedia).toHaveBeenCalledTimes(1);
    expect(closePoster).toHaveBeenCalledTimes(1);

    release();
    dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    expect(closeBase).toHaveBeenCalledTimes(1);
  });

  it('loads poster bytes through media extraction but retains them under the deck owner', async () => {
    const closePoster = vi.fn();
    createImageBitmapSpy.mockResolvedValueOnce({
      width: 640,
      height: 360,
      close: closePoster,
    } as unknown as ImageBitmap);
    const deckOwner = vi.fn(async () => {
      throw new Error('the deck owner is an identity, not the poster byte loader');
    });
    const fetchMedia = vi.fn(async () => new Blob([pngHeader(640, 360) as BlobPart], {
      type: 'image/png',
    }));

    await getPosterBitmap(mediaEl(), fetchMedia, deckOwner);
    expect(fetchMedia).toHaveBeenCalledOnce();
    expect(deckOwner).not.toHaveBeenCalled();

    dropBitmapCacheByPath(deckOwner);
    await Promise.resolve();
    expect(closePoster).toHaveBeenCalledOnce();
  });
});
