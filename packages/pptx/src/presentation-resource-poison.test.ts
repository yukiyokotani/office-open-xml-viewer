import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import type { Slide } from './types.js';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => { resolve = resolvePromise; });
  return { promise, resolve };
}

function slide(): Slide {
  return {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [],
  };
}

function resourceFailure(): OoxmlResourceLimitError {
  return new OoxmlResourceLimitError('media inflation limit', {
    stage: 'parsing',
    violation: {
      format: 'pptx',
      operation: 'extract-media',
      resource: 'archive-entry',
      metric: 'inflated-bytes',
      limit: 1024,
      observed: 1025,
      configurable: true,
      usage: {
        archiveEntryCount: 3,
        declaredInflatedBytes: 4096,
        distinctInflatedBytes: 1024,
        operationInflatedBytes: 1025,
      },
    },
  });
}

describe('PptxPresentation document-level resource poison', () => {
  it('replays the first media violation before cached render, image, and markdown work', async () => {
    const fatal = resourceFailure();
    const request = vi.fn(async () => { throw fatal; });
    const withSlide = vi.fn();
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'main',
      _bridge: { request },
      _preflight: {
        slideCount: 1,
        slideWidth: 914400,
        slideHeight: 914400,
        defaultTextColor: null,
        majorFont: null,
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts: [],
        slides: [{
          index: 0,
          notes: null,
          hidden: false,
          mediaElements: [],
        }],
        fontPreloadNames: [],
      },
      _slides: { withSlide },
      _rawParts: new BoundedRawPartCache({ maxEntries: 2, maxBytes: 1024 }),
      _resourceFailure: null,
    });
    const presentation = instance as unknown as PptxPresentation;

    await expect(presentation.getMedia('ppt/media/video1.mp4')).rejects.toBe(fatal);
    const requestsAfterFailure = request.mock.calls.length;

    await expect(presentation.renderSlide({} as HTMLCanvasElement, 0)).rejects.toBe(fatal);
    await expect(presentation.getImage('ppt/media/image1.png', 'image/png')).rejects.toBe(fatal);
    expect(withSlide).not.toHaveBeenCalled();
    expect(request).toHaveBeenCalledTimes(requestsAfterFailure);
  });

  it('rechecks poison when a previously queued cached render acquires its slide', async () => {
    const fatal = resourceFailure();
    const gate = deferred<void>();
    const request = vi.fn(async () => { throw fatal; });
    const withSlide = vi.fn(async (
      _index: number,
      consume: (value: Slide) => unknown,
    ) => {
      await gate.promise;
      return consume(slide());
    });
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'main',
      _bridge: { request },
      _preflight: {
        slideCount: 1,
        slideWidth: 914400,
        slideHeight: 914400,
        defaultTextColor: null,
        majorFont: null,
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts: [],
        slides: [{ index: 0, notes: null, hidden: false, mediaElements: [] }],
        fontPreloadNames: [],
      },
      _slides: { withSlide },
      _rawParts: new BoundedRawPartCache({ maxEntries: 2, maxBytes: 1024 }),
      _resourceFailure: null,
    });
    const presentation = instance as unknown as PptxPresentation;

    const queuedRender = presentation.renderSlide(
      {} as HTMLCanvasElement,
      0,
      { width: 100 },
    );
    await vi.waitFor(() => expect(withSlide).toHaveBeenCalledTimes(1));
    await expect(presentation.getMedia('ppt/media/video1.mp4')).rejects.toBe(fatal);
    gate.resolve();

    await expect(queuedRender).rejects.toBe(fatal);
  });
});
