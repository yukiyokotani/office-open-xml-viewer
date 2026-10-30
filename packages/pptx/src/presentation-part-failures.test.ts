import { afterEach, describe, expect, it, vi } from 'vitest';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { ProgressiveLayoutObserverNotifier } from '@silurus/ooxml-core/internal/progressive-layout-observers';
import { deferred } from '@silurus/ooxml-core/internal/deferred';
import { PptxPresentation } from './presentation.js';
import { pptxPartFailureLog } from './part-failures.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';

/**
 * #1974 — `PptxPresentation.partFailures` in worker mode: one-record deltas on
 * progressive publications, admitted before the slide becomes available, a
 * reconciling final list, and main-thread resolution of `errorPlaceholder`.
 */
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

const MESSAGE = 'ppt/slides/slide2.xml: unexpected end of stream';
const bootstrap = {
  slideCount: 2,
  slideWidth: 9144000,
  slideHeight: 6858000,
  defaultTextColor: null,
  majorFont: null,
  minorFont: null,
  hlinkColor: null,
  folHlinkColor: null,
  embeddedFonts: [],
  slides: [0, 1].map((index) => ({ index, partName: `ppt/slides/slide${index + 1}.xml` })),
} as const;
const slideFacts = [0, 1].map((index) => ({
  index,
  partName: `ppt/slides/slide${index + 1}.xml`,
  notes: null,
  hidden: false,
  mediaElements: [],
}));
const failure = { kind: 'slide', slideIndex: 1, slideNumber: 2, message: MESSAGE } as const;

function workerPresentation() {
  let resolveFinal!: (response: RenderWorkerResponse) => void;
  const final = new Promise<RenderWorkerResponse>((resolve) => { resolveFinal = resolve; });
  const renderRequests: Extract<RenderWorkerRequest, { kind: 'renderSlide' }>[] = [];
  const bridge = {
    request: (build: (id: number) => RenderWorkerRequest) => {
      const request = build(41);
      if (request.kind === 'parse') return final;
      if (request.kind === 'renderSlide') {
        renderRequests.push(request);
        return Promise.resolve({ kind: 'slideRendered', id: request.id, bitmap: {} as ImageBitmap, runs: [] });
      }
      throw new Error(`unexpected request ${request.kind}`);
    },
    post: vi.fn(),
    terminate: vi.fn(),
  };
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  Object.assign(instance, {
    _mode: 'worker',
    _bridge: bridge,
    _destroyed: false,
    _layoutWaiters: new Set(),
    _availableSlideCount: 0,
    _layoutLifecycle: new ProgressiveLayoutLifecycle(),
    _layoutObservers: new ProgressiveLayoutObserverNotifier(),
    _parseRequestId: null,
    _progressive: null,
    _metrics: null,
  });
  const presentation = instance as unknown as PptxPresentation;
  const lifecycle = { firstPublication: deferred<void>(), published: false, deferred: false, settled: false };
  const internals = presentation as unknown as {
    _parse(...args: unknown[]): Promise<void>;
    _onWorkerLayoutPush(response: RenderWorkerResponse): void;
  };
  const parsing = internals._parse(
    new ArrayBuffer(4), { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null }, false, 500,
    undefined, undefined, lifecycle,
  );
  return { presentation, internals, parsing, resolveFinal, renderRequests };
}

describe('PptxPresentation part failures (#1974)', () => {
  it('admits a failure before its slide becomes available and reconciles the final list', async () => {
    vi.useFakeTimers();
    const { presentation, internals, parsing, resolveFinal, renderRequests } = workerPresentation();
    await Promise.resolve();
    internals._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 41, bootstrap, availableSlides: 1,
      slide: slideFacts[0], fontPreloadNames: [],
    });
    await parsing;
    expect(presentation.partFailures).toEqual([]);

    const availableAtAdmission: number[] = [];
    pptxPartFailureLog(presentation).subscribe(() => availableAtAdmission.push(presentation.availableSlideCount));
    internals._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 41, availableSlides: 2,
      slide: slideFacts[1], fontPreloadNames: [], partFailure: failure,
    });
    expect(availableAtAdmission).toEqual([1]);
    expect(presentation.availableSlideCount).toBe(2);
    expect(presentation.partFailures).toEqual([failure]);

    resolveFinal({
      kind: 'presentationReady', id: 41,
      preflight: { ...bootstrap, slides: slideFacts, fontPreloadNames: [] },
      partFailures: [failure],
    });
    await presentation.waitUntilLayoutComplete();
    expect(presentation.partFailures).toEqual([failure]);

    vi.stubGlobal('OffscreenCanvas', class { constructor(readonly width: number, readonly height: number) {} });
    const resolver = vi.fn((f: { slideNumber: number }) => ({ title: `Slide ${f.slideNumber} is broken`, detail: null }));
    await presentation.renderSlideToBitmap(0, { errorPlaceholder: resolver });
    await presentation.renderSlideToBitmap(1, { errorPlaceholder: resolver });
    expect(resolver).toHaveBeenCalledTimes(1);
    expect(resolver).toHaveBeenCalledWith(failure);
    expect(renderRequests[0]).not.toHaveProperty('errorPlaceholderText');
    expect(renderRequests[1]!.errorPlaceholderText).toEqual({ state: 'resolved', title: 'Slide 2 is broken', detail: null });
    for (const request of renderRequests) {
      expect(request).not.toHaveProperty('errorPlaceholder');
      expect(() => structuredClone({ ...request, canvas: undefined })).not.toThrow();
    }
    await presentation.renderSlideToBitmap(1, { errorPlaceholder: { title: 'Static' } });
    expect(renderRequests[2]!.errorPlaceholderText).toEqual({ state: 'resolved', title: 'Static', detail: undefined });
  });

  it('survives a callback that destroys the presentation during a progressive delta', async () => {
    vi.useFakeTimers();
    const { presentation, internals, parsing } = workerPresentation();
    await Promise.resolve();
    internals._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 41, bootstrap, availableSlides: 1,
      slide: slideFacts[0], fontPreloadNames: [],
    });
    await parsing;
    pptxPartFailureLog(presentation).subscribe(() => presentation.destroy());
    expect(() => internals._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 41, availableSlides: 2,
      slide: slideFacts[1], fontPreloadNames: [], partFailure: failure,
    })).not.toThrow();
    expect(presentation.partFailures).toEqual([]);
  });

  it('rejects a delta that names another slide', async () => {
    vi.useFakeTimers();
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const { presentation, internals, parsing } = workerPresentation();
    await Promise.resolve();
    internals._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 41, bootstrap, availableSlides: 1,
      slide: slideFacts[0], fontPreloadNames: [], partFailure: failure,
    });
    await parsing.catch(() => undefined);
    expect(presentation.partFailures).toEqual([]);
    expect(presentation.availableSlideCount).toBe(0);
  });

  it('validates worker-reported lists', () => {
    const presentation = Object.create(PptxPresentation.prototype) as PptxPresentation;
    const admit = (entries: unknown) => (presentation as unknown as {
      _admitPartFailures(entries: unknown, slideCount: number): void;
    })._admitPartFailures(entries, 2);
    expect(() => admit([{ ...failure, slideIndex: 2 }])).toThrow(/invalid part failure/);
    expect(() => admit([{ ...failure, message: 3 }])).toThrow(/invalid part failure/);
    admit([failure, failure]);
    expect(presentation.partFailures).toEqual([failure]);
  });
});
