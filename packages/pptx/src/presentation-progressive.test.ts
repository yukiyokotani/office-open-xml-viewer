import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { ProgressiveLayoutObserverNotifier } from '@silurus/ooxml-core/internal/progressive-layout-observers';
import {
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { PptxPresentation } from './presentation.js';
import type {
  PptxWorkerRequest,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol.js';
import type { Slide } from './types.js';

type PullResponse = PullSessionResponse<ArrayBuffer, number>;

afterEach(() => vi.useRealTimers());

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function slide(index: number): Slide {
  return {
    index,
    slideNumber: index + 1,
    partName: `ppt/slides/slide${index + 1}.xml`,
    background: null,
    elements: [],
    notes: `notes-${index + 1}`,
  };
}

function pullResponse(
  command: PullSessionCommand<number>,
  value: Record<string, unknown>,
): PullResponse {
  return {
    protocol: PULL_SESSION_PROTOCOL,
    sessionId: command.sessionId,
    operationId: command.operationId,
    generation: command.generation,
    requestId: command.requestId,
    ...value,
  } as PullResponse;
}

describe('PptxPresentation progressive layout lifecycle', () => {
  it('accepts worker-mode prefix pushes before the correlated final response', async () => {
    vi.useFakeTimers();
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const finalResponse = deferred<RenderWorkerResponse>();
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
      slides: [0, 1].map((index) => ({
        index,
        partName: `ppt/slides/slide${index + 1}.xml`,
      })),
    } as const;
    const slideFacts = [0, 1].map((index) => ({
      index,
      partName: `ppt/slides/slide${index + 1}.xml`,
      notes: `notes-${index + 1}`,
      hidden: false,
      mediaElements: [],
    }));
    let requestOptions: { timeoutMs?: number | false } | undefined;
    let renderRequest: Extract<RenderWorkerRequest, { kind: 'renderSlide' }> | undefined;
    const protocolOrder: string[] = [];
    const bridge = {
      request: (
        build: (id: number) => RenderWorkerRequest,
        _transfer: Transferable[] | undefined,
        options: { timeoutMs?: number | false } | undefined,
      ) => {
        const request = build(41);
        if (request.kind === 'parse') {
          expect(request).toMatchObject({ kind: 'parse', id: 41, progressiveLayout: true });
          requestOptions = options;
          return finalResponse.promise;
        }
        if (request.kind === 'renderSlide') {
          renderRequest = request;
          protocolOrder.push('renderSlide');
          return Promise.resolve({
            kind: 'slideRendered', id: request.id, bitmap: {} as ImageBitmap, runs: [],
          });
        }
        throw new Error(`unexpected request ${request.kind}`);
      },
      post: vi.fn(() => protocolOrder.push('continuePresentationPreflight')),
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
    const lifecycle = {
      firstPublication: deferred<void>(),
      published: false,
      deferred: false,
      settled: false,
      onProgress: () => { throw new Error('progress observer failed'); },
      onComplete: () => { throw new Error('complete observer failed'); },
    };
    const policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null } as const;

    const parsing = (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: typeof policy,
        useGoogleFonts: boolean,
      timeoutMs: number,
        onUsage: undefined,
        renderers: undefined,
        progressive: typeof lifecycle,
      ): Promise<void>;
    })._parse(new ArrayBuffer(4), policy, false, 500, undefined, undefined, lifecycle);
    await Promise.resolve();
    expect(requestOptions).toEqual({ timeoutMs: false });

    await vi.advanceTimersByTimeAsync(40);

    (presentation as unknown as {
      _onWorkerLayoutPush(response: RenderWorkerResponse): void;
    })._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial',
      forId: 41,
      bootstrap,
      availableSlides: 1,
      slide: slideFacts[0],
      fontPreloadNames: [],
    });
    await parsing;
    await presentation.renderSlideToBitmap(0, {
      imageResources: { decodedByteBudget: 64 * 1024 * 1024, strategy: 'strict' },
    });
    expect(renderRequest?.imageResources).toEqual({
      decodedByteBudget: 64 * 1024 * 1024,
      strategy: 'strict',
    });
    await vi.waitFor(() => expect(bridge.post).toHaveBeenCalledTimes(1));
    expect(protocolOrder).toEqual(['renderSlide', 'continuePresentationPreflight']);
    await vi.advanceTimersByTimeAsync(40);
    expect(bridge.terminate).not.toHaveBeenCalled();
    expect(presentation.slideCount).toBe(2);
    expect(presentation.availableSlideCount).toBe(1);
    expect(presentation.layoutComplete).toBe(false);

    finalResponse.resolve({
      kind: 'presentationReady',
      id: 41,
      preflight: {
        ...bootstrap,
        slides: slideFacts,
        fontPreloadNames: [],
      },
    });
    await presentation.waitUntilLayoutComplete();
    expect(presentation.availableSlideCount).toBe(2);
    expect(presentation.layoutComplete).toBe(true);
    expect(consoleError).toHaveBeenCalledTimes(2);
    consoleError.mockRestore();
  });

  it('publishes the opening slide while keeping the final slide count stable', async () => {
    const releaseSecondSlide = deferred<void>();
    let secondSlidePullStarted = false;
    const slideIndexBySession = new Map<number, number>();
    let pullRequestId = 1;
    const transport: WorkerBridgeTransport<PullResponse> = {
      request: async (build) => {
        const command = build(pullRequestId++) as PullSessionCommand<number>;
        if (command.kind === 'pull') {
          const index = slideIndexBySession.get(command.sessionId);
          if (index === undefined) throw new Error('missing slide session');
          if (index === 1) {
            secondSlidePullStarted = true;
            await releaseSecondSlide.promise;
          }
          const payload = new TextEncoder().encode(JSON.stringify(slide(index))).buffer;
          return pullResponse(command, {
            kind: 'chunk',
            sequence: command.sequence,
            byteLength: payload.byteLength,
            done: true,
            payload,
          });
        }
        return pullResponse(command, { kind: 'accepted', command: command.kind });
      },
      forgetOrphaned: () => undefined,
      terminate: () => undefined,
    };
    const bootstrap = {
      slideCount: 3,
      slideWidth: 9144000,
      slideHeight: 6858000,
      defaultTextColor: '111111',
      majorFont: 'Aptos Display',
      minorFont: 'Aptos',
      hlinkColor: '0563C1',
      folHlinkColor: null,
      embeddedFonts: [],
      slides: [0, 1, 2].map((index) => ({
        index,
        partName: `ppt/slides/slide${index + 1}.xml`,
      })),
    } as const;
    let ordinaryId = 100;
    const bridge = {
      request: async (build: (id: number) => PptxWorkerRequest) => {
        const request = build(ordinaryId++);
        if (request.kind === 'parse') {
          expect(request.progressiveLayout).toBe(true);
          return { kind: 'presentationOpened', id: request.id, bootstrap };
        }
        if (request.kind === 'openSlideSession') {
          slideIndexBySession.set(request.sessionId, request.slideIndex);
          return {
            kind: 'slideSessionOpened',
            id: request.id,
            sessionId: request.sessionId,
            operationId: request.operationId,
            generation: request.generation,
          };
        }
        throw new Error(`unexpected request ${request.kind}`);
      },
      transport: () => transport,
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance._bridge = bridge;
    instance._embeddedFontFaces = [];
    instance._embeddedFontAliases = new Map();
    instance._embeddedFontAuthoredFamilies = new Map();
    instance._destroyed = false;
    instance._layoutWaiters = new Set();
    instance._layoutLifecycle = new ProgressiveLayoutLifecycle();
    instance._layoutObservers = new ProgressiveLayoutObserverNotifier();
    const presentation = instance as unknown as PptxPresentation;
    const partials: number[] = [];
    const completions: unknown[] = [];
    const policy = {
      maxArchiveEntryBytes: null,
      maxTotalInflatedBytes: null,
    } as const;

    await (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: typeof policy,
        useGoogleFonts: boolean,
        timeoutMs: undefined,
        onUsage: undefined,
        renderers: undefined,
        progressive: {
          onPartial: (progress: { availableUnits: number }) => void;
          onComplete: (error?: unknown) => void;
          firstPublication: ReturnType<typeof deferred<void>>;
          published: boolean;
          deferred: boolean;
          settled: boolean;
        },
      ): Promise<void>;
    })._parse(
      new ArrayBuffer(4),
      policy,
      false,
      undefined,
      undefined,
      undefined,
      {
        onPartial: ({ availableUnits }) => partials.push(availableUnits),
        onComplete: (error) => completions.push(error),
        firstPublication: deferred<void>(),
        published: false,
        deferred: false,
        settled: false,
      },
    );

    expect(presentation.slideCount).toBe(3);
    expect(presentation.availableSlideCount).toBe(1);
    expect(presentation.layoutComplete).toBe(false);
    expect(presentation.getNotes(0)).toBe('notes-1');
    expect(presentation.getNotes(1)).toBeNull();
    expect(partials).toEqual([]);
    expect(secondSlidePullStarted).toBe(false);

    // The opening publication releases the load continuation before the next
    // host task starts slide 2 preflight. A Viewer can therefore enqueue the
    // opening paint/resource work in this gap, matching worker-mode ACK gating.
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect(secondSlidePullStarted).toBe(true);

    let completed = false;
    const completion = presentation.waitUntilLayoutComplete().then(() => {
      completed = true;
    });
    await Promise.resolve();
    expect(completed).toBe(false);

    releaseSecondSlide.resolve();
    await completion;

    expect(presentation.availableSlideCount).toBe(3);
    expect(presentation.layoutComplete).toBe(true);
    expect(presentation.getNotes(1)).toBe('notes-2');
    expect(partials).toEqual([2, 3]);
    expect(completions).toEqual([undefined]);
  });

  it('rethrows a background failure after the opening slide was published', async () => {
    const releaseFailure = deferred<void>();
    const slideIndexBySession = new Map<number, number>();
    let pullRequestId = 1;
    const transport: WorkerBridgeTransport<PullResponse> = {
      request: async (build) => {
        const command = build(pullRequestId++) as PullSessionCommand<number>;
        if (command.kind === 'pull') {
          const index = slideIndexBySession.get(command.sessionId);
          if (index === undefined) throw new Error('missing slide session');
          if (index === 1) {
            await releaseFailure.promise;
            throw new Error('later slide failed');
          }
          const payload = new TextEncoder().encode(JSON.stringify(slide(index))).buffer;
          return pullResponse(command, {
            kind: 'chunk', sequence: command.sequence,
            byteLength: payload.byteLength, done: true, payload,
          });
        }
        return pullResponse(command, { kind: 'accepted', command: command.kind });
      },
      forgetOrphaned: () => undefined,
      terminate: () => undefined,
    };
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
    let ordinaryId = 100;
    const bridge = {
      request: async (build: (id: number) => PptxWorkerRequest) => {
        const request = build(ordinaryId++);
        if (request.kind === 'parse') return { kind: 'presentationOpened', id: request.id, bootstrap };
        if (request.kind === 'openSlideSession') {
          slideIndexBySession.set(request.sessionId, request.slideIndex);
          return {
            ...request,
            kind: 'slideSessionOpened' as const,
          };
        }
        throw new Error(`unexpected request ${request.kind}`);
      },
      transport: () => transport,
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'main', _bridge: bridge, _embeddedFontFaces: [],
      _embeddedFontAliases: new Map(), _embeddedFontAuthoredFamilies: new Map(),
      _destroyed: false, _layoutWaiters: new Set(),
      _layoutLifecycle: new ProgressiveLayoutLifecycle(),
      _layoutObservers: new ProgressiveLayoutObserverNotifier(),
    });
    const presentation = instance as unknown as PptxPresentation;
    const failures: unknown[] = [];
    const policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null } as const;

    await (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: typeof policy,
        useGoogleFonts: boolean,
        timeoutMs: undefined,
        onUsage: undefined,
        renderers: undefined,
        progressive: {
          onComplete: (error?: unknown) => void;
          firstPublication: ReturnType<typeof deferred<void>>;
          published: boolean;
          deferred: boolean;
          settled: boolean;
        },
      ): Promise<void>;
    })._parse(new ArrayBuffer(4), policy, false, undefined, undefined, undefined, {
      onComplete: (error) => failures.push(error),
      firstPublication: deferred<void>(),
      published: false,
      deferred: false,
      settled: false,
    });

    releaseFailure.resolve();
    await expect(presentation.waitUntilLayoutComplete()).rejects.toThrow('later slide failed');
    expect(presentation.layoutComplete).toBe(false);
    expect(failures[0]).toBeInstanceOf(Error);
  });

  it('terminates a worker protocol that publishes a malformed prefix and rejects completion', async () => {
    const finalResponse = deferred<RenderWorkerResponse>();
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
      slides: [0, 1].map((index) => ({
        index,
        partName: `ppt/slides/slide${index + 1}.xml`,
      })),
    } as const;
    const facts = [0, 1].map((index) => ({
      index,
      partName: `ppt/slides/slide${index + 1}.xml`,
      notes: null,
      hidden: false,
      mediaElements: [],
    }));
    const bridge = {
      request: (build: (id: number) => RenderWorkerRequest) => {
        const request = build(73);
        expect(request.kind).toBe('parse');
        return finalResponse.promise;
      },
      post: vi.fn(),
      terminate: vi.fn(() => finalResponse.reject(new Error('worker terminated'))),
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'worker', _bridge: bridge, _destroyed: false,
      _layoutWaiters: new Set(), _availableSlideCount: 0,
      _layoutLifecycle: new ProgressiveLayoutLifecycle(),
      _layoutObservers: new ProgressiveLayoutObserverNotifier(),
      _parseRequestId: null, _progressive: null, _metrics: null,
    });
    const presentation = instance as unknown as PptxPresentation;
    const lifecycle = {
      firstPublication: deferred<void>(), published: false, deferred: false, settled: false,
    };
    const policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null } as const;
    const parsing = (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: typeof policy,
        useGoogleFonts: boolean,
        timeoutMs: undefined,
        onUsage: undefined,
        renderers: undefined,
        progressive: typeof lifecycle,
      ): Promise<void>;
    })._parse(new ArrayBuffer(4), policy, false, undefined, undefined, undefined, lifecycle);
    await Promise.resolve();
    const push = (response: RenderWorkerResponse) => (presentation as unknown as {
      _onWorkerLayoutPush(value: RenderWorkerResponse): void;
    })._onWorkerLayoutPush(response);
    push({
      kind: 'presentationLayoutPartial', forId: 73, bootstrap,
      availableSlides: 1, slide: facts[0], fontPreloadNames: [],
    });
    await parsing;

    push({
      kind: 'presentationLayoutPartial', forId: 73,
      availableSlides: 3, slide: facts[1], fontPreloadNames: [],
    });

    await expect(presentation.waitUntilLayoutComplete())
      .rejects.toThrow('PPTX progressive worker published a non-sequential slide');
    expect(bridge.terminate).toHaveBeenCalledTimes(1);
    expect(presentation.layoutComplete).toBe(false);
  });

  it.each(['main', 'worker'] as const)(
    'treats a one-slide progressive %s load as complete without a deferred callback',
    async (mode) => {
      const bootstrap = {
        slideCount: 1,
        slideWidth: 9144000,
        slideHeight: 6858000,
        defaultTextColor: null,
        majorFont: null,
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts: [],
        slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
      } as const;
      const facts = {
        index: 0,
        partName: 'ppt/slides/slide1.xml',
        notes: null,
        hidden: false,
        mediaElements: [],
      } as const;
      const bridge = { post: vi.fn(), terminate: vi.fn() };
      const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
      Object.assign(instance, {
        _mode: mode, _bridge: bridge, _destroyed: false,
        _layoutWaiters: new Set(), _availableSlideCount: 0,
        _layoutLifecycle: new ProgressiveLayoutLifecycle(),
        _layoutObservers: new ProgressiveLayoutObserverNotifier(),
        _parseRequestId: mode === 'worker' ? 91 : null,
        _progressive: null, _metrics: null,
      });
      const presentation = instance as unknown as PptxPresentation;
      const onComplete = vi.fn();
      const lifecycle = {
        firstPublication: deferred<void>(), published: false, deferred: false,
        settled: false, onComplete,
      };
      instance._progressive = lifecycle;
      const complete = { ...bootstrap, slides: [facts], fontPreloadNames: [] };

      if (mode === 'worker') {
        (presentation as unknown as {
          _onWorkerLayoutPush(response: RenderWorkerResponse): void;
        })._onWorkerLayoutPush({
          kind: 'presentationLayoutPartial', forId: 91, bootstrap,
          availableSlides: 1, slide: facts, fontPreloadNames: [],
        });
      } else {
        (presentation as unknown as {
          _applyProgressivePrefix(prefix: typeof complete, progressive: typeof lifecycle): void;
        })._applyProgressivePrefix(complete, lifecycle);
      }
      let loadReleased = false;
      void lifecycle.firstPublication.promise.then(() => { loadReleased = true; });
      await Promise.resolve();
      expect(loadReleased).toBe(false);
      (presentation as unknown as {
        _finishProgressiveLayout(prefix: typeof complete, progressive: typeof lifecycle): void;
      })._finishProgressiveLayout(complete, lifecycle);
      await lifecycle.firstPublication.promise;

      expect(lifecycle.deferred).toBe(false);
      expect(presentation.layoutComplete).toBe(true);
      expect(onComplete).not.toHaveBeenCalled();
    },
  );

  it('completes an actual one-slide main parse before releasing load', async () => {
    const bootstrap = {
      slideCount: 1, slideWidth: 9144000, slideHeight: 6858000,
      defaultTextColor: null, majorFont: null, minorFont: null,
      hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    } as const;
    let slideSessionId = 0;
    let requestId = 1;
    const transport: WorkerBridgeTransport<PullResponse> = {
      request: async (build) => {
        const command = build(requestId++) as PullSessionCommand<number>;
        if (command.kind === 'pull') {
          const payload = new TextEncoder().encode(JSON.stringify(slide(0))).buffer;
          return pullResponse(command, {
            kind: 'chunk', sequence: command.sequence, byteLength: payload.byteLength,
            done: true, payload,
          });
        }
        return pullResponse(command, { kind: 'accepted', command: command.kind });
      },
      forgetOrphaned: () => undefined,
      terminate: () => undefined,
    };
    const bridge = {
      request: async (build: (id: number) => PptxWorkerRequest) => {
        const request = build(requestId++);
        if (request.kind === 'parse') return { kind: 'presentationOpened', id: request.id, bootstrap };
        if (request.kind === 'openSlideSession') {
          slideSessionId = request.sessionId;
          return { ...request, kind: 'slideSessionOpened' as const };
        }
        throw new Error(`unexpected request ${request.kind}`);
      },
      transport: () => transport,
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'main', _bridge: bridge, _embeddedFontFaces: [],
      _embeddedFontAliases: new Map(), _embeddedFontAuthoredFamilies: new Map(),
      _destroyed: false, _layoutWaiters: new Set(),
      _layoutLifecycle: new ProgressiveLayoutLifecycle(),
      _layoutObservers: new ProgressiveLayoutObserverNotifier(),
    });
    const presentation = instance as unknown as PptxPresentation;
    const onComplete = vi.fn();
    const lifecycle = {
      firstPublication: deferred<void>(), published: false, deferred: false,
      settled: false, onComplete,
    };
    const policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null } as const;

    await (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer, resourcePolicy: typeof policy, useGoogleFonts: boolean,
        timeoutMs: undefined, onUsage: undefined, renderers: undefined,
        progressive: typeof lifecycle,
      ): Promise<void>;
    })._parse(new ArrayBuffer(4), policy, false, undefined, undefined, undefined, lifecycle);

    expect(slideSessionId).toBeGreaterThan(0);
    expect(presentation.layoutComplete).toBe(true);
    expect(presentation.availableSlideCount).toBe(1);
    expect(onComplete).not.toHaveBeenCalled();
  });

  it('waits for the authoritative final response in an actual one-slide worker parse', async () => {
    const finalResponse = deferred<RenderWorkerResponse>();
    const bootstrap = {
      slideCount: 1, slideWidth: 9144000, slideHeight: 6858000,
      defaultTextColor: null, majorFont: null, minorFont: null,
      hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    } as const;
    const facts = {
      index: 0, partName: 'ppt/slides/slide1.xml', notes: null,
      hidden: false, mediaElements: [],
    } as const;
    const bridge = {
      request: (build: (id: number) => RenderWorkerRequest) => {
        const request = build(91);
        expect(request.kind).toBe('parse');
        return finalResponse.promise;
      },
      post: vi.fn(), terminate: vi.fn(),
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    Object.assign(instance, {
      _mode: 'worker', _bridge: bridge, _destroyed: false,
      _layoutWaiters: new Set(), _availableSlideCount: 0,
      _layoutLifecycle: new ProgressiveLayoutLifecycle(),
      _layoutObservers: new ProgressiveLayoutObserverNotifier(),
      _parseRequestId: null, _progressive: null, _metrics: null,
    });
    const presentation = instance as unknown as PptxPresentation;
    const onComplete = vi.fn();
    const lifecycle = {
      firstPublication: deferred<void>(), published: false, deferred: false,
      settled: false, onComplete,
    };
    const policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null } as const;
    let released = false;
    const parsing = (presentation as unknown as {
      _parse(
        buffer: ArrayBuffer, resourcePolicy: typeof policy, useGoogleFonts: boolean,
        timeoutMs: undefined, onUsage: undefined, renderers: undefined,
        progressive: typeof lifecycle,
      ): Promise<void>;
    })._parse(new ArrayBuffer(4), policy, false, undefined, undefined, undefined, lifecycle)
      .then(() => { released = true; });
    await Promise.resolve();
    (presentation as unknown as {
      _onWorkerLayoutPush(response: RenderWorkerResponse): void;
    })._onWorkerLayoutPush({
      kind: 'presentationLayoutPartial', forId: 91, bootstrap,
      availableSlides: 1, slide: facts, fontPreloadNames: [],
    });
    await vi.waitFor(() => expect(bridge.post).toHaveBeenCalledTimes(1));
    expect(released).toBe(false);

    finalResponse.resolve({
      kind: 'presentationReady', id: 91,
      preflight: { ...bootstrap, slides: [facts], fontPreloadNames: [] },
    });
    await parsing;
    expect(presentation.layoutComplete).toBe(true);
    expect(onComplete).not.toHaveBeenCalled();
  });
});
