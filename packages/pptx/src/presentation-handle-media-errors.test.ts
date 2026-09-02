import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createPresentationHandle, type PresentOptions } from './presentation-handle';
import type { MediaElement } from './types';

type Listener = (event: Event) => void;

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function mediaElement(kind: 'audio' | 'video' = 'audio'): MediaElement {
  return {
    type: 'media',
    x: 0,
    y: 0,
    width: 914_400,
    height: 914_400,
    rotation: 0,
    flipH: false,
    flipV: false,
    mediaKind: kind,
    posterPath: '',
    posterMimeType: '',
    mediaPath: kind === 'audio' ? 'ppt/media/clip.mp3' : 'ppt/media/clip.mp4',
    mimeType: kind === 'audio' ? 'audio/mpeg' : 'video/mp4',
  };
}

function makeContext() {
  const gradient = { addColorStop: vi.fn() };
  return {
    setTransform: vi.fn(),
    drawImage: vi.fn(),
    save: vi.fn(),
    restore: vi.fn(),
    beginPath: vi.fn(),
    arc: vi.fn(),
    fill: vi.fn(),
    fillRect: vi.fn(),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    quadraticCurveTo: vi.fn(),
    closePath: vi.fn(),
    createLinearGradient: vi.fn(() => gradient),
    measureText: vi.fn((text: string) => ({ width: text.length * 6 })),
    fillText: vi.fn(),
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    fillStyle: '',
    shadowColor: '',
    shadowBlur: 0,
  };
}

function makeCanvas(ctx = makeContext()) {
  const listeners = new Map<string, Listener[]>();
  return {
    width: 960,
    height: 540,
    style: { cursor: '' },
    getContext: vi.fn(() => ctx),
    getBoundingClientRect: vi.fn(() => ({ left: 0, top: 0, width: 960, height: 540 })),
    addEventListener: vi.fn((type: string, listener: Listener) => {
      const list = listeners.get(type) ?? [];
      list.push(listener);
      listeners.set(type, list);
    }),
    removeEventListener: vi.fn((type: string, listener: Listener) => {
      listeners.set(type, (listeners.get(type) ?? []).filter((item) => item !== listener));
    }),
    setPointerCapture: vi.fn(),
    releasePointerCapture: vi.fn(),
    dispatch(type: string, event: Event) {
      for (const listener of listeners.get(type) ?? []) listener(event);
    },
  };
}

function makeMedia() {
  const listeners = new Map<string, Listener[]>();
  const media = {
    src: '',
    preload: '',
    playsInline: false,
    paused: true,
    duration: Number.NaN,
    currentTime: 0,
    readyState: 0,
    networkState: 0,
    error: null as MediaError | null,
    play: vi.fn(() => Promise.resolve()),
    pause: vi.fn(),
    load: vi.fn(),
    canPlayType: vi.fn(() => 'probably'),
    removeAttribute: vi.fn(),
    addEventListener: vi.fn((type: string, listener: Listener) => {
      const list = listeners.get(type) ?? [];
      list.push(listener);
      listeners.set(type, list);
    }),
    removeEventListener: vi.fn((type: string, listener: Listener) => {
      listeners.set(type, (listeners.get(type) ?? []).filter((item) => item !== listener));
    }),
    dispatch(type: string) {
      for (const listener of listeners.get(type) ?? []) listener(new Event(type));
    },
  };
  return media;
}

function installDom(media = makeMedia()) {
  const baseContext = makeContext();
  const baseCanvas = makeCanvas(baseContext);
  vi.stubGlobal('document', {
    createElement(tag: string) {
      return tag === 'canvas' ? baseCanvas : media;
    },
  });
  vi.stubGlobal('URL', {
    createObjectURL: vi.fn(() => 'blob:media'),
    revokeObjectURL: vi.fn(),
  });
  vi.stubGlobal('requestAnimationFrame', vi.fn(() => 1));
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  return { media, baseCanvas, baseContext };
}

function options(overrides: Partial<PresentOptions> = {}): PresentOptions {
  return {
    width: 960,
    height: 540,
    slideWidthEmu: 9_144_000,
    fetchMedia: vi.fn(async () => new Blob(['media'], { type: 'audio/mpeg' })),
    fetchImage: vi.fn(),
    drawBase: vi.fn(async () => {}),
    ...overrides,
  };
}

beforeEach(() => {
  vi.clearAllMocks();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('presentation media failure reporting', () => {
  it('keeps overlay drawing and hit testing in logical slide coordinates when the backing store is clamped', async () => {
    const { media } = installDom();
    const canvasContext = makeContext();
    const canvas = makeCanvas(canvasContext);
    canvas.width = 1_000;
    canvas.height = 500;
    canvas.getBoundingClientRect.mockReturnValue({
      left: 0,
      top: 0,
      width: 2_000,
      height: 1_000,
    });
    const rightSideMedia = {
      ...mediaElement('video'),
      x: 7_315_200,
      y: 1_828_800,
    };

    const handle = await createPresentationHandle(
      canvas as unknown as HTMLCanvasElement,
      [rightSideMedia],
      options({ width: 2_000, height: 1_000 }),
    );

    // Requested DPR would have been 2, but the allocated buffer represents
    // only 0.5 physical pixels per logical pixel after clamping.
    expect(canvasContext.setTransform).toHaveBeenCalledWith(0.5, 0, 0, 0.5, 0, 0);
    canvas.dispatch('pointerdown', {
      clientX: 1_700,
      clientY: 450,
      pointerId: 1,
      preventDefault: vi.fn(),
    } as unknown as PointerEvent);
    expect(media.play).toHaveBeenCalledTimes(1);
    handle.destroy();
  });

  it('rejects initial fetchMedia failures without also calling onError', async () => {
    installDom();
    const onError = vi.fn();
    const canvasContext = makeContext();
    const canvas = makeCanvas(canvasContext);

    const promise = createPresentationHandle(
      canvas as unknown as HTMLCanvasElement,
      [mediaElement()],
      options({
        fetchMedia: vi.fn(async () => {
          throw new Error('archive read failed');
        }),
        onError,
      }),
    );

    await expect(promise).rejects.toThrow(/ppt\/media\/clip\.mp3.*archive read failed/);
    expect(onError).not.toHaveBeenCalled();
    expect(canvasContext.fillText).not.toHaveBeenCalledWith(
      'Media unavailable',
      expect.any(Number),
      expect.any(Number),
    );
  });

  it('registers readiness/error listeners, explicitly loads media, and reports decode errors', async () => {
    const { media } = installDom();
    const onError = vi.fn();

    const handle = await createPresentationHandle(
      makeCanvas() as unknown as HTMLCanvasElement,
      [mediaElement('video')],
      options({
        fetchMedia: vi.fn(async () => new Blob(['video'], { type: 'video/mp4' })),
        onError,
      }),
    );

    expect(media.addEventListener).toHaveBeenCalledWith('loadedmetadata', expect.any(Function));
    expect(media.addEventListener).toHaveBeenCalledWith('canplay', expect.any(Function));
    expect(media.addEventListener).toHaveBeenCalledWith('error', expect.any(Function));
    expect(media.load).toHaveBeenCalledTimes(1);

    media.error = { code: 4, message: 'unsupported source' } as MediaError;
    media.dispatch('error');

    expect(onError).toHaveBeenCalledTimes(1);
    const error = onError.mock.calls[0][0] as Error;
    expect(error.message).toContain('ppt/media/clip.mp4');
    expect(error.message).toContain('unsupported source');
    expect(error.message).toContain('readyState=0');
    expect(error.message).toContain('canPlayType=probably');
    handle.destroy();
  });

  it('reports play() promise rejections instead of discarding them', async () => {
    const { media } = installDom();
    const onError = vi.fn();
    media.play.mockRejectedValueOnce(new DOMException('codec unavailable', 'NotSupportedError'));

    const handle = await createPresentationHandle(
      makeCanvas() as unknown as HTMLCanvasElement,
      [mediaElement()],
      options({ onError }),
    );
    handle.play();
    await Promise.resolve();

    expect(onError).toHaveBeenCalledTimes(1);
    const error = onError.mock.calls[0][0] as Error;
    expect(error.message).toContain('play');
    expect(error.message).toContain('NotSupportedError');
    expect(error.message).toContain('codec unavailable');
    handle.destroy();
  });

  it('does not report a late play rejection after the handle is destroyed', async () => {
    const { media } = installDom();
    const onError = vi.fn();
    const pendingPlay = deferred<void>();
    media.play.mockReturnValueOnce(pendingPlay.promise);

    const handle = await createPresentationHandle(
      makeCanvas() as unknown as HTMLCanvasElement,
      [mediaElement()],
      options({ onError }),
    );
    handle.play();
    handle.destroy();
    pendingPlay.reject(new DOMException('interrupted by teardown', 'AbortError'));
    await Promise.resolve();

    expect(onError).not.toHaveBeenCalled();
  });

  it('keeps the rendered poster visible while media extraction is pending', async () => {
    installDom();
    const pending = deferred<Blob>();
    const canvasContext = makeContext();
    const promise = createPresentationHandle(
      makeCanvas(canvasContext) as unknown as HTMLCanvasElement,
      [mediaElement('video')],
      options({ fetchMedia: vi.fn(() => pending.promise) }),
    );
    await Promise.resolve();

    expect(canvasContext.fillText).not.toHaveBeenCalledWith(
      expect.stringContaining('Loading'),
      expect.any(Number),
      expect.any(Number),
    );

    pending.resolve(new Blob(['video'], { type: 'video/mp4' }));
    const handle = await promise;
    handle.destroy();
  });
});
