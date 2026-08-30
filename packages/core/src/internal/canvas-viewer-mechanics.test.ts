import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  CanvasViewerErrorRouter,
  readBoundedNativeTextSelection,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from './canvas-viewer-mechanics.js';

afterEach(() => vi.restoreAllMocks());

describe('readBoundedNativeTextSelection', () => {
  function fixture(endInside = true, endOnSelectionSurface = endInside) {
    const insideStart = {} as Node;
    const insideEnd = {} as Node;
    const outside = {} as Node;
    const textNodes = [
      { nodeType: 3, data: 'selected ', childNodes: [] },
      { nodeType: 3, data: 'text', childNodes: [] },
    ] as unknown as Text[];
    const runs = [
      { dataset: { run: '0' }, childNodes: [textNodes[0]] },
      { dataset: { run: '1' }, childNodes: [textNodes[1]] },
    ] as unknown as HTMLElement[];
    const range = {
      startContainer: insideStart,
      endContainer: endInside ? insideEnd : outside,
      intersectsNode: (node: Node) => node === runs[0] || node === runs[1],
      comparePoint: () => 0,
    } as unknown as Range;
    const surface = {
      contains: (node: Node) =>
        node === insideStart || endOnSelectionSurface && node === insideEnd ||
        textNodes.includes(node as Text),
    } as unknown as HTMLElement;
    const root = {
      contains: (node: Node) =>
        node === insideStart || node === insideEnd || textNodes.includes(node as Text),
      matches: () => false,
      querySelectorAll: (selector: string) =>
        selector === '[data-ooxml-selection-surface]' ? [surface] : runs,
    } as unknown as HTMLElement;
    const selection = {
      isCollapsed: false,
      rangeCount: 1,
      getRangeAt: () => range,
      toString: () => 'selected text',
    } as unknown as Selection;
    return { root, runs, selection, textNodes };
  }

  it('returns bounded text and detached run locators', () => {
    const { root, selection } = fixture();
    expect(readBoundedNativeTextSelection(
      root,
      selection,
      (run) => ({ run: Number(run.dataset.run) }),
      { maxChars: 8, maxLocators: 1 },
    )).toEqual({
      text: 'selected',
      locators: [{ run: 0 }],
      truncated: true,
      truncationReasons: ['text', 'runs'],
      textCharacters: 8,
      maxTextCharacters: 8,
      maxLocators: 1,
    });
  });

  it('rejects a range crossing outside the Viewer root', () => {
    const { root, selection } = fixture(false);
    expect(readBoundedNativeTextSelection(root, selection, () => ({ run: 0 }))).toBeNull();
  });

  it('rejects Viewer chrome text outside the tagged selection surface', () => {
    const { root, selection } = fixture(true, false);
    expect(readBoundedNativeTextSelection(root, selection, () => ({ run: 0 }))).toBeNull();
  });

  it('preserves Unicode boundaries and validates public resource limits', () => {
    const { root, selection, textNodes } = fixture();
    Object.assign(textNodes[0], { data: '\ud83d\ude00x' });
    Object.assign(textNodes[1], { data: '' });
    expect(readBoundedNativeTextSelection(
      root, selection, (run) => ({ run: Number(run.dataset.run) }), { maxChars: 1 },
    )?.text).toBe('');
    expect(readBoundedNativeTextSelection(
      root, selection, (run) => ({ run: Number(run.dataset.run) }), { maxChars: 2 },
    )?.text).toBe('\ud83d\ude00');
    expect(() => readBoundedNativeTextSelection(
      root, selection, () => ({ run: 0 }), { maxChars: Number.NaN },
    )).toThrow(/maxTextCharacters/);
  });

  it('never materializes Selection text or includes untagged content between surfaces', () => {
    const { root, selection, textNodes } = fixture();
    Object.assign(textNodes[0], { data: 'public-a ' });
    Object.assign(textNodes[1], { data: 'public-b' });
    Object.assign(selection, {
      toString: () => { throw new Error('must not materialize unbounded or untagged text'); },
    });

    expect(readBoundedNativeTextSelection(
      root, selection, (run) => ({ run: Number(run.dataset.run) }), { maxChars: 8 },
    )).toMatchObject({ text: 'public-a', truncated: true, truncationReasons: ['text'] });
  });

  it('extracts only the selected offsets from a tagged run', () => {
    const { root, runs, selection, textNodes } = fixture();
    Object.assign(textNodes[0], { data: 'before selected after' });
    Object.assign(textNodes[1], { data: '' });
    const range = selection.getRangeAt(0);
    Object.assign(range, {
      startContainer: textNodes[0],
      startOffset: 7,
      endContainer: textNodes[0],
      endOffset: 15,
      intersectsNode: (node: Node) => node === runs[0],
    });

    expect(readBoundedNativeTextSelection(
      root, selection, (run) => ({ run: Number(run.dataset.run) }),
    )?.text).toBe('selected');
  });
});

describe('resolveCanvasViewerMode', () => {
  it('uses the borrowed engine mode and rejects only an explicit conflict', () => {
    const engine = { mode: 'worker' as const };
    expect(resolveCanvasViewerMode('Viewer', undefined, engine)).toBe('worker');
    expect(resolveCanvasViewerMode('Viewer', 'worker', engine)).toBe('worker');
    expect(() => resolveCanvasViewerMode('Viewer', 'main', engine)).toThrow(
      "Viewer: opts.mode='main' conflicts with the borrowed engine's mode='worker'",
    );
  });

  it('defaults a self-loading viewer to main mode', () => {
    expect(resolveCanvasViewerMode('Viewer', undefined, undefined)).toBe('main');
    expect(resolveCanvasViewerMode('Viewer', 'worker', undefined)).toBe('worker');
  });
});

describe('StaticCanvasRenderDispatcher', () => {
  it('acquires the bitmap context once and avoids redundant backing-store resets', () => {
    const bitmapContext = { transferFromImageBitmap: vi.fn() };
    let width = 10;
    let height = 20;
    const setWidth = vi.fn((value: number) => { width = value; });
    const setHeight = vi.fn((value: number) => { height = value; });
    const canvas = {
      get width() { return width; },
      set width(value: number) { setWidth(value); },
      get height() { return height; },
      set height(value: number) { setHeight(value); },
      style: {},
      getContext: vi.fn(() => bitmapContext),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmap(dispatcher.begin(), bitmap)).toBe(true);
    expect(canvas.getContext).toHaveBeenCalledOnce();
    expect(setWidth).not.toHaveBeenCalled();
    expect(setHeight).not.toHaveBeenCalled();
  });

  it('closes a stale worker bitmap instead of committing it', () => {
    const transferFromImageBitmap = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({ transferFromImageBitmap })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const stale = dispatcher.begin();
    dispatcher.begin();
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmap(stale, bitmap)).toBe(false);
    expect(bitmap.close).toHaveBeenCalledOnce();
    expect(transferFromImageBitmap).not.toHaveBeenCalled();
  });

  it('closes the worker bitmap when transfer fails', () => {
    const failure = new Error('context lost');
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({
        transferFromImageBitmap: () => { throw failure; },
      })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(() => dispatcher.commitBitmap(dispatcher.begin(), bitmap)).toThrow(failure);
    expect(bitmap.close).toHaveBeenCalledOnce();
  });

  it('commits and closes a worker bitmap through a 2D context', () => {
    const drawImage = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({ drawImage })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, false);
    const bitmap = { width: 30, height: 40, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmapTo2d(dispatcher.begin(), bitmap, {
      cssWidth: 15,
      cssHeight: 20,
    })).toBe(true);
    expect(drawImage).toHaveBeenCalledWith(bitmap, 0, 0);
    expect(bitmap.close).toHaveBeenCalledOnce();
    expect(canvas.width).toBe(30);
    expect(canvas.height).toBe(40);
    expect(canvas.style.width).toBe('15px');
    expect(canvas.style.height).toBe('20px');
  });

  it('closes a stale 2D worker bitmap without drawing it', () => {
    const drawImage = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({ drawImage })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, false);
    const stale = dispatcher.begin();
    dispatcher.begin();
    const bitmap = { width: 30, height: 40, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmapTo2d(stale, bitmap)).toBe(false);
    expect(bitmap.close).toHaveBeenCalledOnce();
    expect(drawImage).not.toHaveBeenCalled();
  });
});

describe('CanvasViewerErrorRouter', () => {
  it('normalizes failures and becomes silent after close', () => {
    const onError = vi.fn();
    const router = new CanvasViewerErrorRouter('TestViewer', onError);
    router.report('failure');
    router.close();
    router.report(new Error('late'));
    expect(onError).toHaveBeenCalledOnce();
    expect(onError.mock.calls[0][0]).toEqual(new Error('failure'));
  });

  it('delivers one Error identity once and respects an explicit callback owner', () => {
    const onError = vi.fn();
    const router = new CanvasViewerErrorRouter('TestViewer', onError);
    const repeated = new Error('repeated');
    router.report(repeated);
    router.report(repeated);

    const explicitlyHandled = new Error('handled elsewhere');
    router.markHandled(explicitlyHandled);
    router.report(explicitlyHandled);

    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(repeated);
  });

  it('keeps a terminal failure on the active Promise channel', async () => {
    const onError = vi.fn();
    const router = new CanvasViewerErrorRouter('TestViewer', onError);
    const failure = new Error('layout failed');
    let reject!: (error: Error) => void;
    const pending = new Promise<void>((_resolve, rejectPromise) => { reject = rejectPromise; });
    const awaited = router.ownBackgroundLifecycle(() => pending);

    router.reportBackground(failure);
    reject(failure);

    await expect(awaited).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
  });

  it('does not hide an unrelated background failure while a Promise is pending', async () => {
    const onError = vi.fn();
    const router = new CanvasViewerErrorRouter('TestViewer', onError);
    let resolve!: () => void;
    const pending = new Promise<void>((resolvePromise) => { resolve = resolvePromise; });
    const awaited = router.ownAwaitable(() => pending);
    const unrelated = new Error('unrelated render failure');

    router.reportBackground(unrelated);
    expect(onError).toHaveBeenCalledWith(unrelated);

    resolve();
    await awaited;
  });
});

class Resource {
  destroyed = false;
  destroy(): void { this.destroyed = true; }
}

function deferred<T>(): { promise: Promise<T>; resolve(value: T): void } {
  let resolvePromise: (value: T) => void = () => undefined;
  const promise = new Promise<T>((resolve) => { resolvePromise = resolve; });
  return { promise, resolve: resolvePromise };
}

describe('TerminalResourceOwner', () => {
  it('atomically replaces and destroys an owned resource', async () => {
    const first = new Resource();
    const second = new Resource();
    const owner = new TerminalResourceOwner<Resource>('test', first, true);

    await expect(owner.replace(async () => second)).resolves.toBe(second);
    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);
    owner.close();
    expect(second.destroyed).toBe(true);
  });

  it('permanently rejects acquisition after close without invoking the loader', async () => {
    const owner = new TerminalResourceOwner<Resource>('test');
    let invoked = false;
    owner.close();

    await expect(owner.replace(async () => {
      invoked = true;
      return new Resource();
    })).rejects.toThrow('test is closed');
    expect(invoked).toBe(false);
  });

  it('destroys a candidate that resolves after close', async () => {
    const pending = deferred<Resource>();
    const owner = new TerminalResourceOwner<Resource>('test');
    const replacing = owner.replace(() => pending.promise);
    owner.close();
    const candidate = new Resource();
    pending.resolve(candidate);

    await expect(replacing).rejects.toThrow('test is closed');
    expect(candidate.destroyed).toBe(true);
  });

  it('direct install supersedes and disposes a pending replacement candidate', async () => {
    const pending = deferred<Resource>();
    const owner = new TerminalResourceOwner<Resource>('test');
    const replacing = owner.replace(() => pending.promise);
    const installed = new Resource();
    owner.install(installed);
    const late = new Resource();
    pending.resolve(late);

    await expect(replacing).resolves.toBeNull();
    expect(owner.current).toBe(installed);
    expect(installed.destroyed).toBe(false);
    expect(late.destroyed).toBe(true);
  });

  it('does not destroy a borrowed initial resource', () => {
    const borrowed = new Resource();
    const owner = new TerminalResourceOwner<Resource>('test', borrowed, false);
    owner.close();
    expect(borrowed.destroyed).toBe(false);
  });

  it('detaches a throwing resource permanently when closed', () => {
    const throwing = { destroy: vi.fn(() => { throw new Error('dispose failed'); }) };
    const owner = new TerminalResourceOwner('test', throwing, true);

    expect(() => owner.close()).not.toThrow();
    expect(owner.current).toBeNull();
    owner.close();
    expect(throwing.destroy).toHaveBeenCalledOnce();
  });

  it('commits a replacement even when disposal of the previous resource throws', async () => {
    const throwing = { destroy: vi.fn(() => { throw new Error('dispose failed'); }) };
    const next = new Resource();
    const owner = new TerminalResourceOwner<Resource | typeof throwing>('test', throwing, true);

    await expect(owner.replace(async () => next)).resolves.toBe(next);
    expect(owner.current).toBe(next);
    expect(throwing.destroy).toHaveBeenCalledOnce();
  });

  it('lets terminal close win in the microtask after replacement installation', async () => {
    const first = new Resource();
    const next = new Resource();
    const owner = new TerminalResourceOwner<Resource>('test', first, true);

    const replacement = owner.replace(async () => next, () => {
      queueMicrotask(() => owner.close());
    });

    await expect(replacement).resolves.toBe(next);
    expect(first.destroyed).toBe(true);
    expect(next.destroyed).toBe(true);
    expect(owner.current).toBeNull();
  });

  it('preserves stale and closed outcomes when candidate disposal throws', async () => {
    const closedPending = deferred<{ destroy(): void }>();
    const closedOwner = new TerminalResourceOwner<{ destroy(): void }>('closed');
    const closed = closedOwner.replace(() => closedPending.promise);
    closedOwner.close();
    closedPending.resolve({ destroy: () => { throw new Error('dispose failed'); } });
    await expect(closed).rejects.toThrow('closed is closed');

    const stalePending = deferred<{ destroy(): void }>();
    const staleOwner = new TerminalResourceOwner<{ destroy(): void }>('stale');
    const stale = staleOwner.replace(() => stalePending.promise);
    staleOwner.install({ destroy: () => undefined });
    stalePending.resolve({ destroy: () => { throw new Error('dispose failed'); } });
    await expect(stale).resolves.toBeNull();
  });
});
