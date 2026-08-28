import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import { PptxScrollViewer } from './scroll-viewer.js';
import {
  FakePptxEngine,
  installDom,
  makeContainer,
  type FakeEl,
} from './scroll-viewer-test-dom.js';

const WIDTH = 9144000;
const HEIGHT = 5143500;

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('PptxScrollViewer progressive layout', () => {
  it('forwards lifecycle options and keeps the final scroll extent from first paint', async () => {
    installDom();
    const engine = new FakePptxEngine(8, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const load = vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const container = makeContainer();
    const viewer = new PptxScrollViewer(container as unknown as HTMLElement, {
      progressiveLayout: true,
    });

    const loading = viewer.load('deck.pptx');
    await vi.waitFor(() => expect(viewer.slideCount).toBe(8));
    const spacer = container.children[0]!.children[0]!.children[0]!;
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(container.clientHeight);
    expect(viewer.slideCount).toBe(8);
    expect(viewer.availableSlideCount).toBe(1);
    expect(viewer.layoutComplete).toBe(false);
    expect(load).toHaveBeenCalledWith('deck.pptx', expect.objectContaining({
      progressiveLayout: true,
    }));

    engine.setLayoutProgress(8, true);
    await loading;
    await expect(viewer.waitUntilLayoutComplete()).resolves.toBeUndefined();
    expect(viewer.layoutComplete).toBe(true);
    viewer.destroy();
  });

  it('shows a per-slide loading state and emits the completion transition', async () => {
    installDom();
    const engine = new FakePptxEngine(4, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const onVisibleSlideChange = vi.fn();
    const container = makeContainer();
    const viewer = PptxScrollViewer.fromPresentation(
      container as unknown as HTMLElement,
      engine.asPres(),
      { overscan: 3, onVisibleSlideChange },
    ) as PptxScrollViewer;
    const slots = (viewer as unknown as { _slots: Map<number, { loadingLayer: FakeEl }> })._slots;

    expect(slots.get(1)?.loadingLayer.style.display).toBe('flex');
    expect(onVisibleSlideChange).toHaveBeenLastCalledWith(0, 4, false);

    engine.setLayoutProgress(4, true);
    await vi.waitFor(() => expect(slots.get(1)?.loadingLayer.style.display).toBe('none'));
    expect(onVisibleSlideChange).toHaveBeenLastCalledWith(0, 4, true);
    viewer.destroy();
  });

  it.each(['main', 'worker'] as const)(
    'discovers later comments and rebuilds a borrowed %s-mode surface',
    async (mode) => {
      installDom();
      const engine = new FakePptxEngine(3, WIDTH, HEIGHT, mode);
      engine.commentsBySlide = [[], [], [{
        id: 'later-comment', author: 'Reviewer', text: 'Later', x: 100, y: 100,
      }]];
      engine.setLayoutProgress(1, false);
      const container = makeContainer();
      const viewer = PptxScrollViewer.fromPresentation(
        container as unknown as HTMLElement,
        engine.asPres(),
        { comments: true, overscan: 0 },
      ) as PptxScrollViewer;
      const state = viewer as unknown as {
        _hasComments: boolean;
        _slots: Map<number, {
          canvas: FakeEl;
          dispatcher: unknown;
          wrapper: FakeEl;
          commentMarkerLayer: FakeEl | null;
          commentMargin: FakeEl | null;
        }>;
      };
      const spacer = container.children[0]!.children[0]!.children[0]!;
      const scrollHost = container.children[0]!.children[0]!;
      scrollHost.scrollTop = 19;
      const openingScale = viewer.getScale();
      const openingHeight = spacer.style.height;
      const openingWidth = parseFloat(spacer.style.width);
      const openingSlots = new Map(state._slots);
      const openingLeft = state._slots.get(0)!.wrapper.style.left;
      expect(state._hasComments).toBe(false);
      expect([...state._slots.values()].every((slot) =>
        slot.commentMarkerLayer !== null && slot.commentMargin !== null)).toBe(true);

      engine.setLayoutProgress(3, true);
      await vi.waitFor(() => expect(state._hasComments).toBe(true));
      expect([...state._slots.values()].every((slot) =>
        slot.commentMarkerLayer !== null && slot.commentMargin !== null)).toBe(true);
      expect(viewer.getScale()).toBe(openingScale);
      expect(spacer.style.height).toBe(openingHeight);
      expect(scrollHost.scrollTop).toBe(19);
      expect(parseFloat(spacer.style.width)).toBeGreaterThan(openingWidth);
      for (const [index, openingSlot] of openingSlots) {
        expect(state._slots.get(index)?.canvas).toBe(openingSlot.canvas);
        expect(state._slots.get(index)?.dispatcher).toBe(openingSlot.dispatcher);
        expect(state._slots.get(index)?.wrapper.style.left).toBe(openingLeft);
      }
      scrollHost.scrollTop = 10_000;
      scrollHost.dispatch('scroll');
      expect(state._slots.get(2)?.wrapper.style.left).toBe(openingLeft);
      viewer.relayout();
      expect(state._slots.get(2)?.wrapper.style.left).toBe(openingLeft);
      viewer.destroy();
    },
  );

  it('keeps authored slide screen position stable when a left review rail appears', async () => {
    installDom();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT, 'worker');
    engine.commentsBySlide = [[], [], [{
      id: 'later-left', author: 'Reviewer', text: 'Later', x: 100, y: 100,
    }]];
    engine.setLayoutProgress(1, false);
    const container = makeContainer();
    const viewer = PptxScrollViewer.fromPresentation(
      container as unknown as HTMLElement,
      engine.asPres(),
      { comments: { side: 'left' }, overscan: 0 },
    ) as PptxScrollViewer;
    const state = viewer as unknown as {
      _hasComments: boolean;
      _reviewOriginPx: number;
      _slots: Map<number, { wrapper: FakeEl }>;
    };
    const scrollHost = container.children[0]!.children[0]!;
    const spacer = scrollHost.children[0]!;
    let browserScrollLeft = 0;
    Object.defineProperty(scrollHost, 'scrollLeft', {
      configurable: true,
      get: () => browserScrollLeft,
      set: (value: number) => {
        const max = Math.max(0, parseFloat(spacer.style.width) - scrollHost.clientWidth);
        browserScrollLeft = Math.min(max, Math.max(0, value));
      },
    });
    const openingLeft = state._slots.get(0)!.wrapper.style.left;
    const authoredLeft = Number(openingLeft.match(/calc\(([-\d.]+)px/)?.[1]);
    const openingScreenX = authoredLeft + state._reviewOriginPx - scrollHost.scrollLeft;

    engine.setLayoutProgress(3, true);
    await vi.waitFor(() => expect(state._hasComments).toBe(true));
    expect(state._slots.get(0)!.wrapper.style.left).toBe(openingLeft);
    expect(state._reviewOriginPx).toBeGreaterThan(0);
    expect(scrollHost.scrollLeft).toBe(state._reviewOriginPx);
    expect(authoredLeft + state._reviewOriginPx - scrollHost.scrollLeft).toBe(openingScreenX);

    scrollHost.scrollTop = 10_000;
    scrollHost.dispatch('scroll');
    expect(state._slots.get(2)!.wrapper.style.left).toBe(openingLeft);
    viewer.relayout();
    expect(state._slots.get(2)!.wrapper.style.left).toBe(openingLeft);
    expect(authoredLeft + state._reviewOriginPx - scrollHost.scrollLeft).toBe(openingScreenX);
    viewer.destroy();
  });

  it.each(['main', 'worker'] as const)(
    'discovers later comments after a self-loaded %s-mode presentation publishes them',
    async (mode) => {
      installDom();
      const engine = new FakePptxEngine(3, WIDTH, HEIGHT, mode);
      engine.commentsBySlide = [[], [], [{
        id: 'later-comment', author: 'Reviewer', text: 'Later', x: 100, y: 100,
      }]];
      engine.setLayoutProgress(1, false);
      vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
      const viewer = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
        comments: true,
        progressiveLayout: true,
        mode,
        overscan: 0,
      });
      await viewer.load('deck.pptx');
      const state = viewer as unknown as {
        _hasComments: boolean;
        _slots: Map<number, { commentMarkerLayer: FakeEl | null }>;
      };
      expect(state._hasComments).toBe(false);

      engine.setLayoutProgress(3, true);
      await vi.waitFor(() => expect(state._hasComments).toBe(true));
      expect([...state._slots.values()].every((slot) => slot.commentMarkerLayer !== null)).toBe(true);
      viewer.destroy();
    },
  );

  it.each(['main', 'worker'] as const)(
    'waits for a later comment slide in %s mode instead of treating it as absent',
    async (mode) => {
      installDom();
      const engine = new FakePptxEngine(3, WIDTH, HEIGHT, mode);
      engine.commentsBySlide = [[], [], [{
        id: 'later-comment', author: 'Reviewer', text: 'Later', x: 100, y: 100,
      }]];
      engine.setLayoutProgress(1, false);
      const viewer = PptxScrollViewer.fromPresentation(
        makeContainer() as unknown as HTMLElement,
        engine.asPres(),
      ) as PptxScrollViewer;

      let settled = false;
      const navigation = viewer.goToComment(2, 0).then((result) => {
        settled = true;
        return result;
      });
      await Promise.resolve();
      expect(settled).toBe(false);

      engine.setLayoutProgress(3, true);
      await expect(navigation).resolves.toBe(true);
      viewer.destroy();
    },
  );

  it.each(['main', 'worker'] as const)(
    'does not retain %s render work for unavailable slots that scroll out before publication',
    async (mode) => {
    installDom();
    const engine = new FakePptxEngine(8, WIDTH, HEIGHT, mode);
    engine.setLayoutProgress(1, false);
    const container = makeContainer(200, 400);
    const viewer = PptxScrollViewer.fromPresentation(
      container as unknown as HTMLElement,
      engine.asPres(),
      { overscan: 3 },
    ) as PptxScrollViewer;
    const scrollHost = container.children[0]!.children[0]!;

    const calls = mode === 'main' ? engine.renderCalls : engine.bitmapCalls;
    expect(calls.map((call) => call.slide)).toEqual([0]);
    scrollHost.scrollTop = 10_000;
    scrollHost.dispatch('scroll');
    expect(viewer.mountedSlideIndicesForTest()).not.toContain(1);

    engine.setLayoutProgress(2, false);
    await Promise.resolve();
    expect(calls.some((call) => call.slide === 1)).toBe(false);

    viewer.scrollToSlide(1);
    await vi.waitFor(() => expect(calls.filter((call) => call.slide === 1)).toHaveLength(1));
    viewer.destroy();
    },
  );

  it('settles a superseded future-comment wait immediately', async () => {
    installDom();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.commentsBySlide = [[], [], [{
      id: 'later-comment', author: 'Reviewer', text: 'Later', x: 100, y: 100,
    }]];
    engine.setLayoutProgress(1, false);
    const viewer = PptxScrollViewer.fromPresentation(
      makeContainer() as unknown as HTMLElement,
      engine.asPres(),
    ) as PptxScrollViewer;

    const first = viewer.goToComment(2, 0);
    await expect(viewer.goToComment(0, 0)).resolves.toBe(false);
    await expect(first).resolves.toBe(false);
    viewer.destroy();
  });

  it('keeps a concurrent progressive failure on the canceled comment-navigation Promise', async () => {
    installDom();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.commentsBySlide = [[], [], [{
      id: 'later-comment', author: 'Reviewer', text: 'Later', x: 100, y: 100,
    }]];
    engine.setLayoutProgress(1, false);
    const onError = vi.fn();
    const viewer = PptxScrollViewer.fromPresentation(
      makeContainer() as unknown as HTMLElement,
      engine.asPres(),
      { onError },
    ) as PptxScrollViewer;
    const failure = new Error('layout failed during comment cancellation');

    const first = viewer.goToComment(2, 0);
    void viewer.goToComment(0, 0);
    engine.setLayoutProgress(1, false, failure);

    await expect(first).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('does not lose a layout failure when clearFind supersedes a progressive search', async () => {
    installDom();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const onError = vi.fn();
    const viewer = PptxScrollViewer.fromPresentation(
      makeContainer() as unknown as HTMLElement,
      engine.asPres(),
      { onError },
    ) as PptxScrollViewer;
    const failure = new Error('layout failed during find cancellation');

    const find = viewer.findText('needle');
    viewer.clearFind();
    engine.setLayoutProgress(1, false, failure);

    await expect(find).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });
});
