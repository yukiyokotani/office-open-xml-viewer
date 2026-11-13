import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  FakePptxEngine,
  installDom,
  makeBorrowedPptxScrollViewer,
  makeContainer,
  viewerState,
  type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { HiddenSlideMode } from './hidden';
import type { PptxScrollViewerOptions } from './scroll-viewer';
import type { PptxTextRunInfo } from './renderer';
import { parseRelativeSlideJump, resolveRelativeSlideJump } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

// Natural slide 200 × 120 px; with gap 10 the stride is 130 px and the first
// slide starts at the 10 px top padding.
const SLIDE_W_EMU = 9525 * 200;
const SLIDE_H_EMU = 9525 * 120;
const STRIDE = 130;
const LEADING = 10;

const NEEDLE: PptxTextRunInfo = {
  text: 'needle',
  inShapeX: 0,
  inShapeY: 0,
  w: 10,
  h: 12,
  fontSize: 12,
  font: '12px serif',
  shapeX: 0,
  shapeY: 0,
  shapeW: 100,
  shapeH: 50,
  rotation: 0,
};

interface SetupOptions {
  opts?: Partial<PptxScrollViewerOptions>;
  engine?: FakePptxEngine;
  /** Viewport height. 100 px shows one slide (overscan 0). */
  viewport?: number;
}

function setup(count: number, hidden: number[], mode: HiddenSlideMode | undefined, options: SetupOptions = {}) {
  installDom();
  const engine = options.engine ?? new FakePptxEngine(count, SLIDE_W_EMU, SLIDE_H_EMU);
  for (const i of hidden) engine.hiddenSlides.add(i);
  engine.feedTextRuns = [NEEDLE];
  const renderSlide = vi.spyOn(engine, 'renderSlide');
  const viewport = options.viewport ?? 100;
  const container = makeContainer(200, viewport);
  const onVisibleSlideChange = vi.fn();
  const v = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
    presentation: engine.asPres(),
    gap: 10,
    overscan: 0,
    paddingLeft: 0,
    paddingRight: 0,
    onVisibleSlideChange,
    ...(mode === undefined ? {} : { hiddenSlideMode: mode }),
    ...options.opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = viewport;
  scrollHost.clientWidth = 200;
  v.relayout();
  const slots = (v as unknown as { _slots: Map<number, { wrapper: FakeEl }> })._slots;
  const topOf = (slide: number) => slots.get(slide)?.wrapper.style.top;
  const scrollTo = (top: number) => {
    scrollHost.scrollTop = top;
    scrollHost.dispatch('scroll');
  };
  const spacer = (v as unknown as { _spacer: FakeEl })._spacer;
  const drawn = () => renderSlide.mock.calls.map((call) => call[1]);
  return { v, engine, scrollHost, renderSlide, drawn, onVisibleSlideChange, topOf, scrollTo, spacer };
}

describe("PptxScrollViewer hiddenSlideMode: 'skip' (#2181)", () => {
  it('leaves hidden slides out of the layout and reports absolute indices', () => {
    const { v, topOf, scrollTo, spacer, onVisibleSlideChange, drawn } = setup(6, [1, 2], 'skip');
    // Shown slides 0, 3, 4, 5 take layout positions 0..3: four slides of extent.
    expect(spacer.style.height).toBe(`${LEADING + 4 * STRIDE}px`);
    expect(topOf(0)).toBe(`${LEADING}px`);

    // The second laid-out slide is deck slide 3: indices stay absolute.
    scrollTo(LEADING + STRIDE);
    expect(topOf(3)).toBe(`${LEADING + STRIDE}px`);
    expect(v.topVisibleSlide).toBe(3);
    expect(onVisibleSlideChange).toHaveBeenLastCalledWith(3, 6, true);
    scrollTo(LEADING + 3 * STRIDE);
    expect(v.topVisibleSlide).toBe(5);
    expect(drawn()).not.toContain(1);
    expect(drawn()).not.toContain(2);
    expect(v.visibleSlideCount).toBe(4);
    expect(v.slideCount).toBe(6);
    v.destroy();
  });

  it('scrollToSlide and internal links to a hidden slide land on the next shown slide, else the previous', () => {
    const { v, engine } = setup(6, [1, 2, 5], 'skip');
    v.scrollToSlide(1);
    expect(v.topVisibleSlide).toBe(3);
    // No shown slide after 5 → the nearest one before it (clamped too).
    v.scrollToSlide(5);
    expect(v.topVisibleSlide).toBe(4);
    v.scrollToSlide(99);
    expect(v.topVisibleSlide).toBe(4);

    // An internal hyperlink to a hidden slide follows the same rule.
    v.scrollToSlide(0);
    engine.internalTargets.set('../slides/slide2.xml', 1);
    (v as unknown as { _onHyperlinkClick(target: unknown): void })._onHyperlinkClick({
      kind: 'internal', ref: '../slides/slide2.xml',
    });
    expect(v.topVisibleSlide).toBe(3);
    v.destroy();
  });

  it('a previousslide jump to a hidden slide lands on the shown slide before it', () => {
    const click = (v: unknown) => (v as { _onHyperlinkClick(target: unknown): void })._onHyperlinkClick({
      kind: 'internal', ref: 'ppaction://hlinkshowjump?jump=previousslide',
    });
    const { v, engine } = setup(6, [1, 2], 'skip');
    engine.resolveFn = (ref, current) => resolveRelativeSlideJump(parseRelativeSlideJump(ref)!, current, 6);
    v.scrollToSlide(3);
    expect(v.topVisibleSlide).toBe(3);
    click(v);
    expect(v.topVisibleSlide).toBe(0);
    v.destroy();

    // No shown slide before the target: the jump stays on the first shown slide.
    const first = setup(3, [0], 'skip');
    first.engine.resolveFn = (ref, current) => resolveRelativeSlideJump(parseRelativeSlideJump(ref)!, current, 3);
    expect(first.v.topVisibleSlide).toBe(1);
    click(first.v);
    expect(first.v.topVisibleSlide).toBe(1);
    first.v.destroy();
  });

  it('find does not match hidden slides', async () => {
    const { v } = setup(4, [1], 'skip');
    const matches = await v.findText('needle');
    expect(matches.map((match) => match.location.slide)).toEqual([0, 2, 3]);
    v.destroy();
  });

  it('goToComment on a hidden slide resolves false', async () => {
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.commentsBySlide = [[], [{ x: 0, y: 0, text: 'c' } as never], []];
    const { v } = setup(3, [1], 'skip', { engine });
    expect(await v.goToComment(1, 0)).toBe(false);
    v.destroy();
  });

  it('every slide hidden: lays out, navigates and searches as in show', async () => {
    const { v, topOf, spacer } = setup(3, [0, 1, 2], 'skip');
    expect(spacer.style.height).toBe(`${LEADING + 3 * STRIDE}px`);
    v.scrollToSlide(1);
    expect(v.topVisibleSlide).toBe(1);
    expect(topOf(1)).toBe(`${LEADING + STRIDE}px`);
    expect((await v.findText('needle')).length).toBe(3);
    v.destroy();
  });

  it('progressive layout: a slide drops out once it is known hidden, keeping the top slide in place', async () => {
    const engine = new FakePptxEngine(6, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.setLayoutProgress(2, false);
    const { v, topOf, scrollTo, scrollHost } = setup(6, [3], 'skip', { engine });
    // Slide 3's metadata is unknown, so it is laid out for now.
    scrollTo(LEADING + 4 * STRIDE + 20); // 20 px into slide 4
    expect(v.topVisibleSlide).toBe(4);
    engine.setLayoutProgress(6, true);
    await Promise.resolve();
    expect(v.topVisibleSlide).toBe(4);
    expect(topOf(4)).toBe(`${LEADING + 3 * STRIDE}px`);
    expect(scrollHost.scrollTop).toBe(LEADING + 3 * STRIDE + 20);
    v.scrollToSlide(3);
    expect(v.topVisibleSlide).toBe(4);
    v.destroy();
  });
});

describe("PptxScrollViewer hiddenSlideMode: 'dim' and 'show'", () => {
  it("'dim' draws only hidden slides under the overlay; 'show' never dims", () => {
    const dimmed = setup(3, [1], 'dim', { viewport: 400, opts: { hiddenSlideDim: { opacity: 0.5 } } });
    expect(dimmed.renderSlide.mock.calls.map((call) => [call[1], call[2]?.dim])).toEqual([
      [0, undefined],
      [1, { color: '#ffffff', opacity: 0.5 }],
      [2, undefined],
    ]);
    dimmed.v.destroy();

    const shown = setup(3, [1], undefined, { viewport: 400 });
    expect(shown.v.hiddenSlideMode).toBe('show');
    expect(shown.drawn()).toEqual([0, 1, 2]);
    expect(shown.renderSlide.mock.calls.every((call) => call[2]?.dim === undefined)).toBe(true);
    shown.v.destroy();
  });
});

describe('PptxScrollViewer.setHiddenSlideMode', () => {
  it('re-lays out while scrolled, keeps the top slide, and clears find', async () => {
    const { v, topOf, scrollTo, scrollHost } = setup(6, [1, 2], 'show');
    scrollTo(LEADING + 3 * STRIDE + 30); // 30 px into slide 3
    await v.findText('needle');
    const find = (v as unknown as { _find: { slideHighlights(slide: number): unknown[] } })._find;
    expect(find.slideHighlights(3).length).toBeGreaterThan(0);

    await v.setHiddenSlideMode('skip');
    expect(v.hiddenSlideMode).toBe('skip');
    expect(v.topVisibleSlide).toBe(3);
    expect(topOf(3)).toBe(`${LEADING + STRIDE}px`);
    expect(scrollHost.scrollTop).toBe(LEADING + STRIDE + 30);
    expect(find.slideHighlights(3)).toEqual([]);
    expect(await v.findNext()).toBeNull();

    await v.setHiddenSlideMode('show');
    expect(v.topVisibleSlide).toBe(3);
    expect(topOf(3)).toBe(`${LEADING + 3 * STRIDE}px`);
    expect(scrollHost.scrollTop).toBe(LEADING + 3 * STRIDE + 30);
    v.destroy();
  });

  it('a hidden top slide gives way to the next shown slide; dim toggles redraw only hidden slides', async () => {
    const { v, scrollTo, renderSlide } = setup(8, [1], 'show', { viewport: 400 });
    scrollTo(LEADING + STRIDE + 40); // inside hidden slide 1
    expect(v.topVisibleSlide).toBe(1);
    await v.setHiddenSlideMode('skip');
    expect(v.topVisibleSlide).toBe(2);
    expect(viewerState(v).mounted()).not.toContain(1);

    await v.setHiddenSlideMode('show');
    scrollTo(0);
    renderSlide.mockClear();
    await v.setHiddenSlideMode('dim');
    expect(renderSlide.mock.calls.map((call) => [call[1], call[2]?.dim?.opacity])).toEqual([[1, 0.6]]);
    v.destroy();
  });
});

describe('PptxScrollViewer hidden-slide review fixes', () => {
  it('a dim switch during a zoom burst repaints at the current scale and the settle still covers the rest', async () => {
    vi.useFakeTimers();
    try {
      const { v, renderSlide } = setup(3, [1], 'show', { viewport: 400, opts: { enableTextSelection: true } });
      await Promise.resolve();
      v.setScale(v.getScale() * 2); // preview stretch; settle debounced
      renderSlide.mockClear();
      await v.setHiddenSlideMode('dim');
      const slots = (v as unknown as {
        _slots: Map<number, { renderedScale: number; textLayer: FakeEl | null }>;
      })._slots;
      // The hidden slide repainted at the new scale, without a preview stretch.
      expect(renderSlide.mock.calls.map((call) => call[1])).toEqual([1]);
      expect(renderSlide.mock.calls[0]?.[2]?.width).toBe(400);
      expect(slots.get(1)?.renderedScale).toBe(v.getScale());
      expect(slots.get(1)?.textLayer?.style.transform ?? '').toBe('');
      // Shown slides still preview at the old scale until the settle repaints them.
      expect(slots.get(0)?.renderedScale).not.toBe(v.getScale());
      vi.advanceTimersByTime(200);
      for (let i = 0; i < 5; i++) await Promise.resolve();
      expect([...slots.values()].every((slot) => slot.renderedScale === v.getScale())).toBe(true);
      v.destroy();
    } finally {
      vi.useRealTimers();
    }
  });

  it('a switch to skip during goToComment bounds resolution does not select the now-hidden comment', async () => {
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.commentsBySlide = [[], [{
      x: 0, y: 0, text: 'c',
      anchors: [{ type: 'slide' }, { type: 'drawingElement', elementId: 'e1' }],
    } as never], []];
    const { v } = setup(3, [1], 'show', { engine, opts: { comments: true } });
    const navigation = v.goToComment(1, 0);
    void v.setHiddenSlideMode('skip');
    expect(await navigation).toBe(false);
    expect(v.getSelectionContext()).toBeNull();
    v.destroy();
  });

  it('setHiddenSlideMode resolves only after every affected mounted slide has painted', async () => {
    const engine = new FakePptxEngine(6, SLIDE_W_EMU, SLIDE_H_EMU, 'main', true);
    const { v } = setup(6, [1], 'show', { engine, viewport: 300 });
    for (const call of engine.renderCalls) call.resolve();
    const track = (promise: Promise<void>) => {
      let done = false;
      void promise.then(() => { done = true; });
      return async () => {
        for (let i = 0; i < 20; i++) await Promise.resolve();
        return done;
      };
    };
    // 'dim' repaints the mounted hidden slide 1.
    let before = engine.renderCalls.length;
    let isDone = track(v.setHiddenSlideMode('dim'));
    expect(await isDone()).toBe(false);
    expect(engine.renderCalls.slice(before).map((call) => call.slide)).toEqual([1]);
    for (const call of engine.renderCalls.slice(before)) call.resolve();
    expect(await isDone()).toBe(true);
    // 'skip' leaves slide 1 out, so slide 3 moves into the window.
    before = engine.renderCalls.length;
    isDone = track(v.setHiddenSlideMode('skip'));
    expect(await isDone()).toBe(false);
    const fresh = engine.renderCalls.slice(before);
    expect(fresh.map((call) => call.slide)).toContain(3);
    for (const call of fresh) call.resolve();
    expect(await isDone()).toBe(true);
    v.destroy();
  });

  it('worker mode: setHiddenSlideMode waits for the paint that replaces an in-flight bitmap', async () => {
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const bitmap = vi.spyOn(engine, 'renderSlideToBitmap');
    const { v } = setup(3, [1], 'show', { engine, viewport: 400 });
    const settle = async () => { for (let i = 0; i < 20; i++) await Promise.resolve(); };
    // Slide 1's first bitmap is still in flight when the mode changes.
    const first = engine.bitmapCalls.find((call) => call.slide === 1);
    expect(first).toBeDefined();
    let done = false;
    void v.setHiddenSlideMode('dim').then(() => { done = true; });
    await settle();
    expect(done).toBe(false);
    // The superseded bitmap lands; the dimmed replacement is dispatched.
    first?.resolve();
    await settle();
    expect(done).toBe(false);
    const replacement = engine.bitmapCalls.filter((call) => call.slide === 1).at(-1);
    expect(replacement).not.toBe(first);
    expect(bitmap.mock.calls.filter((call) => call[0] === 1).at(-1)?.[1]?.dim?.opacity).toBe(0.6);
    replacement?.resolve();
    await settle();
    expect(done).toBe(true);
    for (const call of engine.bitmapCalls) call.resolve();
    v.destroy();
  });

  it('progressive skip inspects each slide once over the whole load, and show mode not at all', () => {
    for (const mode of ['skip', 'show'] as const) {
      const engine = new FakePptxEngine(40, SLIDE_W_EMU, SLIDE_H_EMU);
      engine.setLayoutProgress(1, false);
      const { v } = setup(40, [5, 20], mode, { engine });
      const isHidden = vi.spyOn(engine, 'isHidden');
      for (let available = 2; available <= 40; available++) {
        engine.setLayoutProgress(available, available === 40);
      }
      expect(isHidden).toHaveBeenCalledTimes(mode === 'skip' ? 39 : 0);
      v.destroy();
    }
  });

  it('keeps the signed offset of the top slide: leading padding and the trailing gap survive a relayout', async () => {
    // scrollTop 0 sits in the leading padding above hidden slide 0.
    const first = setup(6, [0], 'show');
    await first.v.setHiddenSlideMode('skip');
    expect(first.scrollHost.scrollTop).toBe(0);
    first.v.destroy();

    // 5 px into the gap below slide 2 stays 5 px into the gap below it.
    const gap = setup(6, [1], 'show');
    gap.scrollTo(LEADING + 3 * STRIDE - 5);
    expect(gap.v.topVisibleSlide).toBe(2);
    await gap.v.setHiddenSlideMode('skip');
    expect(gap.scrollHost.scrollTop).toBe(LEADING + 2 * STRIDE - 5);
    gap.v.destroy();
  });
});
