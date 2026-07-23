import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeContainer, makeEl, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';
import * as pptxIndex from './index.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('PptxScrollViewer — skeleton (T1)', () => {
  it('builds the wrapper → scrollHost → spacer DOM inside the container', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    // container → wrapper
    const wrapper = container.children[0];
    expect(wrapper.tag).toBe('div');
    expect(wrapper.style.position).toBe('relative');
    // wrapper → scrollHost
    const scrollHost = wrapper.children[0];
    expect(scrollHost.style.overflow).toBe('auto');
    // scrollHost → spacer
    const spacer = scrollHost.children[0];
    expect(spacer.style.position).toBe('absolute');
    v.destroy();
  });

  it('exposes slideCount from the injected engine', () => {
    installDom();
    const engine = new FakePptxEngine(5, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(v.slideCount).toBe(5);
    v.destroy();
  });

  it('load() is unsupported when an engine is injected', async () => {
    installDom();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
  });

  it('destroy() removes the DOM and does NOT destroy an injected engine', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(container.children.length).toBe(1); // wrapper mounted
    v.destroy();
    expect(container.children.length).toBe(0); // wrapper removed
    expect(engine.destroyed).toBe(false); // injected engine preserved (caller owns it)
  });

  it('slideCount is 0 before load resolves (no injected engine)', () => {
    installDom();
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    expect(v.slideCount).toBe(0);
    v.destroy();
  });

  // A <canvas> type-checks as HTMLElement (the pager API takes one), but canvas
  // children never render — the viewer would come up silently blank. Rejected
  // loudly at construction instead.
  it('throws when the container is a <canvas> (pager-API confusion)', () => {
    installDom();
    expect(
      () => new PptxScrollViewer(makeEl('canvas') as unknown as HTMLElement, {}),
    ).toThrow(/container element .* not a <canvas>/i);
    // A plain div (the documented contract) constructs fine.
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    v.destroy();
  });

  // O1 (design §11): an injected engine's own `mode` is authoritative. An
  // EXPLICITLY conflicting opts.mode is a mis-configuration rejected at
  // construction; a matching or absent opts.mode constructs fine.
  it('throws when opts.mode conflicts with an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU, 'worker');
    expect(
      () =>
        new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
          presentation: engine.asPres(),
          mode: 'main',
        }),
    ).toThrow(/mode/i);
  });

  it('does NOT throw when opts.mode matches an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU, 'worker');
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
      mode: 'worker',
    });
    expect(v.slideCount).toBe(1);
    v.destroy();
    // Injected engine is caller-owned even in the worker case: destroy() leaves it intact.
    expect(engine.destroyed).toBe(false);
  });

  it('constructs a default-main injected engine with absent opts.mode (load still rejects; destroy preserves engine)', async () => {
    installDom();
    // Default mode is 'main'; opts.mode is absent ⇒ no conflict, resolved path is main.
    const engine = new FakePptxEngine(2, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    expect(v.slideCount).toBe(2);
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
    // Injected engine is caller-owned: destroy() must not tear it down.
    expect(engine.destroyed).toBe(false);
  });

  // `background` paints the scroll surface (the "desk") visible behind/between
  // slides. It applies to the viewer-owned scrollHost; slides keep their own white.
  it('applies opts.background to the scrollHost element', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      background: '#525659',
    });
    const scrollHost = container.children[0].children[0]; // wrapper → scrollHost
    expect(scrollHost.style.background).toBe('#525659');
    v.destroy();
  });

  it('sets no background on the scrollHost by default (transparent — container shows through)', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    const scrollHost = container.children[0].children[0];
    expect(scrollHost.style.background).toBe('');
    v.destroy();
  });
});

// Fake slide geometry, chosen so the DIMENSIONLESS base scale is a clean 1.0 and
// arithmetic stays tidy. `_scale` is now a multiplier over the 96-dpi natural
// size (natural px = EMU / EMU_PER_PX, EMU_PER_PX = 9525), mirroring docx.
//   slide width  = 9525 × 200 = 1,905,000 EMU ⇒ natural 200px
//   slide height = 9525 × 120 = 1,143,000 EMU ⇒ natural 120px
// Container width 200 ⇒ base = 200 / 200 = 1.0 ⇒ slide 200×120px at base.
// (Old geometry was 1000×600 EMU with a px-per-EMU scale of 0.2; that yielded the
// same 200×120px slide, so the stride math below — 120px tall, gap 10, stride 130
// — is unchanged. Only the SCALE VALUE changes: base 0.2 → 1.0, zoom ×2 → 2.0.)
const SLIDE_W_EMU = 9525 * 200; // 1,905,000 ⇒ natural 200px
const SLIDE_H_EMU = 9525 * 120; // 1,143,000 ⇒ natural 120px
// Tall variant used by the worker recycle/re-mount test: natural 400px tall ⇒
// 200×400px at base 1.0, so the visible window is exactly one slide + overscan.
// (Was 1000×2000 EMU with the old 0.2 px/EMU scale — the same 200×400px slide.)
const SLIDE_H_TALL_EMU = 9525 * 400; // 3,810,000 ⇒ natural 400px

describe('PptxScrollViewer — layout + virtualization (T2)', () => {
  // See the SLIDE_W_EMU/SLIDE_H_EMU note above: container 200 ⇒ dimensionless base
  // scale 1.0 ⇒ each slide is 200px wide, 120px tall.
  function setup(slideCount: number, opts = {}) {
    const dom = installDom();
    const container = makeContainer(200, 400); // clientWidth 200, clientHeight 400
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      width: undefined, // fit container width (200) → base scale below
      // Flush horizontal gutters: paddingLeft/Right default to `gap`, which would
      // shrink the fit width to 180 and break the base-scale-1.0 geometry the T2
      // assertions assume (slide width px = 200). Pin them to 0 so the fit is the
      // full 200 container width (same pattern the vertical commit used).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const wrapper = container.children[0] as FakeEl;
    const scrollHost = wrapper.children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    // Drive layout: container width 200, natural slide width 200px → dimensionless
    // base scale 1.0. Provide the viewport height.
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { dom, container, engine, v, wrapper, scrollHost, spacer };
  }

  it('sizes the spacer to computeVisibleRange.totalHeight and mounts only the visible window', () => {
    const { v, scrollHost, spacer } = setup(10);
    v.relayout(); // T2 exposes an explicit relayout() the viewer calls after load/resize
    // Spacer height > 0 and equals n * slideHeightPx + (n-1)*gap in px.
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // At scrollTop 0 with a 400px viewport and 120px-tall slides, several slides
    // fit; overscan 1 adds one more. Assert the mounted slot count is bounded.
    const mounted = scrollHost.children.filter((c) => c !== spacer);
    expect(mounted.length).toBeGreaterThanOrEqual(1);
    // Viewport 400 / stride 130 ⇒ ~4 slides visible + overscan; comfortably bounded.
    expect(mounted.length).toBeLessThanOrEqual(6);
    v.destroy();
  });

  it('spacer height is exact for uniform slide heights + gap', () => {
    const dom = installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      paddingLeft: 0, // full-width fit (see T2 setup note) → base scale 1.0
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Dimensionless base scale = 200 / 200 (natural width) = 1.0 ⇒ slide height px =
    // (SLIDE_H_EMU / 9525) * 1.0 = 120.
    const slideH = SLIDE_H_EMU / 9525; // 120
    const n = 3;
    // New default: paddingTop/paddingBottom each default to `gap` (10), so the
    // spacer is padTop + Σheights + (n-1)*gap + padBottom.
    const expected = 10 + n * slideH + (n - 1) * 10 + 10;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
    void dom;
  });

  it('recycles slots on scroll without unbounded canvas growth (pool reuse)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    const initialMount = scrollHost.children.length;
    // Scroll far down and fire the scroll listener repeatedly.
    for (let top = 0; top <= 4000; top += 400) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
    }
    // The DOM child count (spacer + mounted slots) must stay bounded — the pool
    // reuses slots rather than appending a new canvas per slide.
    expect(scrollHost.children.length).toBeLessThanOrEqual(initialMount + 2);
  });

  it('scrolling far then back reuses pooled slot wrappers (bounded distinct allocations)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    // Track EVERY distinct wrapper element the viewer ever appends. If slots are
    // pooled (recycled), scrolling across the whole deck then back reuses a small
    // fixed set of wrappers rather than allocating one per visited slide.
    const seen = new Set<FakeEl>();
    const collect = () => {
      for (const c of scrollHost.children) if (c.tag === 'div') seen.add(c);
    };
    collect();
    // Sweep deep into the deck and back to the top.
    for (const top of [3000, 6000, 3000, 0]) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
      collect();
    }
    // 50 slides visited across the sweep; a per-slide allocator would have created
    // dozens of wrappers. The pool must keep the distinct-wrapper count tiny:
    // spacer(1) + at most a couple of window-sized generations of slots.
    expect(seen.size).toBeLessThanOrEqual(12);
    // Coming back to the top mounts slide 0 again, drawn from the pool.
    expect(v.mountedSlideIndicesForTest()).toContain(0);
  });

  it('mounts the correct window for a mid-deck scrollTop', () => {
    const { v, scrollHost } = setup(20);
    v.relayout();
    // Jump to a scrollTop deep in the deck; the mounted slots must include the
    // slide under the viewport top (topVisibleSlide) and its overscan neighbours.
    scrollHost.scrollTop = 1300;
    scrollHost.dispatch('scroll');
    const top = v.topVisibleSlide;
    const mountedSlides = v.mountedSlideIndicesForTest(); // T2 test hook
    expect(mountedSlides).toContain(top);
    expect(Math.max(...mountedSlides) - Math.min(...mountedSlides)).toBeLessThanOrEqual(6);
  });
});

describe('PptxScrollViewer — rendering (T3)', () => {
  it("main mode: calls renderSlide once per mounted slot with that slide's px width", () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingLeft: 0, // full-width fit → slide width px = 200 (asserted below)
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Every mounted slide got exactly one renderSlide call, no more no less.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    const rendered = engine.renderCalls.map((c) => c.slide).sort((a, b) => a - b);
    expect(rendered).toEqual(mounted);
    // Each renderSlide call carried the per-call px width. pptx slides are uniform,
    // so every width equals the same fit-width (dimensionless base scale 1.0 over
    // the natural 200px width ⇒ slide width px = 200), but the per-call width
    // contract still passes a width per slide (mirrors docx's per-page width, §7).
    const widths = engine.renderSlideWidths();
    expect(widths.length).toBe(mounted.length);
    for (const w of widths) expect(w).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('main mode: passes the uniform px width per slide (per-call width contract, §7)', () => {
    installDom();
    const container = makeContainer(200, 400);
    // pptx slide size is UNIFORM deck-wide, so unlike docx (mixed page sizes) every
    // slide renders at the SAME px width. The per-call width contract still holds:
    // each mounted slide receives its own width argument, equal to the uniform px.
    const engine = new FakePptxEngine(2, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingLeft: 0, // full-width fit → base scale 1.0 (slide width px = 200)
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Both slides mount (slide 0 height px = 120 at base scale 1.0; slide 1 top 130
    // within the 400px viewport), so both get exactly one render.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    expect(mounted).toEqual([0, 1]);
    // Per-slide px widths in call order: uniform 200 each.
    const widths = engine.renderCalls
      .slice()
      .sort((a, b) => a.slide - b.slide)
      .map((c) => c.width ?? NaN);
    expect(widths[0]).toBeCloseTo(200, 3);
    expect(widths[1]).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('does not re-render a mounted slot for the same slide on a no-op scroll', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const before = engine.renderCalls.length;
    scrollHost.scrollTop = 0; // unchanged window
    scrollHost.dispatch('scroll');
    expect(engine.renderCalls.length).toBe(before); // no duplicate renders
    v.destroy();
  });

  it('worker mode: never calls renderSlide — routes every slot through renderSlideToBitmap', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // mode:'worker' — the real PptxPresentation.renderSlide THROWS synchronously in
    // worker mode; a viewer that mis-routed would blow up (and renderCalls would
    // record the attempt). The direct _mode routing must never touch renderSlide.
    const engine = new FakePptxEngine(10, SLIDE_W_EMU, SLIDE_H_EMU, 'worker');
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(0); // renderSlide never touched
    // Every mounted slide dispatched exactly one bitmap render.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    const bitmapSlides = [...new Set(engine.bitmapCalls.map((c) => c.slide))].sort((a, b) => a - b);
    expect(bitmapSlides).toEqual(mounted);
    v.destroy();
  });

  it('worker mode: paints a resolved bitmap into the slot canvas (transfer)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, SLIDE_W_EMU, SLIDE_H_EMU, 'worker');
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // A bitmap was produced and transferred into the mounted slot's canvas.
    expect(engine.createdBitmaps.length).toBeGreaterThan(0);
    const slotWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const canvas = slotWrapper.children.find((k) => k.tag === 'canvas') as FakeEl;
    // The bitmaprenderer ctx received the bitmap (fake records lastBitmap).
    expect(canvas._bitmapCtx?.lastBitmap).toBe(engine.createdBitmaps[0]);
    v.destroy();
  });

  it('worker mode: closes the ImageBitmap on recycle (deferred — render in flight when the slot recycles)', async () => {
    // Bitmap-close observability path: the primary path (deterministic slot.bitmap
    // hold under synchronous resolve) does NOT hold — in non-deferred mode
    // transferFromImageBitmap consumes the bitmap and nulls slot.bitmap
    // synchronously BEFORE any scroll-away, so a later recycle has nothing to
    // close. We use the DOCUMENTED FALLBACK: a deferred fake so the render is
    // genuinely in flight when the slot recycles, and the on-resolution
    // stale-check closes the orphan (design §11).
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const firstBatch = engine.bitmapCalls.length;
    expect(firstBatch).toBeGreaterThan(0);
    // Scroll far away so the first slides' slots recycle while their renders are
    // still in flight (deferred → not yet resolved).
    scrollHost.scrollTop = 6000;
    scrollHost.dispatch('scroll');
    // Now resolve the early (now-stale) renders — the viewer must close them.
    for (const call of engine.bitmapCalls.slice(0, firstBatch)) call.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.createdBitmaps.some((b) => b.close.mock.calls.length > 0)).toBe(true);
    v.destroy();
  });

  it('worker mode: drops a stale in-flight render — no paint + bitmap closed when the slot moved on', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred: renderSlideToBitmap resolves only when the test calls resolve(),
    // so we can scroll a slot's slide out of the window BEFORE the bitmap arrives.
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Slide 0's bitmap is dispatched but NOT yet resolved.
    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    // Scroll far away so slide 0's slot recycles while its render is in flight.
    scrollHost.scrollTop = 6000;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).not.toContain(0);
    // Now the stale render for slide 0 resolves — the viewer must NOT paint it and
    // must close the orphaned bitmap.
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0!)];
    slide0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed
    v.destroy();
  });

  it('worker mode: a slide that recycles then re-mounts while in flight still gets a fresh render', async () => {
    // Coalescing keys on slide index; a naive Set<number> would swallow the
    // re-mounted slot's render (the stale resolve clears in-flight AFTER the
    // remount coalesced away → the new slot never paints). The viewer must
    // re-dispatch the live slot's render so the slide never stays blank.
    installDom();
    const container = makeContainer(200, 400);
    // Tall slides (natural 200×400px ⇒ 200×400px at base scale 1.0, stride 410) so the
    // visible window is exactly one slide + overscan = [0,1], mirroring the docx
    // recycle/re-mount pool dynamics: the re-mounted slide-0 slot is a DIFFERENT
    // pooled object than the in-flight one, so `live !== slot` triggers the
    // re-dispatch. (A short-slide window packs several slots, whose LIFO reuse can
    // hand slide 0 back its own in-flight slot — an epoch-only case not under test.)
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_TALL_EMU, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0, // flush top so slide 0's slot sits at top:0px (asserted below)
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const slide0Initial = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Initial).toBeDefined();
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);
    // Scroll away (slide 0 recycles, render still in flight) then back to the top
    // (slide 0 re-mounts on a fresh slot) — all while the initial render is deferred.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).not.toContain(0);
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    // Coalescing pin: slide 0's render is STILL in flight (initial deferred call not
    // yet resolved), so the re-mount must NOT dispatch a second render for slide 0
    // — the in-flight guard swallows it. Exactly one dispatch so far.
    expect(dispatchesForSlide0()).toBe(1);
    // The initial (now stale) render resolves — the orphan is dropped and the
    // re-mounted slot's render is (re-)dispatched.
    slide0Initial!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(dispatchesForSlide0()).toBeGreaterThanOrEqual(2); // fresh render issued
    // Resolve the fresh render and confirm it paints into the live slot.
    for (const c of engine.bitmapCalls.filter((c) => c.slide === 0)) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted, not blank
    v.destroy();
  });

  it('worker mode: a PLAIN render rejection (slot still live, epoch unchanged) does NOT re-dispatch — no retry storm', async () => {
    // B1 regression: the finally re-dispatch must gate on STALENESS, not merely
    // `!painted`. When `renderSlideToBitmap` rejects while the slot is still live
    // and the epoch is unchanged, `painted` is false and `live === slot`, but
    // NEITHER staleness test fires, so we must NOT re-dispatch. A `!painted`-only
    // gate would loop reject → re-dispatch → reject … unbounded (empirically 1→2
    // →3→4 with onError every round). The onError contract leaves the slide blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    expect(v.mountedSlideIndicesForTest()).toContain(0); // slot still live
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);

    // Reject the render with the slot STILL LIVE and no scale change (epoch fixed).
    slide0!.reject(new Error('worker render failed'));
    // Flush microtasks generously — a retry storm would keep queuing dispatches.
    for (let k = 0; k < 8; k++) await Promise.resolve();

    // Dispatch count for slide 0 stayed at 1 — the storm is gone.
    expect(dispatchesForSlide0()).toBe(1);
    // onError fired exactly once (the single failure), never again.
    expect(onError).toHaveBeenCalledTimes(1);
    // Still no further dispatches after more microtask flushing.
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForSlide0()).toBe(1);
    expect(onError).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('worker mode: destroy() mid-flight closes the resolving bitmap and does not fire onError post-destroy', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred so a render is genuinely in flight when we destroy the viewer.
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    // Tear down while slide 0's render is still in flight. destroy() recycles the
    // slot (no bitmap held yet — none received), so the on-resolution stale-check
    // must close the orphan; the identity guard fails (slot no longer live).
    v.destroy();
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0!)];
    slide0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed, no leak
    expect(onError).not.toHaveBeenCalled(); // no error surfaced after teardown
  });

  // ⚠ MANDATORY render-epoch test (T3 review finding I-3). Worker path, deferred:
  // dispatch at scale A, setScale(B) mid-flight, resolve — the old-scale bitmap is
  // closed, NOT painted, and a fresh dispatch at scale B repaints the slot.
  it('render epoch: a bitmap dispatched at the old scale is dropped (closed, not painted) after setScale, and re-dispatched at the new scale', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0, // flush top so slide 0's slot sits at top:0px (asserted below)
      overscan: 1,
      // REAL defaults [0.1, 4]: base fit is 1.0, and setScale(×2)=2.0 is inside them
      // (the old `zoomMin: 0.05, zoomMax: 3` dodged the px-per-EMU unit bug — WS4b-2).
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    // Slide 0's bitmap is dispatched at the base (old) scale but NOT yet resolved.
    const slide0Old = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Old).toBeDefined();
    const oldDispatchCount = engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(oldDispatchCount).toBe(1);
    const oldWidth = slide0Old!.width;

    // Zoom mid-flight: bumps the render epoch and re-mounts. Slide 0's re-mount is
    // coalesced away by the in-flight guard (same index still in flight).
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedSlideIndicesForTest()).toContain(0); // slide 0 still visible at top
    expect(engine.bitmapCalls.filter((c) => c.slide === 0).length).toBe(1); // no double-dispatch yet

    // The OLD-scale render resolves. Epoch moved ⇒ STALE: close, do not paint,
    // then re-dispatch slide 0 at the NEW scale.
    const bmpOld = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0Old!)];
    slide0Old!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmpOld.close.mock.calls.length).toBeGreaterThan(0); // old bitmap closed
    // A fresh dispatch for slide 0 was issued at the new scale.
    const slide0Calls = engine.bitmapCalls.filter((c) => c.slide === 0);
    expect(slide0Calls.length).toBeGreaterThanOrEqual(2);
    const freshCall = slide0Calls[slide0Calls.length - 1];
    // The fresh dispatch's width reflects the NEW scale (double the old width).
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0);

    // The stale old bitmap was NOT transferred; only the fresh one paints.
    freshCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).not.toBe(bmpOld); // never painted the stale bitmap
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh one
    v.destroy();
  });

  // B1: epoch-then-reject must stay BOUNDED. setScale bumps the epoch mid-flight,
  // so the OLD dispatch is stale on resolution; whether it resolves or REJECTS,
  // the finally must issue exactly ONE fresh dispatch at the new epoch. The fresh
  // dispatch captures the new epoch, so if IT later rejects (same epoch, still
  // live) nothing re-dispatches — the retry is bounded to a single fresh attempt.
  it('render epoch: rejecting the OLD-scale dispatch after setScale still yields exactly ONE fresh dispatch at the new scale (bounded)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0, // flush top so slide 0's slot sits at top:0px (asserted below)
      overscan: 1,
      // REAL defaults [0.1, 4]: base fit is 1.0, and setScale(×2)=2.0 is inside them
      // (the old `zoomMin: 0.05, zoomMax: 3` dodged the px-per-EMU unit bug — WS4b-2).
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const slide0Old = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Old).toBeDefined();
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);
    const oldWidth = slide0Old!.width;

    // Zoom mid-flight: bumps the epoch. Slide 0's re-mount is coalesced away.
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    expect(dispatchesForSlide0()).toBe(1);

    // REJECT the old-scale dispatch. Epoch moved ⇒ stale ⇒ re-dispatch (not a plain
    // failure retry). Exactly ONE fresh dispatch at the new epoch.
    slide0Old!.reject(new Error('old-scale render failed'));
    for (let k = 0; k < 8; k++) await Promise.resolve();
    const afterReject = engine.bitmapCalls.filter((c) => c.slide === 0);
    expect(afterReject.length).toBe(2); // one fresh dispatch, no storm
    const freshCall = afterReject[afterReject.length - 1];
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0); // at the NEW (larger) scale

    // The fresh dispatch then SUCCEEDS and paints — the slide is not left blank.
    freshCall.resolve();
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForSlide0()).toBe(2); // still exactly two; success does not re-dispatch
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh bitmap
    v.destroy();
  });
});

describe('PptxScrollViewer — zoom (T4)', () => {
  // `_scale` is a DIMENSIONLESS multiplier over the 96-dpi natural slide size
  // (mirrors docx). Container 200×400, width:undefined ⇒ base fit maps the natural
  // 200px slide width (SLIDE_W_EMU / 9525) to the 200px container:
  //   base = 200 / 200 = 1.0
  //   slide height px = (SLIDE_H_EMU / 9525) * 1.0 = 120
  //   offset[i] = i * (120 + gap) = i * 130
  // Zoom ×2 ⇒ scale 2.0, height 240, stride 250. Because base is now 1.0, base×2 =
  // 2.0 sits inside the DEFAULT [0.1, 4] bounds, so setup uses REAL defaults — the
  // old `zoomMin: 0.05, zoomMax: 3` workarounds only existed to dodge the px-per-EMU
  // base of 0.2 (WS4b-2 review finding) and are removed. The two MANDATORY render-
  // epoch tests (resolve + reject) live in the T3 rendering block above.
  function setup(slideCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      // Flush top/bottom (paddingTop/Bottom default to `gap`; the T4 offset
      // arithmetic below is written for offset[0]===0). This exercises the
      // "explicit 0 ⇒ old flush behavior reachable" contract.
      paddingTop: 0,
      paddingBottom: 0,
      // Flush left/right so the fit width is the full 200 container ⇒ base 1.0.
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  it('base fit scale maps the slide width to the container width', () => {
    const { v } = setup();
    // base = 200 / 200 (natural width) = 1.0
    expect(v.scaleForTest()).toBeCloseTo(1.0, 6);
    expect(v.baseScaleForTest()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('re-anchors so the slide under the viewport top stays fixed across a zoom (intraFrac 0)', () => {
    const { v, scrollHost } = setup();
    // slide 3 top offset at base (1.0): 3 * (120 + 10) = 390
    scrollHost.scrollTop = 390;
    scrollHost.dispatch('scroll');
    expect(v.topVisibleSlide).toBe(3);
    const cur = v.scaleForTest(); // 1.0
    v.setScale(cur * 2); // 2.0 (within default [0.1, 4], no clamp)
    expect(v.scaleForTest()).toBeCloseTo(2.0, 6);
    // slide 3 stays the top slide; new offset = 3 * (240 + 10) = 750
    expect(v.topVisibleSlide).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - 750)).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchors preserving the intra-slide fraction (intraFrac ≠ 0)', () => {
    const { v, scrollHost } = setup();
    // Scroll so HALF of slide 3 (60 of its 120px) has passed above the viewport
    // top: scrollTop = offset[3] + 0.5*120 = 390 + 60 = 450 → intraFrac 0.5.
    scrollHost.scrollTop = 450;
    scrollHost.dispatch('scroll');
    expect(v.topVisibleSlide).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 1.0 → 2.0; heights' = 240
    // newScrollTop = offset'[3] + 0.5 * 240 = 750 + 120 = 870
    expect(Math.abs(scrollHost.scrollTop - 870)).toBeLessThan(2);
    // Slide 3 is still under the viewport top: offset'[3]=750 <= 870 < 1000.
    expect(v.topVisibleSlide).toBe(3);
    v.destroy();
  });

  it('clamps setScale to the absolute [zoomMin, zoomMax] dimensionless bounds', () => {
    // Explicit bounds so the clamp is unambiguous (base fit is 1.0, inside them).
    const { v } = setup(20, { zoomMin: 0.5, zoomMax: 3 });
    v.setScale(100); // above zoomMax 3 (absolute, NOT a multiple of base)
    expect(v.scaleForTest()).toBeCloseTo(3, 6);
    v.setScale(0.001); // below zoomMin 0.5
    expect(v.scaleForTest()).toBeCloseTo(0.5, 6);
    v.destroy();
  });

  it('setScale defaults clamp to [0.1, 4] when zoomMin/zoomMax are unset', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(5, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    v.setScale(999);
    expect(v.scaleForTest()).toBeCloseTo(4, 5); // default zoomMax
    v.setScale(0.0001);
    expect(v.scaleForTest()).toBeCloseTo(0.1, 5); // default zoomMin
    v.destroy();
  });

  // ⚠ MANDATORY real-magnitude regression (WS4b review finding A). Before the
  // dimensionless-scale fix, `_scale` was px-per-EMU (~6.6e-5 for a real 12.192M-EMU
  // deck), which the DEFAULT [0.1, 4] clamp then inflated by ~1000× on the first
  // Ctrl+wheel or resize (spacer 6,274px → 6,172,330px). With the dimensionless
  // scale, a real 16:9 deck's base fit lands INSIDE the default bounds and a wheel
  // tick multiplies it by ~e, never explodes.
  it('real 16:9 deck: base fit is inside [0.1, 4] (NOT clamped) and one wheel tick multiplies the spacer by ~e, not ~1000×', () => {
    installDom();
    // A real PowerPoint 16:9 deck: 13.333in × 7.5in = 12,192,000 × 6,858,000 EMU.
    // Natural width px = 12,192,000 / 9525 = 1280. Container ~1200px.
    const container = makeContainer(1200, 800);
    const engine = new FakePptxEngine(10, 12192000, 6858000);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 16,
      // No zoomMin/zoomMax ⇒ REAL defaults [0.1, 4].
      // Flush horizontal gutters so the fit is the full 1200 container ⇒ base
      // 1200/1280 = 0.9375 (the default `gap` gutters would shrink it to 1168/1280).
      paddingLeft: 0,
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    scrollHost.clientHeight = 800;
    scrollHost.clientWidth = 1200;
    v.relayout();

    // base = 1200 / 1280 = 0.9375 — INSIDE [0.1, 4], so relayout did NOT clamp.
    const base = v.scaleForTest();
    expect(base).toBeCloseTo(0.9375, 6);
    expect(base).toBeGreaterThan(0.1);
    expect(base).toBeLessThan(4);
    const spacerBase = parseFloat(spacer.style.height);
    expect(spacerBase).toBeGreaterThan(0);

    // One Ctrl+wheel tick up (deltaY −100): zoomStepScale multiplies by exp(1) ≈
    // 2.71828, so 0.9375 × e ≈ 2.548 — still inside [0.1, 4], NOT clamped to 4.
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, metaKey: false, preventDefault() {} });
    const zoomed = v.scaleForTest();
    expect(zoomed).toBeCloseTo(0.9375 * Math.E, 4); // ≈ 2.548, not the clamped 4
    expect(zoomed).toBeLessThan(4);

    // The spacer grew by the scale ratio (≈ e), NOT by ~1000× (the old px-per-EMU
    // clamp explosion). Slide height px ∝ _scale, so the spacer's slide-height
    // portion scales by (zoomed / base); with the fixed gap term the total ratio is
    // very close to e. Bound it well below the old 1000× blowup.
    const spacerZoomed = parseFloat(spacer.style.height);
    const ratio = spacerZoomed / spacerBase;
    expect(ratio).toBeGreaterThan(2.5); // grew ~e×
    expect(ratio).toBeLessThan(3); // and NOT ~1000× (the pre-fix explosion)

    // setScale(999) still clamps hard to the default zoomMax 4.
    v.setScale(999);
    expect(v.scaleForTest()).toBeCloseTo(4, 5);
    v.destroy();
  });

  it('Ctrl+wheel zooms in and calls preventDefault; bare wheel does not zoom or preventDefault', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();

    // Bare wheel: no zoom, native scroll (preventDefault NOT called).
    const barePrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: false, preventDefault: barePrevent });
    expect(v.scaleForTest()).toBe(before);
    expect(barePrevent).not.toHaveBeenCalled();

    // Ctrl+wheel (deltaY<0 = zoom in): scale increases, preventDefault called.
    const ctrlPrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, metaKey: false, preventDefault: ctrlPrevent });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    expect(ctrlPrevent).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('Cmd(meta)+wheel also zooms', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: true, preventDefault() {} });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('positive deltaY (ctrl+wheel down) zooms out', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: 100, ctrlKey: true, metaKey: false, preventDefault() {} });
    expect(v.scaleForTest()).toBeLessThan(before);
    v.destroy();
  });

  it('enableZoom:false installs no wheel handler — ctrl+wheel is inert', () => {
    const { v, scrollHost } = setup(20, { enableZoom: false });
    const before = v.scaleForTest();
    // No wheel listener was registered at all.
    expect(scrollHost._listeners.has('wheel')).toBe(false);
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    expect(prevent).not.toHaveBeenCalled();
    v.destroy();
  });

  it('a wheel with deltaY 0 is a no-op even with ctrl held', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: 0, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    v.destroy();
  });

  it('setScale to the SAME scale is a no-op (no re-anchor, no epoch bump churn)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 390;
    scrollHost.dispatch('scroll');
    const topBefore = scrollHost.scrollTop;
    const epochBefore = v.renderEpochForTest();
    v.setScale(v.scaleForTest()); // identical scale
    expect(scrollHost.scrollTop).toBe(topBefore); // unchanged
    expect(v.renderEpochForTest()).toBe(epochBefore); // no epoch bump
    v.destroy();
  });
});

describe('PptxScrollViewer — self-load path (T7 story)', () => {
  it('load(url) lays out and mounts the first window (relayout wired into load)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Mock the static loader so load() resolves to a fake engine WITHOUT touching
    // a real Worker / WASM. The viewer must call relayout() after assignment so a
    // self-loaded (non-injected) viewer is not left blank (I-2).
    const engine = new FakePptxEngine(10, SLIDE_W_EMU, SLIDE_H_EMU);
    const loadSpy = vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { gap: 10 });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    await v.load('sample.pptx');
    expect(loadSpy).toHaveBeenCalledTimes(1);
    // Layout happened: slots mounted and the spacer was sized.
    expect(v.mountedSlideIndicesForTest().length).toBeGreaterThan(0);
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // The mounted slides were actually rendered (main mode → renderSlide).
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    v.destroy();
  });
});

describe('PptxScrollViewer — text selection (T5)', () => {
  // A minimal PptxTextRunInfo: richer than docx's flat run — it carries per-shape
  // frame geometry (shapeX/Y/W/H, rotation) so buildPptxTextLayer can group runs
  // into a positioned shape <div> and nest the run <span> inside it.
  const RUN = {
    text: 'Hi',
    inShapeX: 1,
    inShapeY: 2,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    shapeX: 0,
    shapeY: 0,
    shapeW: 100,
    shapeH: 40,
    rotation: 0,
  };

  function findSlotWrapper(scrollHost: FakeEl): FakeEl {
    return scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
  }
  function findTextLayer(slotWrapper: FakeEl): FakeEl {
    return slotWrapper.children.find((c) => c.tag === 'div') as FakeEl;
  }

  it('main mode: builds a shape div with a nested run span for each visible slot', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Every mounted slot got its overlay built from that slide's render.
    const mounted = v.mountedSlideIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = findTextLayer(wrapper);
      // pptx nests spans INSIDE a per-shape <div> (one level deeper than docx):
      // textLayer → shape <div> → run <span>.
      expect(textLayer.children.length).toBe(1);
      const shapeDiv = textLayer.children[0] as FakeEl;
      expect(shapeDiv.tag).toBe('div');
      expect(shapeDiv.children.length).toBe(1);
      expect(shapeDiv.children[0].tag).toBe('span');
      expect(shapeDiv.children[0].textContent).toBe('Hi');
    }
    v.destroy();
  });

  it('main mode: a rotated shape gets a CSS rotate() on its shape div', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    // Rotated shape frame: buildPptxTextLayer applies transform:rotate(totalRot).
    engine.feedTextRuns = [{ ...RUN, rotation: 30 }];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    const shapeDiv = textLayer.children[0] as FakeEl;
    expect(shapeDiv.tag).toBe('div');
    expect(shapeDiv.style.transform).toBe('rotate(30deg)');
    // The run span still nests inside the rotated shape div.
    expect(shapeDiv.children[0].textContent).toBe('Hi');
    v.destroy();
  });

  it('main mode: sizes the overlay to the slot CSS box (literal px, dpr 1)', async () => {
    installDom(); // window.devicePixelRatio = 1
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
      paddingLeft: 0, // full-width fit → slide 200×120px (literal sizes asserted below)
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    // The overlay container is NOT pinned to a literal px size: it keeps its
    // creation `width:100%;height:100%` so it tracks the wrapper's actual box
    // even when a consumer scales the canvas down with external CSS (the reported
    // scrollbar bug). The wrapper IS sized to the CSS box (200×120 at base scale
    // 1.0), NOT the canvas backing store — so `100%` resolves to the CSS box.
    expect(textLayer.style.width).toBe('100%');
    expect(textLayer.style.height).toBe('100%');
    expect(wrapper.style.width).toBe('200px');
    expect(wrapper.style.height).toBe('120px');
    v.destroy();
  });

  it('main mode (dpr 2): overlay is the CSS box, NOT the 2× backing store', async () => {
    // Finding B regression: on retina (dpr 2) the real renderer sets canvas.width =
    // cssWidth × 2, so a slot canvas is 400×240 for a 200×120 CSS box. The overlay
    // must stay 200×120 (the CSS box); copying canvas.width would size it 2× and
    // overflow the wrapper (inflating the scroll area). The fake renderSlide mirrors
    // the real backing-store sizing when dpr is passed.
    installDom();
    vi.stubGlobal('window', { devicePixelRatio: 2 });
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(1, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
      paddingLeft: 0, // full-width fit → 200×120 CSS box / 400×240 backing store
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const canvas = wrapper.children.find((c) => c.tag === 'canvas') as FakeEl;
    const textLayer = findTextLayer(wrapper);
    // The canvas backing store IS 2× (dpr 2): 400×240.
    expect(canvas.width).toBe(400);
    expect(canvas.height).toBe(240);
    // …but the overlay container stays `100%` of the wrapper, and the wrapper is
    // the CSS box (200×120) — half the backing store. `100%` of that box never
    // copies the 2× backing store, so the overlay does not overflow the wrapper /
    // inflate the scroll area.
    expect(textLayer.style.width).toBe('100%');
    expect(textLayer.style.height).toBe('100%');
    expect(wrapper.style.width).toBe('200px');
    expect(wrapper.style.height).toBe('120px');
    v.destroy();
  });

  it('main mode: recycling a slot clears its overlay so the free pool holds no stale spans', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, SLIDE_W_EMU, SLIDE_H_EMU);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderSlide .then() microtask.
    await Promise.resolve();
    await Promise.resolve();
    // A slide-0 slot exists with a built overlay (one shape div).
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    expect(textLayer.children.length).toBe(1);
    // Scroll far away so slide 0 recycles into the free pool.
    scrollHost.scrollTop = 8000;
    scrollHost.dispatch('scroll');
    // The recycled slot's overlay was cleared (no stale spans linger in the pool).
    expect(textLayer.children.length).toBe(0);
    v.destroy();
  });

  it('worker mode: builds a shape div with a nested run span (IX6, no warning)', async () => {
    installDom();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU, 'worker');
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderSlideToBitmap resolution microtask; give
    // the bitmap promise chain a couple of turns to settle.
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // IX6 — worker mode no longer warns; the runs ride back beside the bitmap so
    // the overlay is populated identically to main mode.
    expect(warn).not.toHaveBeenCalled();
    // The worker path (renderSlideToBitmap), NOT renderSlide, was used.
    expect(engine.renderCalls.length).toBe(0);
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    const mounted = v.mountedSlideIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = wrapper.children.find((c) => c.tag === 'div') as FakeEl;
      // Same nesting as main mode: textLayer → shape <div> → run <span>.
      expect(textLayer.children.length).toBe(1);
      const shapeDiv = textLayer.children[0] as FakeEl;
      expect(shapeDiv.tag).toBe('div');
      expect(shapeDiv.children.length).toBe(1);
      expect(shapeDiv.children[0].tag).toBe('span');
      expect(shapeDiv.children[0].textContent).toBe('Hi');
    }
    warn.mockRestore();
    v.destroy();
  });

  it('stale main render (epoch moved by setScale mid-flight) does NOT rebuild the overlay', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred main mode: the test resolves renderSlide manually so setScale can
    // move the epoch WHILE slide 0's first render is in flight.
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU, 'main', true);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
      // REAL defaults [0.1, 4]: base fit is 1.0, and setScale(×2)=2.0 is inside them
      // (the old `zoomMin: 0.05` dodged the px-per-EMU unit bug — WS4b-2).
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // Grab the first (stale) slide-0 render call before resolving it. Capture the
    // slot wrapper it targets so we can inspect its overlay afterwards.
    const staleWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const staleLayer = staleWrapper.children.find((c) => c.tag === 'div') as FakeEl;
    const staleCall = engine.renderCalls.find((c) => c.slide === 0)!;
    // Zoom mid-flight: bumps the epoch and force-re-mounts every slot (which
    // re-dispatches a fresh slide-0 render at the new scale).
    v.setScale(v.scaleForTest() * 2);
    // Now resolve the STALE (old-epoch) render. Its .then must bail on the epoch
    // guard and NOT build the overlay.
    staleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // The stale render built nothing (epoch guard). The fresh render is still
    // deferred (not resolved), so its overlay isn't built either — the stale
    // slot/layer holds zero children from the superseded render.
    expect(staleLayer.children.length).toBe(0);
    v.destroy();
  });
});

describe('PptxScrollViewer — navigation, resize, empty (T6)', () => {
  // container fit width 200, natural slide width 200px → DIMENSIONLESS base scale
  // 200/200 = 1.0. slide height px = (SLIDE_H_EMU / 9525) * 1.0 = 120. gap 10 ⇒
  // stride 130. (Only the base VALUE moved 0.2 → 1.0; the px geometry is identical.)
  const BASE = 1.0;
  const SLIDE_H = SLIDE_H_EMU / 9525; // 120 (natural height px at base 1.0)
  const GAP = 10;
  const STRIDE = SLIDE_H + GAP; // 130 — the top-to-top distance between slides

  function setup(slideCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: GAP,
      // Flush top/bottom: the T6 STRIDE/offset arithmetic assumes offset[0]===0.
      // paddingTop/Bottom default to `gap`; pinning them to 0 keeps the pre-padding
      // geometry (and exercises the "explicit 0 ⇒ old flush behavior" contract).
      paddingTop: 0,
      paddingBottom: 0,
      // Flush left/right: the fit width must be the full 200 container so BASE = 1.0
      // and SLIDE_H/STRIDE hold (the default `gap` gutters would shrink the fit).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  it('scrollToSlide sets scrollTop to the slide offset and clamps out-of-range', () => {
    const { v, scrollHost } = setup();
    // Pick a clientHeight that makes maxTop STRICTLY LESS than the last slide's top
    // offset, so the `Math.min(maxTop, target)` clamp is actually exercised.
    // totalHeight = 20*120 + 19*10 = 2590; offsets[19] = 19*130 = 2470.
    scrollHost.clientHeight = 200; // maxTop = 2590 − 200 = 2390 < offsets[19] 2470
    v.scrollToSlide(3);
    expect(Math.abs(scrollHost.scrollTop - 3 * STRIDE)).toBeLessThan(2);
    v.scrollToSlide(999); // clamps to last slide index (19)
    const totalHeight = 20 * SLIDE_H + 19 * GAP; // 2590
    const maxTop = totalHeight - 200; // 2390
    const lastOffset = 19 * STRIDE; // 2470
    expect(maxTop).toBeLessThan(lastOffset); // precondition: clamp is exercised
    expect(Math.abs(scrollHost.scrollTop - maxTop)).toBeLessThan(2);
    v.scrollToSlide(-5); // clamps to 0
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('scrollToSlide: viewport taller than total content ⇒ scrollTop pinned to 0 (negative maxTop guard)', () => {
    // A 2-slide deck is shorter than a very tall viewport, so
    // totalHeight − clientHeight is NEGATIVE; the `Math.max(0, …)` maxTop guard must
    // pin scrollTop to 0 rather than a negative top.
    const { v, scrollHost } = setup(2);
    scrollHost.clientHeight = 5000; // >> totalHeight (2*120 + 10 = 250)
    v.scrollToSlide(1); // last slide; its offset (130) is above maxTop (0)
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('onVisibleSlideChange fires only when topIndex changes', () => {
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    scrollHost.scrollTop = STRIDE; // slide 1 top
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = STRIDE + 10; // still within slide 1
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 3 * STRIDE; // slide 3 top
    scrollHost.dispatch('scroll');
    // Deduped: 0 (initial) → 1 → 3, no repeat for the STRIDE+10 no-op.
    expect(changes).toEqual([0, 1, 3]);
    v.destroy();
  });

  it('scrollToSlide does not re-fire onVisibleSlideChange for the same top slide', () => {
    const changes: number[] = [];
    const { v } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    // Initial mount fired 0. scrollToSlide(0) lands on the same top slide — no new fire.
    v.scrollToSlide(0);
    expect(changes).toEqual([0]);
    // Navigate to slide 5 → one fire; navigate there again → no duplicate.
    v.scrollToSlide(5);
    v.scrollToSlide(5);
    expect(changes).toEqual([0, 5]);
    v.destroy();
  });

  it('ResizeObserver re-fit preserves the zoom multiplier', () => {
    // zoomMax headroom: base 1.0, 2× zoom → 2.0; after the width doubles the new
    // base is 2.0 so the preserved scale is 4.0. zoomMin/zoomMax are ABSOLUTE
    // dimensionless bounds (design §8.2), so a low zoomMax would legitimately CLAMP
    // the re-fit and break preservation. Give the multiplier room to survive.
    const { v, container } = setup(20, { zoomMax: 10 });
    v.setScale(v.baseScaleForTest() * 2); // 2× zoom
    const zoomMultiplier = v.scaleForTest() / v.baseScaleForTest();
    // Container widens; the observer callback re-fits base and re-applies the mult.
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest(); // fires the observed resize path
    expect(v.scaleForTest() / v.baseScaleForTest()).toBeCloseTo(zoomMultiplier, 5);
    v.destroy();
  });

  it('ResizeObserver re-fit bumps the render epoch (resize is an epoch event)', () => {
    const { v, container } = setup();
    const before = v.renderEpochForTest();
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest();
    // A width change re-fits the base scale (routed through setScale), which bumps
    // the epoch so any bitmap in flight at the old scale is treated as stale.
    expect(v.renderEpochForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('ResizeObserver re-fit with an UNCHANGED width does not re-fit or bump the epoch', () => {
    const { v } = setup();
    const scaleBefore = v.scaleForTest();
    const epochBefore = v.renderEpochForTest();
    v.resizeForTest(); // same width — no-op re-fit
    expect(v.scaleForTest()).toBe(scaleBefore);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    v.destroy();
  });

  it('height-only resize (width unchanged) re-mounts the newly-revealed window without a scroll, no epoch bump, no callback', () => {
    // F1 regression: a height-only grow used to early-return before `_mountVisible`,
    // leaving the newly-revealed rows blank until the user scrolled. At scrollTop 0
    // the initially-mounted window is [0..3] (viewport 400 / stride 130 ≈ 3 slides;
    // +1 overscan). Growing the viewport must mount more purely from the resize —
    // no scroll event, no epoch bump (geometry/scale unchanged so cached canvases
    // stay valid), and no onVisibleSlideChange fire (topIndex stays 0).
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    const mountedBefore = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore.length).toBeGreaterThan(0);
    expect(changes).toEqual([0]); // initial mount fired 0
    const epochBefore = v.renderEpochForTest();
    const scaleBefore = v.scaleForTest();

    // Grow HEIGHT only; width (and thus the fit-width base scale) is unchanged.
    scrollHost.clientHeight = 1600;
    v.resizeForTest(); // NO scroll event dispatched

    const mountedAfter = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    // The revealed window strictly grows (taller viewport intersects more slides).
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    expect(Math.max(...mountedAfter)).toBeGreaterThan(Math.max(...mountedBefore));
    // No epoch bump — nothing was re-scaled, so no in-flight render is stale.
    expect(v.renderEpochForTest()).toBe(epochBefore);
    expect(v.scaleForTest()).toBe(scaleBefore);
    // topIndex is still 0 at scrollTop 0 ⇒ callback did NOT fire again.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('clamped resize (zoomMax clamp + width & height grow) still refreshes the revealed window', () => {
    // F1 clamped-path variant: pin zoomMax to the current base so a width grow's
    // re-fit (newBase × mult) clamps back to the SAME scale, making setScale a no-op
    // (which skips its own force-re-render mount). The post-setScale `_mountVisible`
    // must still mount the rows the taller viewport revealed. base = 1.0, no user
    // zoom ⇒ mult 1; after width→400 newBase 2.0 clamps to zoomMax 1.0 (== _scale).
    const changes: number[] = [];
    const { v, scrollHost, container } = setup(20, {
      zoomMax: BASE, // 1.0 — clamps the re-fit back to the current scale
      onVisibleSlideChange: (i: number) => changes.push(i),
    });
    expect(v.scaleForTest()).toBe(BASE);
    const mountedBefore = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore.length).toBeGreaterThan(0);
    const epochBefore = v.renderEpochForTest();

    // Grow BOTH width and height. The width grow triggers the re-fit branch, but the
    // clamp pins the scale unchanged ⇒ setScale no-ops. Heights are therefore
    // unchanged (scale unchanged), so the taller viewport reveals more of the SAME
    // geometry.
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    scrollHost.clientWidth = 400;
    scrollHost.clientHeight = 1600;
    v.resizeForTest();

    const mountedAfter = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    // Scale stayed pinned at the clamp, so setScale no-oped ⇒ no epoch bump.
    expect(v.scaleForTest()).toBe(BASE);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    // topIndex still 0 ⇒ no extra callback.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('zero-width container defers, then lays out on first non-zero resize', () => {
    installDom();
    const container = makeContainer(0, 0); // unlaid-out
    const engine = new FakePptxEngine(5, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    expect(v.mountedSlideIndicesForTest().length).toBe(0); // deferred
    container.clientWidth = 300;
    scrollHost.clientWidth = 300;
    scrollHost.clientHeight = 400;
    v.resizeForTest();
    expect(v.mountedSlideIndicesForTest().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('empty deck: slideCount 0 ⇒ spacer 0, no slots, scrollToSlide no-op', () => {
    installDom();
    const container = makeContainer(300, 400);
    const engine = new FakePptxEngine(0, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    v.relayout();
    const spacer = (container.children[0] as FakeEl).children[0].children[0] as FakeEl;
    expect(parseFloat(spacer.style.height || '0')).toBe(0);
    expect(v.mountedSlideIndicesForTest().length).toBe(0);
    v.scrollToSlide(0); // no throw
    v.destroy();
  });

  it('empty deck: resize does not crash and fires no callback', () => {
    installDom();
    const container = makeContainer(0, 0);
    const engine = new FakePptxEngine(0, SLIDE_W_EMU, SLIDE_H_EMU);
    const changes: number[] = [];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      onVisibleSlideChange: (i: number) => changes.push(i),
    });
    container.clientWidth = 300;
    (container.children[0] as FakeEl).children[0].clientWidth = 300;
    v.resizeForTest(); // no slides ⇒ no-op
    expect(v.mountedSlideIndicesForTest().length).toBe(0);
    expect(changes).toEqual([]);
    v.destroy();
  });

  it('destroy() disconnects the ResizeObserver', () => {
    installDom();
    let disconnected = 0;
    class SpyRO {
      cb: () => void;
      constructor(cb: () => void) {
        this.cb = cb;
      }
      observe(): void {}
      disconnect(): void {
        disconnected++;
      }
    }
    vi.stubGlobal('ResizeObserver', SpyRO);
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    v.destroy();
    expect(disconnected).toBe(1);
  });
});

describe('PptxScrollViewer — paddingTop/paddingBottom (desk margin)', () => {
  const SLIDE_H = SLIDE_H_EMU / 9525; // 120 (natural height px at base 1.0)
  const GAP = 10;
  const STRIDE = SLIDE_H + GAP; // 130

  function setup(opts = {}, slideCount = 20) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: GAP,
      // This block tests the VERTICAL desk margin; pin the horizontal gutters to 0
      // so the fit width is the full 200 container ⇒ BASE 1.0 / SLIDE_H 120 / STRIDE
      // 130 (the default `gap` gutters would otherwise shrink the fit).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  /** The wrapper currently mounted for slide `i` (identified by its top offset). */
  function slotTopFor(scrollHost: FakeEl, i: number, stride: number, padTop: number): FakeEl | undefined {
    const want = `${padTop + i * stride}px`;
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === want,
    ) as FakeEl | undefined;
  }

  it('explicit paddingTop mounts the first slot at top = paddingTop px', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    expect(slotTopFor(scrollHost, 0, STRIDE, 24)).toBeDefined();
    v.destroy();
  });

  it('spacer height = padTop + Σheights + (n-1)*gap + padBottom', () => {
    const { scrollHost } = setup({ paddingTop: 24, paddingBottom: 40 });
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = 24 + 20 * SLIDE_H + 19 * GAP + 40;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
  });

  it('DEFAULT paddingTop/paddingBottom = gap (uniform rhythm: no options ⇒ first slot at gap px)', () => {
    // No paddingTop/paddingBottom → each defaults to `gap` (10). Slide 0 sits at
    // top:10px, NOT flush 0 (this is the sanctioned pre-release default change).
    const { v, scrollHost } = setup();
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    expect(slotTopFor(scrollHost, 0, STRIDE, GAP)).toBeDefined();
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = GAP + 20 * SLIDE_H + 19 * GAP + GAP;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
  });

  it('explicit 0 ⇒ flush (old behavior reachable): first slot at top 0, spacer has no pad', () => {
    const { v, scrollHost } = setup({ paddingTop: 0, paddingBottom: 0 });
    expect(slotTopFor(scrollHost, 0, STRIDE, 0)).toBeDefined();
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = 20 * SLIDE_H + 19 * GAP; // no pad
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
  });

  it('scrollToSlide(0) lands on offsets[0] (= paddingTop, not 0)', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    scrollHost.scrollTop = 3 * STRIDE;
    scrollHost.dispatch('scroll');
    v.scrollToSlide(0);
    expect(scrollHost.scrollTop).toBe(24);
    v.destroy();
  });

  it('scrollToSlide(k) lands on paddingTop + k*stride', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    v.scrollToSlide(3);
    expect(Math.abs(scrollHost.scrollTop - (24 + 3 * STRIDE))).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchor keeps the slide under the viewport top fixed WITH padding intact after setScale', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24, zoomMin: 0.5, zoomMax: 3 });
    // Scroll so slide 3's top sits at the viewport top: offset[3] = 24 + 3*130 = 414.
    scrollHost.scrollTop = 24 + 3 * STRIDE;
    scrollHost.dispatch('scroll');
    expect(v.topVisibleSlide).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 1.0 → 2.0; SLIDE_H' = 240, STRIDE' = 250
    expect(v.scaleForTest()).toBeCloseTo(2, 5);
    // Slide 3 stays the top slide; padding is intact so offset'[3] = 24 + 3*250 = 774.
    expect(v.topVisibleSlide).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - (24 + 3 * 250))).toBeLessThan(2);
    v.destroy();
  });
});

describe('PptxScrollViewer — paddingLeft/paddingRight (horizontal desk gutters)', () => {
  // Slides are UNIFORM: natural width px = SLIDE_W_EMU / 9525 = 200, height 120.

  /** Build a viewer over a container of `cw`×400. */
  function setup(cw: number, opts = {}, slideCount = 20) {
    installDom();
    const container = makeContainer(cw, 400);
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 16,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = cw;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  /** The wrapper mounted at CSS `top`. */
  function slot(scrollHost: FakeEl, top: string): FakeEl | undefined {
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === top,
    ) as FakeEl | undefined;
  }

  it('fit subtracts the DEFAULT gutters: container 232 with default gap-16 gutters ⇒ fit 200 ⇒ base 1.0', () => {
    // 232 − 16 (padL) − 16 (padR) = 200, the same fit the old 200-container tests
    // used, so the dimensionless base scale matches their 1.0 (natural width 200).
    const { v } = setup(232);
    expect(v.baseScaleForTest()).toBeCloseTo(1.0, 6);
    expect(v.scaleForTest()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('explicit paddingLeft:0/paddingRight:0 ⇒ old full-width fit reachable (base uses the FULL container width)', () => {
    // Gutters pinned to 0 ⇒ fit is the full 200 container ⇒ base 1.0, exactly the
    // pre-feature behavior.
    const { v } = setup(200, { paddingLeft: 0, paddingRight: 0 });
    expect(v.baseScaleForTest()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('explicit opts.width is NOT reduced by the gutters (it is the slide CSS-width contract)', () => {
    // opts.width 150 stays 150 regardless of the gutters: base = 150 / 200 (natural)
    // = 0.75. The gutters still apply to PLACEMENT, just not to the width.
    const { v } = setup(400, { width: 150, paddingLeft: 24, paddingRight: 24 });
    expect(v.baseScaleForTest()).toBeCloseTo(0.75, 6);
    v.destroy();
  });

  it('slot left = paddingLeft when the slide fills the fit exactly (symmetric gutters ⇒ centre lands on the floor)', () => {
    // Container 232, default gutters 16 ⇒ fit 200, slide px = 200. scrollHost
    // clientWidth 232, so centre = (232 − 200)/2 = 16 = padL ⇒ left pinned at padL.
    const { v, scrollHost } = setup(232);
    // Default vertical padding = gap 16 ⇒ slide 0 sits at top:16px.
    const s0 = slot(scrollHost, '16px');
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('16px'); // = paddingLeft
    v.destroy();
  });

  it('slot is CENTERED when the slide is narrower than the viewport (left = (cw − sw)/2 > paddingLeft)', () => {
    // Explicit width 100 (NOT reduced) ⇒ slide px = 100, far narrower than the 400
    // viewport. centre = (400 − 100)/2 = 150, which exceeds padL 16, so the slide is
    // centred at 150 rather than pinned to the gutter floor.
    const { v, scrollHost } = setup(400, { width: 100, paddingLeft: 16, paddingRight: 16 });
    const s0 = slot(scrollHost, '16px'); // default vertical pad = gap 16
    expect(s0).toBeDefined();
    expect(parseFloat(s0!.style.left)).toBeCloseTo(150, 3);
    expect(parseFloat(s0!.style.left)).toBeGreaterThan(16);
    v.destroy();
  });

  it('zoomed-in (slide wider than viewport): left pins to paddingLeft and the spacer width = slideW + padL + padR', () => {
    // Explicit width 400 in a 200 viewport ⇒ slide px 400 > cw 200. centre =
    // (200 − 400)/2 = −100, so the floor pins left at padL 24. The horizontal scroll
    // extent (spacer width) is the slide width plus both gutters: 400 + 24 + 24 = 448.
    const { v, scrollHost } = setup(200, { width: 400, paddingLeft: 24, paddingRight: 24 });
    const s0 = slot(scrollHost, '16px'); // default vertical pad = gap 16
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('24px'); // pinned to paddingLeft
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(400 + 24 + 24, 3);
    v.destroy();
  });

  it('a Ctrl+wheel zoom that widens the slide past the viewport updates the spacer width and pins left to paddingLeft', () => {
    // Container 200 pinned gutters 0 ⇒ fit 200, base 1.0, slide px 200 (== viewport).
    // Zoom ×2 ⇒ slide px 400 > 200 viewport. left floors to padL (0 here) and the
    // spacer width tracks the new slide width + gutters (400 + 0 + 0 = 400).
    const { v, scrollHost } = setup(200, {
      paddingLeft: 0,
      paddingRight: 0,
      paddingTop: 0,
      zoomMin: 0.5,
      zoomMax: 4,
    });
    expect(v.scaleForTest()).toBeCloseTo(1.0, 6);
    v.setScale(v.scaleForTest() * 2); // slide px 200 → 400
    const s0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('0px'); // paddingLeft 0
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(400, 3);
    v.destroy();
  });

  it('spacer width = slide width + both gutters (uniform slides)', () => {
    // Container 200, gutters 12 ⇒ fit 176, base 0.88, slide px = 176. Spacer width =
    // 176 + 12 + 12 = 200 (== container; a spacer that just fits creates no scrollbar).
    const { v, scrollHost } = setup(200, { paddingLeft: 12, paddingRight: 12 });
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(176 + 12 + 12, 3);
    void v;
    v.destroy();
  });

  it('gutters wider than the container defer layout (non-positive fit ⇒ zero-width deferral)', () => {
    // padL + padR = 300 > container 200 ⇒ fit ≤ 0 ⇒ deferred (no slots mounted),
    // mirroring the zero-width container deferral.
    const { v } = setup(200, { paddingLeft: 150, paddingRight: 150 });
    expect(v.mountedSlideIndicesForTest().length).toBe(0);
    v.destroy();
  });

  it('the slot wrapper carries NO CSS auto-centering (explicit JS left replaces margin:0 auto)', () => {
    // The old wrapper used `left:0;right:0;margin:0 auto`; that would fight the
    // explicit per-mount `left`. Assert the auto-centering margin is gone.
    const { v, scrollHost } = setup(232);
    const s0 = slot(scrollHost, '16px');
    expect(s0).toBeDefined();
    expect(s0!.style.margin).toBe(''); // no `0 auto`
    expect(s0!.style.right).toBe(''); // no right:0 pinning
    v.destroy();
  });
});

describe('PptxScrollViewer — flicker-free zoom (T8)', () => {
  // Flicker-free zoom (design §7): setScale must NOT blank a visible slide.
  // Three mechanisms — CSS preview (immediate, no device-buffer resize / no
  // recycle of in-window slots), settle re-render (debounced ZOOM_SETTLE_MS),
  // double-buffer swap (main-mode settle renders into a SPARE off-DOM canvas and
  // swaps it in). Container 200×400, natural slide 200×120px, flush pads ⇒ base
  // 1.0, slide height px 120.
  function setup(slideCount = 20, opts = {}, mode: 'main' | 'worker' = 'main', deferred = false) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU, mode, deferred);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 4,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  /** The wrapper mounted for the slide currently at top:`top`px. */
  function slotAtTop(scrollHost: FakeEl, top: string): FakeEl | undefined {
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === top,
    ) as FakeEl | undefined;
  }

  // (a) setScale does NOT unmount in-window slots (same wrapper before/after) and
  //     does NOT resize any on-screen canvas device buffer during the preview.
  it('CSS preview: setScale keeps the SAME in-window slot wrappers and never resizes their device buffer', () => {
    const { v, scrollHost } = setup();
    const before = slotAtTop(scrollHost, '0px');
    expect(before).toBeDefined();
    const beforeCanvas = before!.children.find((k) => k.tag === 'canvas') as FakeEl;
    const resizesBefore = beforeCanvas._deviceResizes.length;

    v.setScale(v.scaleForTest() * 2); // zoom in

    const after = slotAtTop(scrollHost, '0px');
    expect(after).toBe(before); // same wrapper (not recycled + re-mounted)
    const afterCanvas = after!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(afterCanvas).toBe(beforeCanvas); // same canvas element
    expect(afterCanvas._deviceResizes.length).toBe(resizesBefore); // no device resize
    v.destroy();
  });

  it('CSS preview: setScale CSS-resizes the slot canvas (style.width/height) to the new layout size immediately', () => {
    const { v, scrollHost } = setup();
    const slot = slotAtTop(scrollHost, '0px')!;
    const canvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    // base 1.0 ⇒ slide px 200×120. Zoom ×2 ⇒ CSS 400×240.
    v.setScale(v.scaleForTest() * 2);
    expect(canvas.style.width).toBe('400px');
    expect(canvas.style.height).toBe('240px');
    v.destroy();
  });

  it('CSS preview: text layer gets a transform: scale(ratio) matching newScale / renderedScale', async () => {
    const { v, scrollHost, engine } = setup(20, { enableTextSelection: true });
    engine.feedTextRuns = [
      {
        text: 'Hi',
        inShapeX: 1,
        inShapeY: 2,
        w: 10,
        h: 12,
        fontSize: 12,
        font: '12px serif',
        shapeX: 0,
        shapeY: 0,
        shapeW: 100,
        shapeH: 40,
        rotation: 0,
      },
    ];
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const textLayer = slot.children.find((k) => k.tag === 'div') as FakeEl;
    // Zoom ×2 — overlay built at base 1.0; preview scales by newScale/renderedScale
    // = 2.0/1.0 = 2.
    v.setScale(v.scaleForTest() * 2);
    expect(textLayer.style.transform).toBe('scale(2)');
    expect(textLayer.style.transformOrigin).toBe('0 0');
    v.destroy();
  });

  // (c) NO render dispatch during a burst; ONE dispatch after ZOOM_SETTLE_MS.
  it('debounce: a burst of setScale calls dispatches NO settle render until ZOOM_SETTLE_MS elapses, then exactly one per slot', () => {
    vi.useFakeTimers();
    try {
      const { v, engine } = setup();
      const dispatchesBefore = engine.renderCalls.length;
      v.setScale(v.scaleForTest() * 1.1);
      vi.advanceTimersByTime(50);
      v.setScale(v.scaleForTest() * 1.1);
      vi.advanceTimersByTime(50);
      v.setScale(v.scaleForTest() * 1.1);
      vi.advanceTimersByTime(50); // timer reset each tick — still inside window
      expect(engine.renderCalls.length).toBe(dispatchesBefore);
      vi.advanceTimersByTime(150); // past the settle timeout from the LAST setScale
      const mounted = v.mountedSlideIndicesForTest();
      const settleRenders = engine.renderCalls.length - dispatchesBefore;
      expect(settleRenders).toBe(mounted.length);
    } finally {
      vi.useRealTimers();
    }
  });

  // (d) settle resolution swaps the canvas without a blank (main mode double-buffer).
  it('double-buffer (main): settle renders into a SPARE canvas and swaps it in — the on-screen canvas is never device-resized in place', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup(20, {}, 'main', true);
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
      const onScreenUid = onScreenCanvas._uid;
      const resizesOnScreenBefore = onScreenCanvas._deviceResizes.length;

      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200);

      const settleCall = engine.renderCalls.filter((c) => c.slide === 0).pop()!;
      expect(settleCall.canvas).toBeDefined();
      expect(settleCall.canvas!._uid).not.toBe(onScreenUid); // rendered into a spare
      expect(onScreenCanvas._deviceResizes.length).toBe(resizesOnScreenBefore); // not in-place
      expect(slot.children).toContain(onScreenCanvas); // still attached (no swap yet)

      settleCall.resolve();
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          expect(nowCanvas._uid).toBe(settleCall.canvas!._uid); // spare swapped in
          expect(slot.children).not.toContain(onScreenCanvas); // old retired
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('double-buffer (worker): settle renders into the SAME canvas (sync resize+transfer paints no blank frame)', async () => {
    vi.useFakeTimers();
    const { v, scrollHost, engine } = setup(20, {}, 'worker', true);
    for (const c of engine.bitmapCalls.slice()) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const canvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    const canvasUid = canvas._uid;
    const bitmapsBefore = engine.bitmapCalls.length;

    v.setScale(v.scaleForTest() * 2);
    vi.advanceTimersByTime(200);

    expect(engine.bitmapCalls.length).toBeGreaterThan(bitmapsBefore);
    const settleCall = engine.bitmapCalls.filter((c) => c.slide === 0).pop()!;
    const settleBitmap = engine.createdBitmaps[engine.bitmapCalls.indexOf(settleCall)];
    settleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(nowCanvas._uid).toBe(canvasUid); // SAME canvas (worker keeps it)
    expect(nowCanvas._bitmapCtx?.lastBitmap).toBe(settleBitmap);
    vi.useRealTimers();
    v.destroy();
  });

  // (e) stale settle discards per the epoch gate.
  it('stale settle: an epoch bump during the settle render discards the settle (no swap, spare not attached)', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup(20, {}, 'main', true);
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;

      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200); // settle dispatched (deferred, in flight)
      const settleCall = engine.renderCalls.filter((c) => c.slide === 0).pop()!;

      v.setScale(v.scaleForTest() * 1.1); // bump epoch mid-settle-render

      settleCall.resolve();
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          expect(nowCanvas).toBe(onScreenCanvas); // stale spare NOT swapped in
          expect(nowCanvas._uid).not.toBe(settleCall.canvas!._uid);
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  // (f) destroy during a pending settle timer / in-flight settle render.
  it('destroy during a pending settle timer clears it — no post-destroy render dispatch, no crash', () => {
    vi.useFakeTimers();
    try {
      const { v, engine } = setup();
      const dispatchesBefore = engine.renderCalls.length;
      v.setScale(v.scaleForTest() * 2); // schedules a settle timer
      v.destroy();
      vi.advanceTimersByTime(500);
      expect(engine.renderCalls.length).toBe(dispatchesBefore);
    } finally {
      vi.useRealTimers();
    }
  });

  it('destroy during an in-flight settle render does not swap or fire onError post-destroy', () => {
    vi.useFakeTimers();
    const onError = vi.fn();
    try {
      const { v, scrollHost, engine } = setup(20, { onError }, 'main', true);
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200);
      const settleCall = engine.renderCalls.filter((c) => c.slide === 0).pop()!;
      v.destroy();
      settleCall.resolve();
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          expect(onError).not.toHaveBeenCalled();
          expect(slot.children).toContain(onScreenCanvas);
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('settle clears the preview transform on the text layer (overlay rebuilt at full resolution)', async () => {
    vi.useFakeTimers();
    const { v, scrollHost, engine } = setup(20, { enableTextSelection: true }, 'main', true);
    engine.feedTextRuns = [
      {
        text: 'Hi',
        inShapeX: 1,
        inShapeY: 2,
        w: 10,
        h: 12,
        fontSize: 12,
        font: '12px serif',
        shapeX: 0,
        shapeY: 0,
        shapeW: 100,
        shapeH: 40,
        rotation: 0,
      },
    ];
    for (const c of engine.renderCalls.slice()) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const textLayer = slot.children.find((k) => k.tag === 'div') as FakeEl;

    v.setScale(v.scaleForTest() * 2);
    expect(textLayer.style.transform).toBe('scale(2)'); // preview transform applied
    vi.advanceTimersByTime(200);
    const settleCall = engine.renderCalls.filter((c) => c.slide === 0).pop()!;
    settleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(textLayer.style.transform).toBe(''); // cleared after the crisp render
    vi.useRealTimers();
    v.destroy();
  });
});

describe('PptxScrollViewer — pageShadow (T9)', () => {
  // pageShadow paints a CSS box-shadow on every slide CANVAS (not the wrapper).
  // Default = the recipe drop shadow; `false` disables; a custom string is
  // honoured verbatim. It must reach EVERY canvas: fresh mounts, the settle
  // double-buffer SPARE, and recycled+re-mounted slots. Reuses the T8 setup
  // (deferred engine + fake timers) for the spare-canvas case.
  const DEFAULT_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';

  function setup(opts = {}, mode: 'main' | 'worker' = 'main', deferred = false) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU, mode, deferred);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 4,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine };
  }

  /** Every mounted slide canvas (each slot wrapper's canvas child). */
  function mountedCanvases(scrollHost: FakeEl): FakeEl[] {
    return scrollHost.children
      .filter((c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas'))
      .map((w) => w.children.find((k) => k.tag === 'canvas') as FakeEl);
  }

  it('applies the default recipe shadow to every mounted slide canvas', () => {
    const { v, scrollHost } = setup();
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
    v.destroy();
  });

  it('pageShadow:false sets no box-shadow on the slide canvas', () => {
    const { v, scrollHost } = setup({ pageShadow: false });
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe('');
    v.destroy();
  });

  it('honours a custom pageShadow string verbatim (e.g. a spread-ring border look)', () => {
    const ring = '0 0 0 1px #c8ccd0';
    const { v, scrollHost } = setup({ pageShadow: ring });
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(ring);
    v.destroy();
  });

  it('does NOT paint the shadow on the wrapper (only the canvas)', () => {
    const { v, scrollHost } = setup();
    const wrapper = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas'),
    ) as FakeEl;
    expect(wrapper.style.boxShadow).toBe('');
    v.destroy();
  });

  it('the settle double-buffer SPARE canvas carries the shadow (swapped-in canvas keeps it)', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup({}, 'main', true);
      // Resolve the initial mount renders so the on-screen canvases are painted.
      for (const c of engine.renderCalls.slice()) c.resolve();

      v.setScale(v.scaleForTest() * 2); // zoom → schedules a settle
      vi.advanceTimersByTime(200); // past ZOOM_SETTLE_MS → settle dispatched

      // The settle render for slide 0 targets a fresh SPARE canvas — it must
      // already carry the shadow BEFORE the swap.
      const settleCall = engine.renderCalls.filter((c) => c.slide === 0).pop()!;
      const spare = settleCall.canvas as unknown as FakeEl;
      expect(spare).toBeDefined();
      expect(spare.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);

      // Resolve → swap the spare in. The now-on-screen canvas still has the shadow.
      settleCall.resolve();
      const slot = scrollHost.children.find(
        (c) => c.tag === 'div' && c.style.top === '0px' && c.children.some((k) => k.tag === 'canvas'),
      ) as FakeEl;
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          expect(nowCanvas._uid).toBe(spare._uid); // spare swapped in
          expect(nowCanvas.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('a recycled + re-mounted slot still carries the shadow', () => {
    const { v, scrollHost } = setup();
    // Scroll far down so slide 0's slot recycles into the free pool, then back to
    // the top so a POOLED slot is re-mounted for slide 0.
    scrollHost.scrollTop = 5000;
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
    v.destroy();
  });
});

describe('PptxScrollViewer — barrel export (T7)', () => {
  it('is exported from the package entry', () => {
    expect(typeof pptxIndex.PptxScrollViewer).toBe('function');
  });
});
