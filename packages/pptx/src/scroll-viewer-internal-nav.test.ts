import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';
import type { PptxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX-nav wiring — PptxScrollViewer's internal-hyperlink click default.
 *
 * IX1 built the clickable overlay and resolved the external branch, but the
 * INTERNAL slide-jump branch was a documented no-op (the parser did not surface
 * a slide part → index map). IX-nav supplies that map on `PptxPresentation`
 * (`resolveInternalTarget`); these tests pin that the scroll viewer's default
 * click behaviour now (a) resolves an internal ref to a slide index via the
 * engine and scrolls to it, (b) passes the CURRENT top slide as the relative
 * base, (c) no-ops safely when the ref resolves to no slide, (d) fully delegates
 * to a caller-supplied `onHyperlinkClick` (with the resolved `slideIndex`
 * populated), and (e) still opens external links.
 *
 * The click is driven END TO END through the real overlay: a hyperlink text run
 * is fed to the viewer's `onTextRun`, `buildPptxTextLayer` makes a clickable span,
 * and the test fires its `click` listener — exactly the path a user click takes.
 */

const SLIDE_W_EMU = 9525 * 200; // natural 200px wide
const SLIDE_H_EMU = 9525 * 120; // natural 120px tall

function linkRun(hyperlink: HyperlinkTarget): PptxTextRunInfo {
  return {
    text: 'go',
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
    hyperlink,
  };
}

/** Mount a scroll viewer with one hyperlink run fed into the first slide's
 *  overlay, and return the pieces + a helper to click the link span. */
async function setup(
  slideCount: number,
  hyperlink: HyperlinkTarget,
  opts: Record<string, unknown> = {},
) {
  installDom();
  const container = makeContainer(200, 400);
  const engine = new FakePptxEngine(slideCount, SLIDE_W_EMU, SLIDE_H_EMU);
  engine.feedTextRuns = [linkRun(hyperlink)];
  const v = new PptxScrollViewer(container as unknown as HTMLElement, {
    presentation: engine.asPres(),
    gap: 10,
    overscan: 1,
    paddingLeft: 0,
    paddingRight: 0,
    enableTextSelection: true,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  v.relayout();
  // Let the main-mode render resolve so the overlay is built for the top slot.
  await Promise.resolve();
  await Promise.resolve();

  function clickLink(): void {
    // Find any mounted slot's text-layer div and click its first span.
    for (const slot of scrollHost.children) {
      const layer = slot.children.find((k) => k.tag === 'div') as FakeEl | undefined;
      const shapeDiv = layer?.children[0] as FakeEl | undefined;
      const span = shapeDiv?.children[0] as (FakeEl & { click?: () => void }) | undefined;
      if (span && '_listeners' in span) {
        (span as unknown as { dispatch: (t: string, e?: unknown) => void }).dispatch('click', {
          preventDefault() {},
        });
        return;
      }
    }
    throw new Error('no clickable link span found in any mounted slot');
  }

  return { container, engine, v, scrollHost, clickLink };
}

describe('PptxScrollViewer — internal-hyperlink click default (IX-nav)', () => {
  it('resolves an internal slide-part ref via the engine and scrolls to that slide', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: '../slides/slide3.xml' };
    const { engine, v, scrollHost, clickLink } = await setup(5, target);
    engine.internalTargets.set('../slides/slide3.xml', 2); // slide3 ⇒ index 2

    clickLink();

    expect(engine.resolveCalls).toContainEqual({ ref: '../slides/slide3.xml', current: 0 });
    // The click must land exactly where scrollToSlide(2) lands — decoupled from
    // the padding/clamp convention (which scrollToSlide owns).
    const afterClick = scrollHost.scrollTop;
    v.scrollToSlide(2);
    expect(afterClick).toBeCloseTo(scrollHost.scrollTop, 3);
    expect(afterClick).toBeGreaterThan(0); // it actually moved to a lower slide
    v.destroy();
  });

  it('passes the CURRENT top slide as the relative base for a show-jump verb', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: 'ppaction://hlinkshowjump?jump=nextslide' };
    // nextslide from `current` ⇒ current + 1 (clamped) — model it with a fn.
    const { engine, v, scrollHost, clickLink } = await setup(6, target);
    engine.resolveFn = (_ref, current) => Math.min(current + 1, 5);

    // Scroll down so the top-of-viewport slide is no longer slide 0, then click.
    v.scrollToSlide(3);
    scrollHost.dispatch('scroll');
    clickLink();

    // The dispatch must have asked the engine relative to the CURRENT top slide,
    // not always 0 — that is the whole point of a relative jump. (The exact index
    // is the slide intersecting the viewport top, which is > 0 after scrolling.)
    const lastCall = engine.resolveCalls.at(-1);
    expect(lastCall?.ref).toBe(target.ref);
    expect(lastCall?.current).toBeGreaterThan(0);
    // nextslide(current) ⇒ current+1; the click lands where scrollToSlide does.
    const afterClick = scrollHost.scrollTop;
    v.scrollToSlide((lastCall?.current ?? 0) + 1);
    expect(afterClick).toBeCloseTo(scrollHost.scrollTop, 3);
    v.destroy();
  });

  it('is a safe no-op when the ref resolves to no reachable slide', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: '../slides/nope.xml' };
    const { engine, v, scrollHost, clickLink } = await setup(5, target);
    // internalTargets left empty ⇒ resolveInternalTarget returns undefined.
    const before = scrollHost.scrollTop;
    clickLink();
    expect(engine.resolveCalls.length).toBeGreaterThan(0); // it TRIED to resolve
    expect(scrollHost.scrollTop).toBe(before); // ...but did not move
    v.destroy();
  });

  it('delegates fully to a caller-supplied onHyperlinkClick, with slideIndex populated, and does NOT scroll itself', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: '../slides/slide4.xml' };
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { engine, v, scrollHost, clickLink } = await setup(5, target, { onHyperlinkClick });
    engine.internalTargets.set('../slides/slide4.xml', 3);
    const before = scrollHost.scrollTop;

    clickLink();

    // The callback owns the click: it receives the target ENRICHED with the
    // resolved slideIndex (the field IX1 always left undefined), and the viewer
    // takes NO default navigation.
    expect(onHyperlinkClick).toHaveBeenCalledTimes(1);
    expect(onHyperlinkClick).toHaveBeenCalledWith({ kind: 'internal', ref: '../slides/slide4.xml', slideIndex: 3 });
    expect(scrollHost.scrollTop).toBe(before);
    v.destroy();
  });

  it('still opens an external link in a new tab (external branch unchanged)', async () => {
    const target: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };
    const { v, clickLink } = await setup(3, target);
    // setup()'s installDom stubs a bare window; attach a spy `open` for the
    // external branch (openExternalHyperlink reads the ambient window.open).
    const winOpen = vi.fn();
    (globalThis as unknown as { window: { open: unknown } }).window.open = winOpen;
    clickLink();
    expect(winOpen).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
    v.destroy();
  });
});
