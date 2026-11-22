import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9144000;
const SLIDE_H_EMU = 6858000;

/**
 * IX-nav wiring — PptxViewer's internal-hyperlink click default.
 *
 * The single-canvas viewer's dispatch is `_onHyperlinkClick`, exercised here via
 * the same fake-engine seam the render-error tests use (spy `PptxPresentation.load`
 * to install a FakePptxEngine as the viewer's engine). The overlay is DOM-bound,
 * so the dispatch is invoked directly — it is the branch selection (resolve →
 * goToSlide / delegate / no-op / external) that IX-nav adds and that these pin.
 */
async function mountLoaded(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakePptxEngine(6, SLIDE_W_EMU, SLIDE_H_EMU);
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.pptx');
  // Reach the private dispatch (DOM overlay would call it on a span click).
  const dispatch = (t: HyperlinkTarget): void =>
    (v as unknown as { _onHyperlinkClick: (t: HyperlinkTarget) => void })._onHyperlinkClick(t);
  return { v, engine, dispatch };
}

/** The viewer's navigation entry shared by goToSlide() and link clicks. */
function spyNavigation(v: PptxViewer) {
  return vi.spyOn(v as unknown as { _goToSlide(index: number, generation: number, backward?: boolean): Promise<void> }, '_goToSlide');
}

describe('PptxViewer — internal-hyperlink click default (IX-nav)', () => {
  it('resolves an internal ref via the engine and navigates to that slide', async () => {
    const { v, engine, dispatch } = await mountLoaded();
    engine.internalTargets.set('../slides/slide5.xml', 4);
    const goTo = spyNavigation(v);

    dispatch({ kind: 'internal', ref: '../slides/slide5.xml' });

    expect(engine.resolveCalls.at(-1)).toEqual({ ref: '../slides/slide5.xml', current: 0 });
    expect(goTo.mock.calls.at(-1)?.[0]).toBe(4);
    await vi.waitFor(() => expect(v.slideIndex).toBe(4));
    v.destroy();
  });

  it('resolves a relative jump against the CURRENT slide', async () => {
    const { v, engine, dispatch } = await mountLoaded();
    engine.resolveFn = (_ref, current) => Math.min(current + 1, 5);
    await v.goToSlide(2); // current becomes 2
    const goTo = spyNavigation(v);

    dispatch({ kind: 'internal', ref: 'ppaction://hlinkshowjump?jump=nextslide' });

    expect(engine.resolveCalls.at(-1)?.current).toBe(2); // relative to the current slide
    expect(goTo.mock.calls.at(-1)?.[0]).toBe(3); // nextslide(2) ⇒ 3
    v.destroy();
  });

  it('is a safe no-op when the ref resolves to no reachable slide', async () => {
    const { v, engine, dispatch } = await mountLoaded();
    const goTo = spyNavigation(v);
    dispatch({ kind: 'internal', ref: '../slides/nope.xml' });
    expect(engine.resolveCalls.length).toBeGreaterThan(0); // tried
    expect(goTo).not.toHaveBeenCalled(); // but did not navigate
    v.destroy();
  });

  it('delegates to a caller-supplied onHyperlinkClick with slideIndex populated, and does NOT navigate', async () => {
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { v, engine, dispatch } = await mountLoaded({ onHyperlinkClick });
    engine.internalTargets.set('../slides/slide3.xml', 2);
    const goTo = spyNavigation(v);

    dispatch({ kind: 'internal', ref: '../slides/slide3.xml' });

    // The callback owns the click and receives the ENRICHED target (slideIndex
    // populated — the field IX1 always left undefined); the viewer does not nav.
    expect(onHyperlinkClick).toHaveBeenCalledWith({ kind: 'internal', ref: '../slides/slide3.xml', slideIndex: 2 });
    expect(goTo).not.toHaveBeenCalled();
    v.destroy();
  });

  it('opens an external link in a new tab', async () => {
    const { v, dispatch } = await mountLoaded();
    const winOpen = vi.fn();
    (globalThis as unknown as { window: { open: unknown } }).window.open = winOpen;
    dispatch({ kind: 'external', url: 'https://example.com/' });
    expect(winOpen).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
    v.destroy();
  });
});
