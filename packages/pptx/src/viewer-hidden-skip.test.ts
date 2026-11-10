import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxViewer, type HiddenSlideMode } from './viewer.js';
import { FakePptxEngine, installDom, makeEl } from './scroll-viewer-test-dom.js';
import type { PptxTextRunInfo } from './renderer';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const WIDTH = 9144000;
const HEIGHT = 6858000;

function run(text: string): PptxTextRunInfo {
  return {
    text,
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
}

/** installDom + <canvas> elements with a measuring 2d context (find needs it). */
function mount(): unknown {
  installDom();
  vi.stubGlobal('document', {
    createElement: (t: string) => {
      const el = makeEl(t);
      if (t === 'canvas') {
        el.getContext = (kind: string) =>
          kind === '2d'
            ? { font: '', measureText: (s: string) => ({ width: s.length * 7 }) }
            : {};
      }
      return el;
    },
  });
  const parent = makeEl('div');
  const canvas = makeEl('canvas');
  parent.appendChild(canvas);
  return canvas;
}

function viewerFor(hidden: number[], count: number, mode: HiddenSlideMode = 'skip') {
  const engine = new FakePptxEngine(count, WIDTH, HEIGHT);
  for (const i of hidden) engine.hiddenSlides.add(i);
  engine.feedTextRuns = [run('needle')];
  const renderSlide = vi.spyOn(engine, 'renderSlide');
  const viewer = PptxViewer.fromPresentation(
    mount() as HTMLCanvasElement,
    engine.asPres(),
    { hiddenSlideMode: mode },
  ) as PptxViewer;
  const drawn = () => renderSlide.mock.calls.map((call) => call[1]);
  return { viewer, drawn };
}

describe("PptxViewer hiddenSlideMode: 'skip' (#2181)", () => {
  it('goToSlide on a hidden slide lands on the next shown slide, else the previous one', async () => {
    // Slides 1 and 2 hidden; 3 is the next shown slide after them.
    const { viewer, drawn } = viewerFor([1, 2, 4], 5);
    await viewer.goToSlide(1);
    expect(viewer.slideIndex).toBe(3);
    // No shown slide after 4 → the nearest shown slide before it.
    await viewer.goToSlide(4);
    expect(viewer.slideIndex).toBe(3);
    // Clamped out-of-range index follows the same rule.
    await viewer.goToSlide(99);
    expect(viewer.slideIndex).toBe(3);
    await viewer.goToSlide(0);
    expect(viewer.slideIndex).toBe(0);
    expect(drawn().some((slide) => [1, 2, 4].includes(slide))).toBe(false);
    viewer.destroy();
  });

  it('an all-hidden deck has nothing to skip to, so the slide is drawn as in show mode', async () => {
    const { viewer } = viewerFor([0, 1], 2);
    await viewer.goToSlide(1);
    expect(viewer.slideIndex).toBe(1);
    viewer.destroy();
  });

  it('show and dim modes still honour goToSlide on a hidden slide', async () => {
    for (const mode of ['show', 'dim'] as const) {
      const { viewer } = viewerFor([1], 3, mode);
      await viewer.goToSlide(1);
      expect(viewer.slideIndex).toBe(1);
      viewer.destroy();
    }
  });

  it('findText excludes hidden slides and findNext/findPrev never land on one', async () => {
    const { viewer, drawn } = viewerFor([1], 3);
    const matches = await viewer.findText('needle');
    expect(matches.map((match) => match.location.slide)).toEqual([0, 2]);
    expect((await viewer.findNext())?.location.slide).toBe(0);
    expect((await viewer.findNext())?.location.slide).toBe(2);
    expect(viewer.slideIndex).toBe(2);
    expect((await viewer.findPrev())?.location.slide).toBe(0);
    expect(drawn()).not.toContain(1);

    // Leaving 'skip' clears the find; a new search includes the hidden slide.
    await viewer.setHiddenSlideMode('show');
    expect(await viewer.findNext()).toBeNull();
    const all = await viewer.findText('needle');
    expect(all.map((match) => match.location.slide)).toEqual([0, 1, 2]);
    viewer.destroy();
  });

  it('an internal hyperlink to a hidden slide follows the goToSlide rule', async () => {
    const { viewer, drawn } = viewerFor([1], 3);
    (viewer as unknown as {
      _onHyperlinkClick(target: { kind: 'internal'; ref: string; slideIndex: number }): void;
    })._onHyperlinkClick({ kind: 'internal', ref: '../slides/slide2.xml', slideIndex: 1 });
    await vi.waitFor(() => expect(viewer.slideIndex).toBe(2));
    expect(drawn()).not.toContain(1);
    viewer.destroy();
  });
});
