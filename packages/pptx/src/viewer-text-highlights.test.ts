import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import {
  FakePptxEngine, installDom, makeBorrowedPptxScrollViewer, makeContainer, makeEl, type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { PptxTextRunInfo } from './renderer';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Text highlights (`highlightText`): a persistent set of term groups, painted
 * in each group's background on the same overlay as the find, with their own
 * matches, cursor and clear.
 */
const WIDTH = 9144000;
const HEIGHT = 6858000;
const RUN: PptxTextRunInfo = {
  text: 'Alpha beta', inShapeX: 4, inShapeY: 6, w: 60, h: 12, fontSize: 12, font: '12px serif',
  shapeX: 0, shapeY: 0, shapeW: 100, shapeH: 40, rotation: 0,
};
const GROUPS = [
  { id: 'system', terms: ['alpha'], color: 'rgba(0, 0, 255, 0.3)' },
  { id: 'user', terms: ['beta', 'alpha beta'], color: 'rgba(255, 0, 0, 0.3)' },
];
const FIND_COLORS = { match: 'rgba(1, 2, 3, 0.4)', active: 'rgba(4, 5, 6, 0.7)' };
const RING = '2px solid #111';

function mountCanvas(): HTMLCanvasElement {
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
  return canvas as unknown as HTMLCanvasElement;
}

function makeViewer(engine: FakePptxEngine, options: ConstructorParameters<typeof PptxViewer>[1] = {}) {
  engine.feedTextRuns = [RUN];
  const viewer = PptxViewer.fromPresentation(mountCanvas(), engine.asPres(), {
    findHighlightColors: FIND_COLORS, ...options,
  }) as PptxViewer;
  const boxes = (): FakeEl[] =>
    (viewer as unknown as { highlightLayer: FakeEl }).highlightLayer.children.flatMap((shape) => shape.children);
  return { viewer, boxes };
}

describe.each(['main', 'worker'] as const)('PptxViewer text highlights (%s mode)', (mode) => {
  it('highlights every group on every slide, next to the find, and navigates them', async () => {
    const { viewer, boxes } = makeViewer(new FakePptxEngine(2, WIDTH, HEIGHT, mode), { mode });
    await viewer.goToSlide(0);

    const matches = await viewer.highlightText(GROUPS);
    expect(matches.map((match) => [match.matchIndex, match.location.slide, match.group, match.term, match.text])).toEqual([
      [0, 0, 'system', 'alpha', 'Alpha'],
      [1, 0, 'user', 'alpha beta', 'Alpha beta'],
      [2, 1, 'system', 'alpha', 'Alpha'],
      [3, 1, 'user', 'alpha beta', 'Alpha beta'],
    ]);
    expect(boxes().map((box) => box.style.background)).toEqual(['rgba(0, 0, 255, 0.3)', 'rgba(255, 0, 0, 0.3)']);

    expect(await viewer.findText('beta')).toHaveLength(2);
    expect(boxes().map((box) => box.style.background))
      .toEqual(['rgba(0, 0, 255, 0.3)', 'rgba(255, 0, 0, 0.3)', FIND_COLORS.match]);
    expect((await viewer.findNext())?.matchIndex).toBe(0);

    expect(await viewer.highlightPrev()).toMatchObject({ matchIndex: 3, group: 'user', location: { slide: 1 } });
    expect(viewer.slideIndex).toBe(1);
    expect(boxes().map((box) => box.style.outline === RING)).toEqual([false, true, false]);
    expect(await viewer.highlightNext()).toMatchObject({ matchIndex: 0, group: 'system', location: { slide: 0 } });
    expect(viewer.slideIndex).toBe(0);
    expect(boxes().map((box) => box.style.outline === RING)).toEqual([true, false, false]);
    expect(boxes()[2].style.background).toBe(FIND_COLORS.active);

    viewer.clearFind();
    expect(boxes()).toHaveLength(2);
    expect(await viewer.findNext()).toBeNull();

    await viewer.findText('beta');
    viewer.clearHighlights();
    expect(boxes().map((box) => box.style.background)).toEqual([FIND_COLORS.match]);
    expect(await viewer.highlightNext()).toBeNull();

    // An empty find query leaves a later highlight set in place.
    await viewer.highlightText(GROUPS);
    await viewer.findText('');
    expect(boxes()).toHaveLength(2);
    // A repeated group id rejects, even with no terms, and clears nothing.
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('duplicate group id');
    expect(boxes()).toHaveLength(2);
    expect(await viewer.highlightText([])).toEqual([]);
    expect(boxes()).toEqual([]);
    viewer.destroy();
    expect(await viewer.highlightNext()).toBeNull();
  });
});

describe("PptxViewer text highlights and hiddenSlideMode: 'skip'", () => {
  it('leaves hidden slides out, and drops the highlights when the mode changes', async () => {
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.hiddenSlides.add(1);
    const { viewer } = makeViewer(engine, { hiddenSlideMode: 'skip' });
    const matches = await viewer.highlightText([{ id: 'only', terms: ['alpha'] }]);
    expect(matches.map((match) => match.location.slide)).toEqual([0, 2]);
    expect((await viewer.highlightPrev())?.location.slide).toBe(2);

    await viewer.setHiddenSlideMode('show');
    expect(await viewer.highlightNext()).toBeNull();
    expect((await viewer.highlightText([{ id: 'only', terms: ['alpha'] }])).map((match) => match.location.slide))
      .toEqual([0, 1, 2]);
    viewer.destroy();
  });
});

describe('PptxViewer text highlights on a progressively prepared deck', () => {
  it('waits for the complete deck, and a clear during the wait cancels the request', async () => {
    const engine = new FakePptxEngine(2, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const { viewer } = makeViewer(engine);
    const cleared = viewer.highlightText(GROUPS);
    viewer.clearHighlights();
    let settled = false;
    // The set is taken when the call is made: later changes to the caller's
    // objects are ignored, and an invalid set does not cancel the request.
    const groups = GROUPS.map((group) => ({ ...group, terms: [...group.terms] }));
    const pending = viewer.highlightText(groups).then((matches) => { settled = true; return matches; });
    groups[0].terms[0] = 'nothing';
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }]))
      .rejects.toThrow('duplicate group id');
    const found = viewer.findText('alpha');
    await Promise.resolve();
    await Promise.resolve();
    expect(settled).toBe(false);
    engine.setLayoutProgress(2, true);
    await expect(cleared).resolves.toEqual([]);
    expect(await pending).toHaveLength(4);
    expect(await found).toHaveLength(2);
    viewer.destroy();
  });
});

describe.each(['main', 'worker'] as const)('PptxScrollViewer text highlights (%s mode)', (mode) => {
  function mount(slides: number, hidden: number[] = [], hiddenSlideMode?: 'skip') {
    installDom();
    const container = makeContainer(200, 80);
    const engine = new FakePptxEngine(slides, WIDTH, HEIGHT, mode);
    for (const slide of hidden) engine.hiddenSlides.add(slide);
    engine.feedTextRuns = [RUN];
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      findHighlightColors: FIND_COLORS,
      ...(hiddenSlideMode ? { hiddenSlideMode } : {}),
      gap: 0, overscan: 0, paddingTop: 0, paddingBottom: 0, paddingLeft: 0, paddingRight: 0,
    });
    const scrollHost = container.children[0].children[0] as FakeEl;
    scrollHost.clientHeight = 80;
    viewer.relayout();
    const boxes = (): FakeEl[] => scrollHost.children
      .filter((child) => child.children.some((nested) => nested.tag === 'canvas'))
      .map((slot) => slot.children.find((child) => child.tag === 'div') as FakeEl)
      .flatMap((layer) => layer.children)
      .flatMap((shape) => shape.children);
    return { viewer, scrollHost, boxes, engine };
  }

  it('highlights slides outside the mounted window, navigates, and clears independently', async () => {
    const { viewer, scrollHost, boxes } = mount(4);
    const matches = await viewer.highlightText(GROUPS);
    expect(matches.map((match) => match.location.slide)).toEqual([0, 0, 1, 1, 2, 2, 3, 3]);
    expect(matches.slice(0, 2).map((match) => match.group)).toEqual(['system', 'user']);
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);
    expect(boxes().some((box) => box.style.background === 'rgba(255, 0, 0, 0.3)')).toBe(true);

    await viewer.findText('beta');
    expect(boxes().some((box) => box.style.background === FIND_COLORS.match)).toBe(true);

    expect(await viewer.highlightPrev()).toMatchObject({ matchIndex: 7, location: { slide: 3 } });
    expect(scrollHost.scrollTop).toBeGreaterThan(0);
    expect(boxes().some((box) => box.style.outline === RING)).toBe(true);

    viewer.clearFind();
    expect(boxes().some((box) => box.style.background === FIND_COLORS.match)).toBe(false);
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);

    await viewer.findText('beta');
    viewer.clearHighlights();
    expect(boxes().length).toBeGreaterThan(0);
    expect(boxes().every((box) => box.style.background === FIND_COLORS.match)).toBe(true);
    expect(await viewer.highlightNext()).toBeNull();

    // An empty find query leaves a later highlight set in place.
    await viewer.highlightText(GROUPS);
    await viewer.findText('');
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('duplicate group id');
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);
    viewer.destroy();
    expect(await viewer.highlightNext()).toBeNull();
  });

  it("leaves slides hidden by 'skip' out, and drops the highlights when the mode changes", async () => {
    const { viewer, boxes } = mount(4, [1], 'skip');
    const matches = await viewer.highlightText([{ id: 'only', terms: ['alpha'] }]);
    expect(matches.map((match) => match.location.slide)).toEqual([0, 2, 3]);
    await viewer.setHiddenSlideMode('show');
    expect(await viewer.highlightNext()).toBeNull();
    expect(boxes()).toEqual([]);
    viewer.destroy();
  });

  it('keeps highlight geometry current after a zoom settle render without a find', async () => {
    vi.useFakeTimers();
    try {
      const { viewer, boxes, engine } = mount(1);
      await viewer.highlightText([{ id: 'only', terms: ['alpha'] }]);
      await Promise.resolve();
      const before = boxes()[0].style.left;
      engine.feedTextRuns = [{ ...RUN, inShapeX: 40 }];
      viewer.setScale(viewer.getScale() * 2);
      vi.advanceTimersByTime(200);
      await Promise.resolve();
      await Promise.resolve();
      expect(boxes()[0].style.left).not.toBe(before);
      viewer.destroy();
    } finally {
      vi.useRealTimers();
    }
  });
});
