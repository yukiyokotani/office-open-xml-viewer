import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import {
  FakeDocxEngine, installDom, makeBorrowedDocxScrollViewer, makeContainer, makeEl, type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { DocxDocument } from './document';
import type { DocxTextRunInfo } from './renderer';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Text highlights (`highlightText`): a persistent set of term groups, painted
 * in each group's background on the same overlay as the find, with their own
 * matches, cursor and clear.
 */
const RUN: DocxTextRunInfo = {
  text: 'Alpha beta', x: 4, y: 6, w: 60, h: 12, fontSize: 12, font: '12px serif',
};
const GROUPS = [
  { id: 'system', terms: ['alpha'], color: 'rgba(0, 0, 255, 0.3)' },
  { id: 'user', terms: ['beta', 'alpha beta'], color: 'rgba(255, 0, 0, 0.3)' },
];
const RING = '2px solid #111';

function installMeasuringDom(): void {
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
}

function makeViewer(
  engine: FakeDocxEngine, mode: 'main' | 'worker',
): { viewer: DocxViewer; layer: () => FakeEl } {
  const parent = makeEl('div');
  const canvas = makeEl('canvas');
  parent.appendChild(canvas);
  const viewer = new DocxViewer(canvas as unknown as HTMLCanvasElement, {
    mode,
    findHighlightColors: { match: 'rgba(1, 2, 3, 0.4)', active: 'rgba(4, 5, 6, 0.7)' },
  });
  (viewer as unknown as {
    _documentOwner: { install(document: DocxDocument): void };
  })._documentOwner.install(engine.asDoc());
  return {
    viewer,
    layer: () => (viewer as unknown as { _highlightLayer: FakeEl })._highlightLayer,
  };
}

describe.each(['main', 'worker'] as const)('DocxViewer text highlights (%s mode)', (mode) => {
  it('highlights every group on every page, next to the find, and navigates them', async () => {
    installMeasuringDom();
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }], mode);
    engine.feedTextRuns = [RUN];
    const { viewer, layer } = makeViewer(engine, mode);

    const matches = await viewer.highlightText(GROUPS);
    expect(matches.map((match) => [match.matchIndex, match.location.page, match.group, match.term, match.text])).toEqual([
      [0, 0, 'system', 'alpha', 'Alpha'],
      [1, 0, 'user', 'alpha beta', 'Alpha beta'],
      [2, 1, 'system', 'alpha', 'Alpha'],
      [3, 1, 'user', 'alpha beta', 'Alpha beta'],
    ]);
    expect(layer().children.map((box) => box.style.background))
      .toEqual(['rgba(0, 0, 255, 0.3)', 'rgba(255, 0, 0, 0.3)']);

    // The find is drawn on top and keeps its own matches and cursor.
    expect(await viewer.findText('beta')).toHaveLength(2);
    expect(layer().children.map((box) => box.style.background))
      .toEqual(['rgba(0, 0, 255, 0.3)', 'rgba(255, 0, 0, 0.3)', 'rgba(1, 2, 3, 0.4)']);
    expect((await viewer.findNext())?.matchIndex).toBe(0);

    expect(await viewer.highlightPrev()).toMatchObject({ matchIndex: 3, group: 'user', location: { page: 1 } });
    expect(viewer.currentPage).toBe(1);
    expect(layer().children.map((box) => box.style.outline === RING)).toEqual([false, true, false]);
    expect(await viewer.highlightNext()).toMatchObject({ matchIndex: 0, group: 'system', location: { page: 0 } });
    expect(viewer.currentPage).toBe(0);
    expect(layer().children.map((box) => box.style.outline === RING)).toEqual([true, false, false]);
    // The find's active match is still the first one.
    expect(layer().children[2].style.background).toBe('rgba(4, 5, 6, 0.7)');

    viewer.clearFind();
    expect(layer().children).toHaveLength(2);
    expect(await viewer.findNext()).toBeNull();

    await viewer.findText('beta');
    viewer.clearHighlights();
    expect(layer().children.map((box) => box.style.background)).toEqual(['rgba(1, 2, 3, 0.4)']);
    expect(await viewer.highlightNext()).toBeNull();
    viewer.destroy();
  });

  it('replaces the set on each call and clears on an empty set', async () => {
    installMeasuringDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], mode);
    engine.feedTextRuns = [RUN];
    const { viewer, layer } = makeViewer(engine, mode);
    await viewer.highlightText(GROUPS);
    expect((await viewer.highlightText([{ id: 'only', terms: ['beta'] }])).map((match) => match.group)).toEqual(['only']);
    expect(layer().children.map((box) => box.style.background)).toEqual(['rgba(0, 184, 148, 0.35)']);
    // A repeated group id rejects, even with no terms, and clears nothing.
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('duplicate group id');
    expect(layer().children).toHaveLength(1);
    expect(await viewer.highlightText([])).toEqual([]);
    expect(layer().children).toEqual([]);
    viewer.destroy();
  });

  it('drops the highlights on destroy', async () => {
    installMeasuringDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], mode);
    engine.feedTextRuns = [RUN];
    const { viewer } = makeViewer(engine, mode);
    await viewer.highlightText(GROUPS);
    expect(await viewer.highlightNext()).not.toBeNull();
    viewer.destroy();
    expect(await viewer.highlightNext()).toBeNull();
    expect(await viewer.highlightPrev()).toBeNull();
  });
});

describe('DocxViewer text highlights on a progressively laid out document', () => {
  function progressive(): { viewer: DocxViewer; finish: () => void } {
    installMeasuringDom();
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }]);
    engine.feedTextRuns = [RUN];
    const doc = engine.asDoc();
    let complete = false;
    let settle!: () => void;
    const completion = new Promise<void>((resolve) => { settle = resolve; });
    Object.defineProperties(doc, {
      layoutComplete: { configurable: true, get: () => complete },
      waitUntilLayoutComplete: { configurable: true, value: () => completion },
    });
    const parent = makeEl('div');
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    (viewer as unknown as {
      _documentOwner: { install(document: DocxDocument): void };
    })._documentOwner.install(doc);
    return { viewer, finish: () => { complete = true; settle(); } };
  }

  it('waits for the complete layout before searching', async () => {
    const { viewer, finish } = progressive();
    let settled = false;
    const pending = viewer.highlightText(GROUPS).then((matches) => { settled = true; return matches; });
    await Promise.resolve();
    await Promise.resolve();
    expect(settled).toBe(false);
    finish();
    expect(await pending).toHaveLength(4);
    viewer.destroy();
  });

  it('is cancelled by a clear or a newer set during that wait, without touching a find', async () => {
    const { viewer, finish } = progressive();
    const cleared = viewer.highlightText(GROUPS);
    viewer.clearHighlights();
    const superseded = viewer.highlightText(GROUPS);
    const current = viewer.highlightText([{ id: 'only', terms: ['beta'] }]);
    const found = viewer.findText('alpha');
    finish();
    await expect(cleared).resolves.toEqual([]);
    await expect(superseded).resolves.toEqual([]);
    expect((await current).map((match) => match.group)).toEqual(['only', 'only']);
    expect(await found).toHaveLength(2);
    viewer.destroy();
  });
});

describe('DocxViewer text highlights: the set is taken when the call is made', () => {
  it('ignores later changes to the caller\'s objects, and an invalid set does not cancel a waiting request', async () => {
    installMeasuringDom();
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }]);
    engine.feedTextRuns = [RUN];
    engine.setLayoutComplete(false);
    const { viewer } = makeViewer(engine, 'main');
    const terms = ['alpha'];
    const waiting = viewer.highlightText([{ id: 'a', terms }]);
    terms[0] = 'beta';
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }]))
      .rejects.toThrow('duplicate group id');
    engine.setLayoutComplete(true);
    expect((await waiting).map((match) => match.text)).toEqual(['Alpha', 'Alpha']);
    viewer.destroy();
  });
});

describe.each(['main', 'worker'] as const)('DocxScrollViewer text highlights (%s mode)', (mode) => {
  function mount(pages: number) {
    installDom();
    const container = makeContainer(200, 120);
    const engine = new FakeDocxEngine(
      pages, Array.from({ length: pages }, () => ({ widthPt: 100, heightPt: 200 })), mode,
    );
    engine.feedTextRuns = [RUN];
    const viewer = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      findHighlightColors: { match: 'rgba(1, 2, 3, 0.4)', active: 'rgba(4, 5, 6, 0.7)' },
      gap: 0, overscan: 0, paddingTop: 0, paddingBottom: 0, paddingLeft: 0, paddingRight: 0,
    });
    const scrollHost = container.children[0].children[0] as FakeEl;
    scrollHost.clientHeight = 120;
    viewer.relayout();
    const boxes = (): FakeEl[] => scrollHost.children
      .filter((child) => child.children.some((nested) => nested.tag === 'canvas'))
      .map((slot) => slot.children.find((child) => child.tag === 'div') as FakeEl)
      .flatMap((layer) => layer.children);
    return { viewer, scrollHost, boxes, engine };
  }

  it('highlights pages outside the mounted window, navigates, and clears independently', async () => {
    const { viewer, scrollHost, boxes } = mount(4);
    const matches = await viewer.highlightText(GROUPS);
    expect(matches).toHaveLength(8);
    expect(matches.map((match) => match.location.page)).toEqual([0, 0, 1, 1, 2, 2, 3, 3]);
    expect(matches.slice(0, 2).map((match) => match.group)).toEqual(['system', 'user']);
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);
    expect(boxes().some((box) => box.style.background === 'rgba(255, 0, 0, 0.3)')).toBe(true);

    await viewer.findText('beta');
    expect(boxes().some((box) => box.style.background === 'rgba(1, 2, 3, 0.4)')).toBe(true);

    expect(await viewer.highlightPrev()).toMatchObject({ matchIndex: 7, location: { page: 3 } });
    expect(scrollHost.scrollTop).toBeGreaterThan(0);
    expect(boxes().some((box) => box.style.outline === RING)).toBe(true);

    viewer.clearFind();
    expect(boxes().some((box) => box.style.background === 'rgba(1, 2, 3, 0.4)')).toBe(false);
    expect(boxes().some((box) => box.style.background === 'rgba(0, 0, 255, 0.3)')).toBe(true);

    await viewer.findText('beta');
    viewer.clearHighlights();
    expect(boxes().every((box) => box.style.background === 'rgba(1, 2, 3, 0.4)')).toBe(true);
    expect(boxes().length).toBeGreaterThan(0);
    expect(await viewer.highlightNext()).toBeNull();

    viewer.clearFind();
    expect(boxes()).toEqual([]);
    viewer.destroy();
  });

  it('takes the set when the call is made, and an invalid set does not cancel a waiting request', async () => {
    const { viewer, engine } = mount(2);
    engine.setLayoutComplete(false);
    const terms = ['alpha'];
    const waiting = viewer.highlightText([{ id: 'a', terms }]);
    terms[0] = 'beta';
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }]))
      .rejects.toThrow('duplicate group id');
    engine.setLayoutComplete(true);
    expect((await waiting).map((match) => match.text)).toEqual(['Alpha', 'Alpha']);
    viewer.destroy();
  });

  it('keeps highlight geometry current after a zoom settle render without a find', async () => {
    vi.useFakeTimers();
    try {
      const { viewer, boxes, engine } = mount(1);
      await viewer.highlightText([{ id: 'only', terms: ['alpha'] }]);
      await Promise.resolve();
      expect(boxes()[0].style.left).toBe('2%');

      engine.feedTextRuns = [{ ...RUN, x: 40 }];
      viewer.setScale(viewer.getScale() * 2);
      vi.advanceTimersByTime(200);
      await Promise.resolve();
      await Promise.resolve();
      expect(boxes()[0].style.left).toBe('10%');
      viewer.destroy();
    } finally {
      vi.useRealTimers();
    }
  });

  it('drops the highlights on destroy', async () => {
    const { viewer, boxes } = mount(2);
    await viewer.highlightText(GROUPS);
    const shown = boxes().length;
    expect(shown).toBeGreaterThan(0);
    await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('duplicate group id');
    expect(boxes()).toHaveLength(shown);
    viewer.destroy();
    expect(await viewer.highlightNext()).toBeNull();
  });
});
