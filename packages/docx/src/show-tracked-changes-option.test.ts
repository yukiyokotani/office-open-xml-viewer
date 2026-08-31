import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { subscribeDocxLayoutView } from './document-layout-view.js';
import {
  installDom,
  makeEl,
  makeContainer,
  makeBorrowedDocxScrollViewer,
  FakeDocxEngine,
  type FakeEl,
} from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * ECMA-376 §17.13.5 `showTrackedChanges` — a viewer-level view switch (default
 * `false` = final view). The single gate is the render-path threading: every
 * render / run-collection call passes the live flag, which selects the cached
 * layout variant (final vs markup) down in the engine. These tests pin that
 * threading on the single-canvas viewer (default / true / runtime toggle) and
 * on BOTH scroll-viewer render paths (main `renderPage` and worker
 * `renderPageToBitmap` — a distinct call site).
 */

const PAGE = [{ widthPt: 595, heightPt: 842 }];

async function mountViewer(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakeDocxEngine(3, PAGE);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.docx');
  return { v, engine };
}

describe('DocxViewer — showTrackedChanges option', () => {
  it('renders the final view by default (no flag on the render call)', async () => {
    const { v, engine } = await mountViewer();
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    expect(engine.renderCalls[0]!.showTrackedChanges).toBeUndefined();
    v.destroy();
  });

  it('threads showTrackedChanges: true into the render call', async () => {
    const { v, engine } = await mountViewer({ showTrackedChanges: true });
    expect(engine.renderCalls[0]!.showTrackedChanges).toBe(true);
    v.destroy();
  });

  it('an explicit false matches the default final view', async () => {
    const { v, engine } = await mountViewer({ showTrackedChanges: false });
    expect(engine.renderCalls[0]!.showTrackedChanges).toBe(false);
    v.destroy();
  });

  it('setShowTrackedChanges re-renders at the new variant and back', async () => {
    const { v, engine } = await mountViewer();
    const before = engine.renderCalls.length;
    await v.setShowTrackedChanges(true);
    expect(engine.renderCalls.length).toBe(before + 1);
    expect(engine.renderCalls.at(-1)!.showTrackedChanges).toBe(true);
    await v.setShowTrackedChanges(false);
    expect(engine.renderCalls.at(-1)!.showTrackedChanges).toBe(false);
    // A no-op set (same value) does not re-render.
    const count = engine.renderCalls.length;
    await v.setShowTrackedChanges(false);
    expect(engine.renderCalls.length).toBe(count);
    v.destroy();
  });
});

describe.each(['main', 'worker'] as const)(
  'DocxViewer — shared borrowed document (%s mode)',
  (mode) => {
    it('keeps geometry and paint on the document\'s post-construction active view', async () => {
      installDom();
      const engine = new FakeDocxEngine(2, PAGE, mode);
      engine.layoutView = { showTrackedChanges: false, currentDate: 10 };
      const viewers = [makeEl('canvas'), makeEl('canvas')].map((canvas) =>
        DocxViewer.fromDocument(canvas as unknown as HTMLCanvasElement, engine.asDoc()));
      await Promise.all(viewers.map((viewer) => viewer.goToPage(0)));

      const calls = mode === 'worker' ? engine.bitmapCalls : engine.renderCalls;
      const beforeExternalSwitch = calls.length;
      await engine.setLayoutView({ showTrackedChanges: true, currentDate: 20 });
      await Promise.resolve();

      expect(calls.length).toBeGreaterThan(beforeExternalSwitch);
      for (const call of calls.slice(beforeExternalSwitch)) {
        expect(call.showTrackedChanges).toBe(true);
        expect(call.currentDate).toBe(20);
      }

      const beforeViewerSwitch = calls.length;
      await viewers[1]!.setShowTrackedChanges(false);
      await Promise.resolve();

      expect(calls.length).toBeGreaterThan(beforeViewerSwitch);
      for (const call of calls.slice(beforeViewerSwitch)) {
        expect(call.showTrackedChanges ?? false).toBe(false);
        expect(call.currentDate).toBe(20);
      }
      for (const viewer of viewers) viewer.destroy();
    });
  },
);

describe('DocxViewer — layout-view publication ordering', () => {
  it('rejects an older outer publication after a listener installs a newer view', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    engine.layoutView = { showTrackedChanges: false, currentDate: 0 };
    const unsubscribe = subscribeDocxLayoutView(
      engine.asDoc(),
      (publication) => {
        if (publication.view.currentDate === 10) {
          engine.setLayoutView({ showTrackedChanges: true, currentDate: 20 });
        }
      },
      vi.fn(),
    );
    const viewers = [makeEl('canvas'), makeEl('canvas')].map((canvas) =>
      DocxViewer.fromDocument(canvas as unknown as HTMLCanvasElement, engine.asDoc()));
    await Promise.all(viewers.map((viewer) => viewer.goToPage(0)));

    const before = engine.renderCalls.length;
    engine.setLayoutView({ showTrackedChanges: true, currentDate: 10 });
    await Promise.resolve();

    expect(engine.layoutView).toEqual({ showTrackedChanges: true, currentDate: 20 });
    expect(engine.renderCalls.length).toBeGreaterThan(before);
    for (const call of engine.renderCalls.slice(before)) {
      expect(call.showTrackedChanges).toBe(true);
      expect(call.currentDate).toBe(20);
    }

    unsubscribe();
    for (const viewer of viewers) viewer.destroy();
  });

  it('awaits its owned repaint and preserves repaint rejection', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE, 'main', true);
    const onError = vi.fn();
    const viewer = DocxViewer.fromDocument(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      engine.asDoc(),
      { onError },
    );
    const initial = viewer.goToPage(0);
    engine.renderCalls.at(-1)!.resolve();
    await initial;

    let settled = false;
    const beforeSwitch = engine.renderCalls.length;
    const switching = viewer.setShowTrackedChanges(true).then(() => { settled = true; });
    await vi.waitFor(() => expect(engine.renderCalls.length).toBe(beforeSwitch + 1));
    expect(settled).toBe(false);
    engine.renderCalls.at(-1)!.resolve();
    await switching;
    expect(settled).toBe(true);

    const beforeFailure = engine.renderCalls.length;
    const failing = viewer.setShowTrackedChanges(false);
    await vi.waitFor(() => expect(engine.renderCalls.length).toBe(beforeFailure + 1));
    engine.renderCalls.at(-1)!.reject(new Error('repaint failed'));
    await expect(failing).rejects.toThrow('repaint failed');
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });
});

async function setupScroll(
  opts: Record<string, unknown> = {},
  mode: 'main' | 'worker' = 'main',
  layoutView?: Readonly<{
    showTrackedChanges: boolean;
    currentDate: number;
  }>,
) {
  installDom();
  const container = makeContainer(200, 400);
  const engine = new FakeDocxEngine(
    3,
    [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 200 },
    ],
    mode,
  );
  engine.layoutView = layoutView ?? {
    showTrackedChanges: opts.showTrackedChanges === true,
    currentDate: typeof opts.currentDate === 'number' ? opts.currentDate : 0,
  };
  const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
    document: engine.asDoc(),
    gap: 10,
    overscan: 1,
    paddingLeft: 0,
    paddingRight: 0,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  v.relayout();
  await Promise.resolve();
  await Promise.resolve();
  await new Promise((r) => setTimeout(r, 0));
  return { v, engine, scrollHost };
}

describe('DocxScrollViewer — showTrackedChanges option (main mode)', () => {
  it('inherits an already-loaded document\'s active markup view', async () => {
    const currentDate = 1_700_000_000_000;
    const { v, engine } = await setupScroll(
      {},
      'main',
      { showTrackedChanges: true, currentDate },
    );
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    for (const call of engine.renderCalls) {
      expect(call.showTrackedChanges).toBe(true);
      expect(call.currentDate).toBe(currentDate);
    }

    const before = engine.renderCalls.length;
    await v.setShowTrackedChanges(false);
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));

    // The first OFF is a real transition, not a stale-state no-op, and it keeps
    // the borrowed document's other layout axis intact.
    expect(engine.renderCalls.length).toBeGreaterThan(before);
    expect(engine.layoutViews.at(-1)).toEqual({
      showTrackedChanges: false,
      currentDate,
    });
    for (const call of engine.renderCalls.slice(before)) {
      expect(call.showTrackedChanges ?? false).toBe(false);
      expect(call.currentDate).toBe(currentDate);
    }
    v.destroy();
  });

  it('renders the final view by default', async () => {
    const { v, engine } = await setupScroll();
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    for (const call of engine.renderCalls) {
      // The scroll viewer emits the flag only for the markup view, so default
      // render calls keep their historical option shape (no key at all).
      expect(call.showTrackedChanges ?? false).toBe(false);
    }
    v.destroy();
  });

  it('threads showTrackedChanges: true into every slot render', async () => {
    const { v, engine } = await setupScroll({ showTrackedChanges: true });
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    for (const call of engine.renderCalls) {
      expect(call.showTrackedChanges).toBe(true);
    }
    v.destroy();
  });

  it('setShowTrackedChanges re-renders every mounted slot at the new variant', async () => {
    const { v, engine } = await setupScroll();
    const before = engine.renderCalls.length;
    await v.setShowTrackedChanges(true);
    await Promise.resolve();
    await new Promise((r) => setTimeout(r, 0));
    expect(engine.renderCalls.length).toBeGreaterThan(before);
    for (const call of engine.renderCalls.slice(before)) {
      expect(call.showTrackedChanges).toBe(true);
    }
    v.destroy();
  });
});

describe('DocxScrollViewer — showTrackedChanges in worker mode', () => {
  it('threads the flag through renderPageToBitmap (the worker call site)', async () => {
    const { v, engine } = await setupScroll({ showTrackedChanges: true }, 'worker');
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    expect(engine.renderCalls.length).toBe(0);
    for (const call of engine.bitmapCalls) {
      expect(call.showTrackedChanges).toBe(true);
    }
    v.destroy();
  });

  it('worker mode default stays the final view', async () => {
    const { v, engine } = await setupScroll({}, 'worker');
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    for (const call of engine.bitmapCalls) {
      expect(call.showTrackedChanges ?? false).toBe(false);
    }
    v.destroy();
  });
});

describe.each(['main', 'worker'] as const)(
  'DocxScrollViewer — shared borrowed document (%s mode)',
  (mode) => {
    it('keeps every viewer on the document\'s post-construction active view', async () => {
      installDom();
      const engine = new FakeDocxEngine(
        2,
        [{ widthPt: 100, heightPt: 200 }],
        mode,
      );
      engine.layoutView = { showTrackedChanges: false, currentDate: 10 };
      const viewers = [makeContainer(200, 400), makeContainer(200, 400)].map((container) => {
        const viewer = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
          document: engine.asDoc(),
          gap: 10,
          overscan: 1,
          paddingLeft: 0,
          paddingRight: 0,
        });
        const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
        scrollHost.clientHeight = 400;
        scrollHost.clientWidth = 200;
        viewer.relayout();
        return viewer;
      });
      await Promise.resolve();
      await new Promise((resolve) => setTimeout(resolve, 0));

      const calls = mode === 'worker' ? engine.bitmapCalls : engine.renderCalls;
      const beforeExternalSwitch = calls.length;
      await engine.setLayoutView({ showTrackedChanges: true, currentDate: 20 });
      await Promise.resolve();
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(calls.length).toBeGreaterThan(beforeExternalSwitch);
      for (const call of calls.slice(beforeExternalSwitch)) {
        expect(call.showTrackedChanges).toBe(true);
        expect(call.currentDate).toBe(20);
      }

      const beforeViewerSwitch = calls.length;
      await viewers[1]!.setShowTrackedChanges(false);
      await Promise.resolve();
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(calls.length).toBeGreaterThan(beforeViewerSwitch);
      for (const call of calls.slice(beforeViewerSwitch)) {
        expect(call.showTrackedChanges ?? false).toBe(false);
        expect(call.currentDate).toBe(20);
      }
      for (const viewer of viewers) viewer.destroy();
    });
  },
);
