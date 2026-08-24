import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
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

async function setupScroll(
  opts: Record<string, unknown> = {},
  mode: 'main' | 'worker' = 'main',
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
    v.setShowTrackedChanges(true);
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
