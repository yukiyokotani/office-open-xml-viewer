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
import type { DocxTextRunInfo } from './renderer';
import type { DocComment } from './types';
import type { CommentAnchorRange } from './comment-margin-layout';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * ECMA-376 §17.13.4 `showComments` — the comment margin (default `false`). Off
 * means ZERO comment DOM and no gutter reservation: comments stay data-only.
 * On means the wrapper reserves a right gutter and the render path builds the
 * tint + balloon layers from the engine's threads/anchor ranges. These tests
 * pin the option on the single-canvas viewer (default / true / runtime
 * toggle) and on the scroll viewer in BOTH render modes.
 */

const PAGE = [{ widthPt: 595, heightPt: 842 }];

function commentedRun(): DocxTextRunInfo {
  return {
    text: 'annotated',
    x: 40,
    y: 60,
    w: 90,
    h: 14,
    fontSize: 12,
    font: '12px serif',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    sourceRunIndex: 1,
  };
}

const COMMENTS: DocComment[] = [
  { id: '1', author: 'Alice', date: '2024-03-01', text: 'A remark', paragraphs: ['A remark'] },
  { id: '2', author: 'Bob', text: 'A reply', parentId: '1', paragraphs: ['A reply'] },
  { id: '3', author: 'Carol', text: 'Done already', resolved: true, paragraphs: ['Done already'] },
];

const RANGES: CommentAnchorRange[] = [
  { commentId: '1', paragraphPath: [0], startRunIndex: 1, endRunIndex: 2 },
  { commentId: '3', paragraphPath: [0], startRunIndex: 1, endRunIndex: 2 },
];

function seedComments(engine: FakeDocxEngine): void {
  engine.comments = COMMENTS;
  engine.feedCommentAnchorRanges = RANGES;
  engine.feedTextRuns = [commentedRun()];
}

interface ViewerInternals {
  _wrapper: FakeEl;
  _commentTintLayer: FakeEl | null;
  _commentGutterLayer: FakeEl | null;
  _selectedCommentId: string | null;
}

async function mountViewer(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakeDocxEngine(3, PAGE);
  seedComments(engine);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, {
    // A concrete width so the fake engine stamps a canvas CSS box (the
    // overlay's percent denominators).
    width: 595,
    ...opts,
  });
  await v.load('x.docx');
  const internals = v as unknown as ViewerInternals;
  return { v, engine, internals };
}

/** The balloon elements in a gutter layer: children with a click listener. */
function balloons(gutter: FakeEl | null): FakeEl[] {
  return (gutter?.children ?? []).filter((child) => child._listeners.has('click'));
}

describe('DocxViewer — showComments option', () => {
  it('builds no comment DOM and reserves no gutter by default', async () => {
    const { v, internals } = await mountViewer();
    expect(internals._commentTintLayer).toBeNull();
    expect(internals._commentGutterLayer).toBeNull();
    expect(internals._wrapper.style.marginRight).toBe('');
    v.destroy();
  });

  it('reserves the gutter and builds tint + balloons when true', async () => {
    const { v, internals } = await mountViewer({ showComments: true });
    expect(internals._wrapper.style.marginRight).toBe('260px');
    // Tint layer: at least one range box over the commented run.
    expect(internals._commentTintLayer!.children.length).toBeGreaterThan(0);
    // Gutter: one balloon per UNRESOLVED root thread (thread 1 with its
    // reply; the resolved thread 3 is hidden).
    const balloonEls = balloons(internals._commentGutterLayer);
    expect(balloonEls).toHaveLength(1);
    const balloonText = JSON.stringify(
      balloonEls[0]!.children.map((child) => child.textContent),
    );
    expect(balloonText).toContain('Alice');
    expect(balloonText).toContain('A remark');
    expect(balloonText).toContain('Bob');
    expect(balloonText).toContain('A reply');
    expect(balloonText).not.toContain('Carol');
    v.destroy();
  });

  it('honours a custom commentsGutterWidth', async () => {
    const { v, internals } = await mountViewer({ showComments: true, commentsGutterWidth: 180 });
    expect(internals._wrapper.style.marginRight).toBe('180px');
    expect(internals._commentGutterLayer!.style.width).toBe('180px');
    v.destroy();
  });

  it('clicking a balloon selects its thread; clicking again deselects', async () => {
    const { v, internals } = await mountViewer({ showComments: true });
    const first = balloons(internals._commentGutterLayer)[0]!;
    first.dispatch('click', { preventDefault() {} });
    expect(internals._selectedCommentId).toBe('1');
    // The rebuild replaced the balloon; the new one deselects on click.
    const reselected = balloons(internals._commentGutterLayer)[0]!;
    reselected.dispatch('click', { preventDefault() {} });
    expect(internals._selectedCommentId).toBeNull();
    v.destroy();
  });

  it('clicking the tinted range selects the thread; clicking it again deselects', async () => {
    const { v, internals } = await mountViewer({ showComments: true });
    // The wrapper carries ONE shared hit-test listener; give the tint layer a
    // laid-out box (the fake's getBoundingClientRect reads client sizes) that
    // matches the page CSS box, so hit coordinates map 1:1.
    const tint = internals._commentTintLayer!;
    tint.clientWidth = 595;
    tint.clientHeight = 842;
    // The commented run covers x 40–130, y 60–74. A click outside it is a no-op.
    internals._wrapper.dispatch('click', { clientX: 300, clientY: 65 });
    expect(internals._selectedCommentId).toBeNull();
    // A click ON the tinted range selects the thread…
    internals._wrapper.dispatch('click', { clientX: 60, clientY: 65 });
    expect(internals._selectedCommentId).toBe('1');
    // …and a second click on it toggles the selection off again.
    internals._wrapper.dispatch('click', { clientX: 60, clientY: 65 });
    expect(internals._selectedCommentId).toBeNull();
    v.destroy();
  });

  it('the selected balloon becomes scrollable; unselected balloons stay clipped', async () => {
    const { v, internals } = await mountViewer({ showComments: true });
    const before = balloons(internals._commentGutterLayer)[0]!;
    expect(before.style.overflow).toBe('hidden');
    before.dispatch('click', { preventDefault() {} });
    expect(internals._selectedCommentId).toBe('1');
    // The rebuild replaced the balloon; the selected one scrolls its body.
    const selected = balloons(internals._commentGutterLayer)[0]!;
    expect(selected.style['overflow-y']).toBe('auto');
    expect(selected.style.overflow).toBe('');
    // Deselect: back to the clipped form.
    selected.dispatch('click', { preventDefault() {} });
    const after = balloons(internals._commentGutterLayer)[0]!;
    expect(after.style.overflow).toBe('hidden');
    v.destroy();
  });

  it('setShowComments toggles the margin and comment DOM at runtime', async () => {
    const { v, internals } = await mountViewer();
    await v.setShowComments(true);
    expect(internals._wrapper.style.marginRight).toBe('260px');
    expect(balloons(internals._commentGutterLayer)).toHaveLength(1);
    await v.setShowComments(false);
    expect(internals._wrapper.style.marginRight).toBe('');
    expect(internals._commentGutterLayer).toBeNull();
    expect(internals._commentTintLayer).toBeNull();
    await v.setShowComments(true);
    expect(balloons(internals._commentGutterLayer)).toHaveLength(1);
    v.destroy();
  });
});

// ---------------------------------------------------------------------------
// DocxScrollViewer — per-slot comment layers, both render modes.
// ---------------------------------------------------------------------------

async function setupScroll(
  opts: Record<string, unknown> = {},
  mode: 'main' | 'worker' = 'main',
) {
  installDom();
  const container = makeContainer(400, 400);
  const engine = new FakeDocxEngine(
    3,
    [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 200 },
    ],
    mode,
  );
  seedComments(engine);
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
  scrollHost.clientWidth = 400;
  v.relayout();
  await Promise.resolve();
  await Promise.resolve();
  await new Promise((r) => setTimeout(r, 0));

  /** All balloon elements across mounted slots. */
  function allBalloons(): FakeEl[] {
    const out: FakeEl[] = [];
    for (const slot of scrollHost.children) {
      for (const layer of slot.children) {
        out.push(...(layer.children ?? []).filter((child) => child._listeners.has('click')));
      }
    }
    return out;
  }

  return { v, engine, scrollHost, allBalloons };
}

describe('DocxScrollViewer — showComments option', () => {
  it('builds no comment DOM by default', async () => {
    const { v, allBalloons } = await setupScroll();
    expect(allBalloons()).toHaveLength(0);
    v.destroy();
  });

  it('builds per-slot balloons for the unresolved thread when true (main mode)', async () => {
    const { v, allBalloons } = await setupScroll({ showComments: true });
    // The commented run is fed on every page render, so every mounted slot
    // shows the thread's balloon.
    expect(allBalloons().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('worker mode builds the same balloons from the bitmap path', async () => {
    const { v, engine, allBalloons } = await setupScroll({ showComments: true }, 'worker');
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    expect(engine.renderCalls.length).toBe(0);
    expect(allBalloons().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('setShowComments(true) at runtime mounts layers and re-renders', async () => {
    const { v, allBalloons } = await setupScroll();
    expect(allBalloons()).toHaveLength(0);
    (v as unknown as { setShowComments(value: boolean): void }).setShowComments(true);
    await Promise.resolve();
    await new Promise((r) => setTimeout(r, 0));
    expect(allBalloons().length).toBeGreaterThan(0);
    (v as unknown as { setShowComments(value: boolean): void }).setShowComments(false);
    await Promise.resolve();
    await new Promise((r) => setTimeout(r, 0));
    expect(allBalloons()).toHaveLength(0);
    v.destroy();
  });
});
