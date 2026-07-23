import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';
import type { PptxPresentation } from './presentation';
import type { PptxTextRunInfo } from './renderer';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX2 — destroy() must drop the find state (twin of the DocxViewer test).
 * findNext()/findPrev() on a torn-down viewer must return null, not a stale
 * match pointing into a dead viewer.
 */
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

/** installDom + a document whose <canvas> els carry a measuring 2d context. */
function installFindDom(): void {
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

describe('PptxViewer.destroy() — find state invalidation', () => {
  it('findNext()/findPrev() return null after destroy() (no stale matches)', async () => {
    installFindDom();
    const parent = makeEl('div');
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
    const engine = new FakePptxEngine(1, 9144000, 6858000);
    engine.feedTextRuns = [run('hello world')];
    (v as unknown as { engine: PptxPresentation }).engine = engine.asPres();

    // A live find with a real active match…
    const matches = await v.findText('hello');
    expect(matches).toHaveLength(1);
    const active = await v.findNext();
    expect(active).not.toBeNull();
    expect(active?.matchIndex).toBe(0);

    // …must not survive teardown.
    v.destroy();
    expect(await v.findNext()).toBeNull();
    expect(await v.findPrev()).toBeNull();
  });
});
