import { afterEach, describe, expect, it, vi } from 'vitest';

// A table whose rows grow paints beyond its authored frame. Main-mode hit
// testing uses the extent the latest successful display render reported
// (TableHitExtents): published after the render, never by a failed render or
// by text collection, bounded, and cleared on destroy. An unrendered slide
// is hit by its authored frame.
const EMU = 12700;
type Opts = { onTableExtent?: (index: number, height: number) => void; onTextRun?: unknown };
const renderer = vi.hoisted(() => ({
  paint: async (_opts: Opts): Promise<void> => undefined,
}));
vi.mock('./renderer', async (importOriginal) => ({
  ...await importOriginal<typeof import('./renderer')>(),
  renderSlideWithEmbeddedFonts: async (...args: unknown[]) => renderer.paint(args[4] as Opts),
}));
vi.mock('./font-requests', async (importOriginal) => ({
  ...await importOriginal<typeof import('./font-requests')>(),
  pptxSlideOfficeFontRequests: () => [],
}));

const { PptxPresentation } = await import('./presentation');

class SilentWorker {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

const G = globalThis as Record<string, unknown>;
const ORIGINAL = { Worker: G.Worker, location: G.location, OffscreenCanvas: G.OffscreenCanvas };
afterEach(() => {
  Object.assign(G, ORIGINAL);
  renderer.paint = async () => undefined;
  vi.restoreAllMocks();
});

type Internals = {
  _parse(): Promise<void>;
  _waitForSlide(): Promise<void>;
  _assertSlideIndex(): void;
  _officeRoutesForRequests(): Promise<{ routes: Map<string, string>; installed: Set<string> }>;
};

const table = {
  type: 'table', x: 0, y: 0, width: 100 * EMU, height: 20 * EMU, rotation: 0, flipH: false, flipV: false,
  cols: [100 * EMU], rows: [{ height: 20 * EMU, cells: [] }],
};

async function loadMain() {
  G.Worker = SilentWorker;
  G.location = { href: 'http://localhost/' };
  G.OffscreenCanvas = class { constructor(public width: number, public height: number) {} };
  const proto = PptxPresentation.prototype as unknown as Internals;
  vi.spyOn(proto, '_parse').mockImplementationOnce(async function (this: object) {
    Object.assign(this, {
      _preflight: {
        slideCount: 20, slideWidth: 914400, slideHeight: 914400,
        defaultTextColor: null, majorFont: null, minorFont: null,
        hlinkColor: null, folHlinkColor: null, embeddedFonts: [], slides: [],
        fontPreloadNames: [],
      },
      _slides: {
        withSlide: (_index: number, use: (slide: object) => unknown) => use({ elements: [table] }),
        clear: () => undefined,
      },
    });
  });
  vi.spyOn(proto, '_waitForSlide').mockResolvedValue(undefined);
  vi.spyOn(proto, '_assertSlideIndex').mockReturnValue(undefined);
  vi.spyOn(proto, '_officeRoutesForRequests').mockResolvedValue({ routes: new Map(), installed: new Set() });
  const presentation = await PptxPresentation.load(new ArrayBuffer(0), {});
  const canvas = { width: 0, height: 0 } as unknown as OffscreenCanvas;
  const render = (slide: number, heightPt: number | null) => {
    renderer.paint = async (opts) => { if (heightPt !== null) opts.onTableExtent?.(0, heightPt * EMU); };
    return presentation.renderSlide(canvas, slide, { width: 100, dpr: 1 });
  };
  const hits = async (slide: number, yPt: number) =>
    (await presentation.getElementContextAt(slide, { x: 10 * EMU, y: yPt * EMU }))?.elementType === 'table';
  return { presentation, canvas, render, hits };
}

describe('main-mode table hit testing', () => {
  it('uses the authored frame until a display render publishes the painted extent', async () => {
    const { presentation, render, hits } = await loadMain();
    expect(await hits(0, 22)).toBe(false);
    await render(0, 24);
    expect(await hits(0, 22)).toBe(true);
    // The authored frame is still reported, with its authored pivot.
    expect((await presentation.getElementContextAt(0, { x: 10 * EMU, y: 22 * EMU }))?.bounds.height).toBe(20 * EMU);
    expect(await hits(0, 25)).toBe(false);
    // A later render with no grown table publishes that too.
    await render(0, null);
    expect(await hits(0, 22)).toBe(false);
    presentation.destroy();
  });

  it('publishes atomically: an in-flight render keeps, a failed render leaves, the previous snapshot', async () => {
    const { presentation, canvas, render, hits } = await loadMain();
    await render(0, 24);
    let finish!: () => void;
    renderer.paint = async (opts) => {
      opts.onTableExtent?.(0, 30 * EMU);
      await new Promise<void>((resolve) => { finish = resolve; });
    };
    const inFlight = presentation.renderSlide(canvas, 0, { width: 100, dpr: 1 });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(await hits(0, 27)).toBe(false);
    expect(await hits(0, 22)).toBe(true);
    finish();
    await inFlight;
    expect(await hits(0, 27)).toBe(true);

    renderer.paint = async (opts) => {
      opts.onTableExtent?.(0, 40 * EMU);
      throw new Error('paint failed');
    };
    await expect(presentation.renderSlide(canvas, 0, { width: 100, dpr: 1 })).rejects.toThrow('paint failed');
    expect(await hits(0, 35)).toBe(false);
    expect(await hits(0, 27)).toBe(true);
    presentation.destroy();
  });

  it('does not publish from text collection', async () => {
    const { presentation, render, hits } = await loadMain();
    await render(0, 24);
    renderer.paint = async (opts) => { opts.onTableExtent?.(0, 50 * EMU); };
    await presentation.collectSlideRuns(0, 100);
    expect(await hits(0, 45)).toBe(false);
    expect(await hits(0, 22)).toBe(true);
    presentation.destroy();
  });

  it('retains the most recently rendered slides only, and clears on destroy', async () => {
    const { presentation, render, hits } = await loadMain();
    for (let slide = 0; slide < 10; slide++) await render(slide, 24);
    const store = (presentation as unknown as { _tableHitExtentsStore: { size: number } })._tableHitExtentsStore;
    expect(store.size).toBe(8);
    expect(await hits(0, 22)).toBe(false);
    expect(await hits(1, 22)).toBe(false);
    expect(await hits(9, 22)).toBe(true);
    expect(await hits(2, 22)).toBe(true);
    presentation.destroy();
    expect(store.size).toBe(0);
  });
});
