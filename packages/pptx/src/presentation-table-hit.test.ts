import { afterEach, describe, expect, it, vi } from 'vitest';

// A table whose rows grow paints beyond its authored frame; main-mode hit
// testing uses the extent the renderer reported (onTableExtent), and sizes
// the tables without painting when the slide was never rendered.
const EMU = 12700;
const rendered = vi.hoisted(() => ({ opts: [] as Array<Record<string, unknown>> }));
vi.mock('./renderer', async (importOriginal) => ({
  ...await importOriginal<typeof import('./renderer')>(),
  renderSlideWithEmbeddedFonts: async (...args: unknown[]) => {
    const opts = args[4] as { onTableExtent?: (index: number, height: number) => void };
    rendered.opts.push(opts as Record<string, unknown>);
    // The row painted through 24pt although authored 20pt.
    opts.onTableExtent?.(0, 24 * 12700);
  },
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
  rendered.opts = [];
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
        slideCount: 1, slideWidth: 914400, slideHeight: 914400,
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
  return PptxPresentation.load(new ArrayBuffer(0), {});
}

describe('main-mode table hit testing', () => {
  it('hits the painted extent of a grown table after a render', async () => {
    const presentation = await loadMain();
    await presentation.renderSlide({ width: 0, height: 0 } as unknown as OffscreenCanvas, 0, { width: 100, dpr: 1 });
    const hit = await presentation.getElementContextAt(0, { x: 10 * EMU, y: 22 * EMU });
    expect(hit?.elementType).toBe('table');
    // The authored frame is still reported, with its authored pivot.
    expect(hit?.bounds.height).toBe(20 * EMU);
    expect(await presentation.getElementContextAt(0, { x: 10 * EMU, y: 25 * EMU })).toBeNull();
    expect(rendered.opts).toHaveLength(1);
    presentation.destroy();
  });

  it('sizes the tables without painting when the slide was never rendered', async () => {
    const presentation = await loadMain();
    const hit = await presentation.getElementContextAt(0, { x: 10 * EMU, y: 22 * EMU });
    expect(hit?.elementType).toBe('table');
    expect(rendered.opts).toHaveLength(1);
    expect(rendered.opts[0]!.measureTablesOnly).toBe(true);
    presentation.destroy();
  });
});
