import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ModelSource } from '@silurus/ooxml-core';

// Issue #1989 paints text whose parsed chain authors no colour in
// PowerPoint's hard default (black). Only slides from this package's own
// parser have that proven chain; a caller-supplied model source keeps the
// 0.89.7 theme or `defaultTextColor` fallback, so white default text on a dark
// slide does not turn invisible.
const rendered = vi.hoisted(() => ({ opts: [] as Array<{ parsedTextChain?: boolean }> }));
vi.mock('./renderer', async (importOriginal) => ({
  ...await importOriginal<typeof import('./renderer')>(),
  renderSlideWithEmbeddedFonts: async (...args: unknown[]) => {
    rendered.opts.push(args[4] as { parsedTextChain?: boolean });
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
const ORIGINAL = { Worker: G.Worker, location: G.location };
afterEach(() => {
  G.Worker = ORIGINAL.Worker;
  G.location = ORIGINAL.location;
  rendered.opts = [];
  vi.restoreAllMocks();
});

type Internals = {
  _parse(): Promise<void>;
  _waitForSlide(): Promise<void>;
  _assertSlideIndex(): void;
  _officeRoutesForRequests(): Promise<{ routes: Map<string, string>; installed: Set<string> }>;
};

async function parsedTextChainFor(variant: 'parser' | 'model-source'): Promise<boolean | undefined> {
  G.Worker = SilentWorker;
  G.location = { href: 'http://localhost/' };
  const proto = PptxPresentation.prototype as unknown as Internals;
  vi.spyOn(proto, '_parse').mockImplementationOnce(async function (this: object) {
    Object.assign(this, {
      _preflight: {
        slideCount: 1, slideWidth: 914400, slideHeight: 914400,
        defaultTextColor: 'FFFFFF', majorFont: null, minorFont: null,
        hlinkColor: null, folHlinkColor: null, embeddedFonts: [], slides: [],
        fontPreloadNames: [],
      },
      _slides: { withSlide: (_index: number, use: (slide: object) => unknown) => use({}), clear: () => undefined },
    });
  });
  vi.spyOn(proto, '_waitForSlide').mockResolvedValue(undefined);
  vi.spyOn(proto, '_assertSlideIndex').mockReturnValue(undefined);
  vi.spyOn(proto, '_officeRoutesForRequests').mockResolvedValue({ routes: new Map(), installed: new Set() });
  const source: ModelSource<'pptx'> = { target: 'pptx', claim: () => true, beginLoad: () => ({
    module: { protocol: 'ooxml-model-source-module/v2', target: 'pptx', moduleUrl: 'https://example.test/source.mjs', config: {} },
    release: () => undefined,
  }) };
  const presentation = await PptxPresentation.load(new ArrayBuffer(0),
    variant === 'model-source' ? { legacyPpt: source } : {});
  const canvas = { width: 0, height: 0 } as unknown as OffscreenCanvas;
  await presentation.renderSlide(canvas, 0, { width: 100, dpr: 1 });
  presentation.destroy();
  return rendered.opts.at(-1)?.parsedTextChain;
}

describe('parsed text-colour chain provenance', () => {
  it('marks slides from the built-in parser as a parsed chain', async () => {
    expect(await parsedTextChainFor('parser')).toBe(true);
  });

  it('keeps the 0.89.7 fallback for a caller-supplied model source', async () => {
    expect(await parsedTextChainFor('model-source')).toBe(false);
  });
});
