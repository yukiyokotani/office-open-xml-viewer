import { afterEach, describe, expect, it, vi } from 'vitest';

const renderSlideWithEmbeddedFonts = vi.hoisted(() => vi.fn(async (..._args: unknown[]) => undefined));
vi.mock('./renderer', () => ({ renderSlideWithEmbeddedFonts, dropImageBitmapCache: vi.fn() }));

import { PptxPresentation } from './presentation.js';

afterEach(() => { vi.unstubAllGlobals(); renderSlideWithEmbeddedFonts.mockClear(); });

describe('PptxPresentation canvas font ownership', () => {
  it('uses the popup document for a foreign-realm canvas and the main document for OffscreenCanvas', async () => {
    const mainFonts = {} as FontFaceSet;
    const popupFonts = {} as FontFaceSet;
    class MainCanvas {}
    class PopupCanvas {
      readonly nodeType = 1;
      readonly localName = 'canvas';
      readonly offsetWidth = 720;
      readonly ownerDocument = { defaultView: { HTMLCanvasElement: PopupCanvas }, fonts: popupFonts };
    }
    class FakeOffscreenCanvas { readonly width = 100; readonly height = 100; }
    vi.stubGlobal('HTMLCanvasElement', MainCanvas);
    vi.stubGlobal('document', { fonts: mainFonts });

    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance._preflight = {
      slideCount: 1, slideWidth: 914400, slideHeight: 914400,
      majorFont: null, minorFont: null, defaultTextColor: null,
    };
    instance._slides = { withSlide: async (_index: number, read: (slide: object) => unknown) => read({ elements: [] }) };
    instance._assertResourceHealthy = () => undefined;
    instance._assertSlideIndex = () => undefined;
    instance._waitForSlide = async () => undefined;
    instance._rethrowWithResourceFailure = (error: unknown) => { throw error; };
    const selectRoutes = vi.fn(async (requests: unknown[], set: FontFaceSet | null) => ({ set }));
    instance._officeRoutesForRequests = selectRoutes;
    const presentation = instance as unknown as PptxPresentation;

    await presentation.renderSlide(new PopupCanvas() as unknown as HTMLCanvasElement, 0);
    expect(selectRoutes).toHaveBeenLastCalledWith(expect.any(Array), popupFonts);
    expect(renderSlideWithEmbeddedFonts.mock.calls[0]?.[4]).toMatchObject({
      width: 720,
      officeFontRoutes: { set: popupFonts },
    });

    await presentation.renderSlide(new FakeOffscreenCanvas() as unknown as OffscreenCanvas, 0);
    expect(selectRoutes).toHaveBeenLastCalledWith(expect.any(Array), mainFonts);
    expect(renderSlideWithEmbeddedFonts.mock.calls[1]?.[4]).toMatchObject({
      width: 960,
      officeFontRoutes: { set: mainFonts },
    });
  });
});
