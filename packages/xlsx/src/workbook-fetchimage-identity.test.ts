import { afterEach, describe, expect, it, vi } from 'vitest';
import type { RenderViewportOptions, Worksheet } from './types.js';

const { renderWorksheetViewport } = vi.hoisted(() => ({
  renderWorksheetViewport: vi.fn(async (..._args: unknown[]) => undefined),
}));

vi.mock('./render-orchestrator.js', () => ({ renderWorksheetViewport }));

import { XlsxWorkbook } from './workbook.js';

afterEach(() => { vi.unstubAllGlobals(); renderWorksheetViewport.mockClear(); });

describe('XlsxWorkbook.renderViewport() fetchImage identity', () => {
  /**
   * The stable instance closure keys the shared image caches, render-pass lease
   * counter, and destroy-time cache drops. A per-call closure would split that
   * namespace and leave its decoded images alive after destroy().
   */
  it('uses the instance fetchImage closure owned by the workbook', async () => {
    const stableClosure = vi.fn(async () => new Blob());
    const minimalWorksheet = {} as Worksheet;
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance.parsedWorkbook = { styles: {} };
    instance.sheetCache = new Map([[0, minimalWorksheet]]);
    instance.imageCache = new Map();
    instance._fetchImage = stableClosure;

    await (instance as unknown as XlsxWorkbook).renderViewport(
      {} as HTMLCanvasElement,
      0,
      { row: 1, col: 1, rows: 1, cols: 1 },
    );

    const opts = renderWorksheetViewport.mock.calls[0]?.[3] as RenderViewportOptions;
    expect(opts.fetchImage).toBe(stableClosure);
  });

  it('selects the popup canvas FontFaceSet and retains the main document for OffscreenCanvas', async () => {
    class MainCanvas {}
    class PopupCanvas {
      readonly nodeType = 1;
      readonly localName = 'canvas';
      readonly ownerDocument = { defaultView: { HTMLCanvasElement: PopupCanvas }, fonts: popupFonts };
    }
    class FakeOffscreenCanvas { readonly width = 100; readonly height = 100; }
    const mainFonts = {} as FontFaceSet;
    const popupFonts = {} as FontFaceSet;
    vi.stubGlobal('HTMLCanvasElement', MainCanvas);
    vi.stubGlobal('document', { fonts: mainFonts });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance.parsedWorkbook = { styles: {} };
    instance.sheetCache = new Map([[0, {} as Worksheet]]);
    instance._fetchImage = vi.fn(async () => new Blob());
    const popupRoutes = { popup: true };
    const mainRoutes = { main: true };
    instance.retainedFontSets = new Map([
      [popupFonts, { loaded: { office: { routes: popupRoutes } } }],
      [mainFonts, { loaded: { office: { routes: mainRoutes } } }],
    ]);

    await (instance as unknown as XlsxWorkbook).renderViewport(
      new PopupCanvas() as unknown as HTMLCanvasElement,
      0,
      { row: 1, col: 1, rows: 1, cols: 1 },
    );
    expect((renderWorksheetViewport.mock.calls[0]?.[3] as RenderViewportOptions).officeFontRoutes).toBe(popupRoutes);

    await (instance as unknown as XlsxWorkbook).renderViewport(
      new FakeOffscreenCanvas() as unknown as OffscreenCanvas,
      0,
      { row: 1, col: 1, rows: 1, cols: 1 },
    );
    expect((renderWorksheetViewport.mock.calls[1]?.[3] as RenderViewportOptions).officeFontRoutes).toBe(mainRoutes);
  });
});
