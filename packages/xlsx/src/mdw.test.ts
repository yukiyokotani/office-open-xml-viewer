import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { bindXlsxOfficeFontRoutes, bindXlsxWorksheetOfficeFontRoutes, computeMdw, getGridGeometryForWorksheet, getMdwForWorksheet } from './renderer.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';

beforeEach(() => vi.stubGlobal('navigator', { platform: 'Win32', userAgent: 'Windows' }));
afterEach(() => vi.unstubAllGlobals());

function measuringContext(width: number): CanvasRenderingContext2D {
  let savedFont = '';
  return {
    font: '',
    save(this: { font: string }) { savedFont = this.font; },
    restore(this: { font: string }) { this.font = savedFont; },
    measureText: () => ({ width }),
  } as unknown as CanvasRenderingContext2D;
}

describe('ECMA-376 maximum digit width authority', () => {
  it('keeps the painted fallback face authoritative when Calibri is unavailable', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const worksheet = { defaultFontFamily: 'Calibri', defaultFontSize: 12 };
    // Canvas would measure the substituted 9px face; the Office Calibri hmtx
    // maximum is 1038/2048 em, i.e. 8.109375 CSS px before Mac quantization.
    expect(getMdwForWorksheet(worksheet)).toBe(9);
    // A bound route map alone does not establish a paintable face.
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0], {});
    expect(getMdwForWorksheet(worksheet)).toBe(9);
    // A completed lookup cannot make the unavailable Calibri face paint text.
    // Narrowing columns to its catalog digit width would newly clip fallback
    // glyphs, so the same painted fallback remains authoritative.
    bindXlsxOfficeFontRoutes(measuringContext(9), worksheet as Parameters<typeof bindXlsxOfficeFontRoutes>[1], {}, false);
    expect(getMdwForWorksheet(worksheet)).toBe(9);
  });

  it('keeps cached column geometry stable when no paint face changes', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const worksheet = {
      defaultFontFamily: 'Calibri', defaultFontSize: 12,
      defaultColWidth: 10, defaultRowHeight: 15, colWidths: {}, rowHeights: {},
    } as Parameters<typeof getGridGeometryForWorksheet>[0];
    const routes = {};
    const ctx = measuringContext(9);
    bindXlsxOfficeFontRoutes(ctx, worksheet, routes, false);
    const before = getGridGeometryForWorksheet(worksheet);
    bindXlsxOfficeFontRoutes(ctx, worksheet, routes, false);
    const after = getGridGeometryForWorksheet(worksheet);
    expect(before.maximumDigitWidth).toBe(9);
    expect(after.maximumDigitWidth).toBe(9);
    expect(after.col.sizeOf(1)).toBe(before.col.sizeOf(1));
  });

  it('keeps two viewer-owned worksheet projections independent', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const source = {
      defaultFontFamily: 'Calibri', defaultFontSize: 12,
      defaultColWidth: 10, defaultRowHeight: 15,
      colWidths: {}, rowHeights: {}, rows: [],
    } as unknown as Parameters<typeof createSheetViewModel>[0];
    const first = createSheetViewModel(source);
    const second = createSheetViewModel(source);
    bindXlsxOfficeFontRoutes(measuringContext(9), first, {}, false);
    bindXlsxOfficeFontRoutes(measuringContext(12), second, {}, false);
    expect(getGridGeometryForWorksheet(first).maximumDigitWidth).toBe(9);
    expect(getGridGeometryForWorksheet(second).maximumDigitWidth).toBe(12);
    expect(getGridGeometryForWorksheet(first).maximumDigitWidth).toBe(9);
  });

  it('measures a popup-owned application face on its own Canvas context', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    class MainOffscreenCanvas {
      getContext() { return { font: '', measureText: () => ({ width: 9 }) }; }
    }
    vi.stubGlobal('OffscreenCanvas', MainOffscreenCanvas);
    class PopupCanvas { constructor(readonly ownerDocument: { fonts: FontFaceSet }) {} }
    vi.stubGlobal('HTMLCanvasElement', PopupCanvas);
    const popupFonts = { *[Symbol.iterator]() { yield { family: 'Calibri' }; } } as unknown as FontFaceSet;
    const initialFont = 'italic 13px serif';
    let savedFont = '';
    const ctx = { canvas: new PopupCanvas({ fonts: popupFonts }), font: initialFont,
      save(this: { font: string }) { savedFont = this.font; },
      restore(this: { font: string }) { this.font = savedFont; },
      measureText: () => ({ width: 12 }) } as unknown as CanvasRenderingContext2D;
    const worksheet = { defaultFontFamily: 'Calibri', defaultFontSize: 12 } as Parameters<typeof bindXlsxOfficeFontRoutes>[1];
    bindXlsxOfficeFontRoutes(ctx, worksheet, {}, false);
    expect(getMdwForWorksheet(worksheet)).toBe(12);
    expect(ctx.font).toBe(initialFont);
  });

  it('binds a loaded regular alias from a foreign canvas owner document', () => {
    class MainCanvas {}
    class PopupCanvas {
      readonly nodeType = 1;
      readonly localName = 'canvas';
      readonly ownerDocument: { fonts: FontFaceSet; defaultView: { HTMLCanvasElement: typeof PopupCanvas } };
      constructor(fonts: FontFaceSet) {
        this.ownerDocument = { fonts, defaultView: { HTMLCanvasElement: PopupCanvas } };
      }
    }
    vi.stubGlobal('HTMLCanvasElement', MainCanvas);
    vi.stubGlobal('document', { fonts: [] });
    const popupFonts = [{ family: 'Lato', status: 'loaded', weight: '400', style: 'normal' }] as unknown as FontFaceSet;
    const observedFonts: string[] = [];
    const ctx = {
      canvas: new PopupCanvas(popupFonts), font: '',
      save() {}, restore() {},
      measureText(this: { font: string }) {
        observedFonts.push(this.font);
        return { width: 9 };
      },
    } as unknown as CanvasRenderingContext2D;
    const worksheet = { defaultFontFamily: 'Lato Regular', defaultFontSize: 11 } as Parameters<typeof bindXlsxOfficeFontRoutes>[1];
    bindXlsxOfficeFontRoutes(ctx, worksheet, {}, true);
    expect(observedFonts[0]).toContain('"Lato Regular", "Lato",');
  });

  it('keeps actual font measurements authoritative for exact and app faces', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const worksheet = { defaultFontFamily: 'Calibri', defaultFontSize: 12 };
    const route = {
      requestedFamily: 'Calibri', family: '__local_calibri', source: 'local',
      resourceIdentity: 'office-local:test', weight: 400, style: 'normal',
      metric: { family: '__local_calibri' },
    } as const;
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0], { calibri: route });
    expect(getMdwForWorksheet(worksheet)).toBe(9);
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0]);
    vi.stubGlobal('document', {
      fonts: [{ family: 'Calibri' }],
      createElement: () => ({ getContext: () => ({ font: '', measureText: () => ({ width: 9 }) }) }),
    });
    expect(getMdwForWorksheet(worksheet)).toBe(9);
  });

  it('uses the painted fallback rather than an authored bold Normal tuple that is unavailable', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const worksheet = {
      defaultFontFamily: 'Meiryo UI', defaultFontSize: 12,
      defaultFontBold: true,
    };
    bindXlsxOfficeFontRoutes(measuringContext(9), worksheet as Parameters<typeof bindXlsxOfficeFontRoutes>[1], {}, false);
    // The catalogued bold face would quantize to 11px, but cannot paint.
    expect(getMdwForWorksheet(worksheet)).toBe(9);
  });

  it('uses Mac Excel point-quantized widths across a font-size boundary', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    let digitWidth = 0;
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '',
      measureText: () => ({ width: digitWidth }),
    }) }) });

    // Independent Office-produced controls: Calibri 9/10/11/12/14 pt have
    // first-column widths 45/45/53/53/61 pt on Mac Excel. Digit advances
    // come from the Office-bundled Calibri face (1038/2048 em).
    for (const [sizePt, expectedMdw] of [[9, 7], [10, 7], [11, 8], [12, 8], [14, 9]]) {
      digitWidth = 1038 / 2048 * sizePt * 96 / 72;
      expect(computeMdw('Calibri', sizePt)).toBe(expectedMdw);
    }
    // At 11 pt Arial and Meiryo UI independently land in the 6 pt and 7 pt
    // classes of the same rule; no font-name-specific branch is needed.
    digitWidth = 1139 / 2048 * 11 * 96 / 72;
    expect(computeMdw('Arial', 11)).toBe(8);
    digitWidth = 1272 / 2048 * 11 * 96 / 72;
    expect(computeMdw('Meiryo UI', 11)).toBe(9);
  });

  it('retains CSS-pixel quantization on non-Mac platforms', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'Win32', userAgent: 'Windows' });
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 1038 / 2048 * 11 * 96 / 72 }),
    }) }) });
    expect(computeMdw('Calibri', 11)).toBe(7);
  });

  it('measures the retained Calibri resource when supplied by the worksheet', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    let selected = '';
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      set font(value: string) { selected = value; },
      measureText: () => ({ width: 7.6 }),
    }) }) });
    const route = {
      requestedFamily: 'Calibri' as const, family: '__pinned_regular',
      source: 'substitute' as const, resourceIdentity: 'bundled:carlito:test',
      weight: 400 as const, style: 'normal' as const,
      metric: { family: '__pinned_regular' },
    };
    expect(computeMdw('Calibri', 11, route)).toBe(8);
    expect(selected).toContain('"__pinned_regular"');
  });
  it('measures the resolved Normal-style face without family-specific overrides', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', {
      createElement: () => ({
        getContext: () => ({
          font: '',
          measureText: () => ({ width: 6.2 }),
        }),
      }),
    });

    expect(computeMdw('Meiryo UI', 10)).toBe(6);
  });

  it('remeasures after the active font realm changes', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    let width = 6.2;
    vi.stubGlobal('document', {
      createElement: () => ({
        getContext: () => ({
          font: '',
          measureText: () => ({ width }),
        }),
      }),
    });

    expect(computeMdw('Example', 11)).toBe(6);
    width = 8.1;
    expect(computeMdw('Example', 11)).toBe(8);
  });

});
