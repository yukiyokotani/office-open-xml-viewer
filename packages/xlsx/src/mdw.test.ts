import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { bindXlsxOfficeFontRoutes, bindXlsxWorksheetOfficeFontRoutes, computeMdw, getGridGeometryForWorksheet, getMdwForWorksheet } from './renderer.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';

beforeEach(() => vi.stubGlobal('navigator', { platform: 'Win32', userAgent: 'Windows' }));
afterEach(() => vi.unstubAllGlobals());

describe('ECMA-376 maximum digit width authority', () => {
  it('uses the authored Normal face digit scalar when Calibri is unavailable', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    vi.stubGlobal('document', { createElement: () => ({ getContext: () => ({
      font: '', measureText: () => ({ width: 9 }),
    }) }) });
    const worksheet = { defaultFontFamily: 'Calibri', defaultFontSize: 12 };
    // Canvas would measure the substituted 9px face; the Office Calibri hmtx
    // maximum is 1038/2048 em, i.e. 8.109375 CSS px before Mac quantization.
    expect(getMdwForWorksheet(worksheet)).toBe(9);
    // A bound route map does not prove this tuple was attempted: bounded
    // preflight may omit it. Canvas remains authoritative until it completes.
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0], {});
    expect(getMdwForWorksheet(worksheet)).toBe(9);
    // Once the exact tuple was attempted without a retained route, use the
    // authored face's digit scalar rather than substituted Canvas digits.
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0], {}, false, ['calibri']);
    expect(getMdwForWorksheet(worksheet)).toBe(8);
  });

  it('rebuilds cached column geometry when the Normal tuple completes preflight', () => {
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
    bindXlsxWorksheetOfficeFontRoutes(worksheet, routes, false, []);
    const before = getGridGeometryForWorksheet(worksheet);
    bindXlsxWorksheetOfficeFontRoutes(worksheet, routes, false, ['calibri']);
    const after = getGridGeometryForWorksheet(worksheet);
    expect(before.maximumDigitWidth).toBe(9);
    expect(after.maximumDigitWidth).toBe(8);
    expect(after).not.toBe(before);
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
    bindXlsxWorksheetOfficeFontRoutes(first, {}, false, ['calibri']);
    bindXlsxWorksheetOfficeFontRoutes(second, {}, false, []);
    expect(getGridGeometryForWorksheet(first).maximumDigitWidth).toBe(8);
    expect(getGridGeometryForWorksheet(second).maximumDigitWidth).toBe(9);
    expect(getGridGeometryForWorksheet(first).maximumDigitWidth).toBe(8);
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
    bindXlsxOfficeFontRoutes(ctx, worksheet, {}, false, []);
    expect(getMdwForWorksheet(worksheet)).toBe(12);
    expect(ctx.font).toBe(initialFont);
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

  it('uses the authored bold Normal tuple rather than regular digit widths', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', userAgent: 'Macintosh' });
    const worksheet = {
      defaultFontFamily: 'Meiryo UI', defaultFontSize: 12,
      defaultFontBold: true,
    };
    bindXlsxWorksheetOfficeFontRoutes(worksheet as Parameters<typeof bindXlsxWorksheetOfficeFontRoutes>[0], {}, false, ['meiryo ui:700:normal']);
    // Office hmtx: 1386/2048 for bold vs 1272/2048 for regular.
    expect(getMdwForWorksheet(worksheet)).toBe(11);
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
