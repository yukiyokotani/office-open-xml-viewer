import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { computeMdw } from './renderer.js';

beforeEach(() => vi.stubGlobal('navigator', { platform: 'Win32', userAgent: 'Windows' }));
afterEach(() => vi.unstubAllGlobals());

describe('ECMA-376 maximum digit width authority', () => {
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
