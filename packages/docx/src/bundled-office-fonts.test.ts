import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { register, unregister } = vi.hoisted(() => ({
  register: vi.fn(), unregister: vi.fn(),
}));
vi.mock('@silurus/ooxml-core', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  registerEmbeddedFonts: register,
  unregisterEmbeddedFonts: unregister,
  parseOpenTypeResourceMetrics: () => null,
}));
vi.mock('./assets/carlito/urls.js', () => ({
  CARLITO_URLS: { regular: '/regular.ttf', bold: '/bold.ttf',
    italic: '/italic.ttf', boldItalic: '/bold-italic.ttf' },
}));

import { loadBundledCalibri, unloadBundledOfficeFonts } from './bundled-office-fonts.js';

describe('offline DOCX font fallback', () => {
  beforeEach(() => {
    register.mockImplementation(async (faces: Array<{ weight: string; style: string }>) =>
      faces.map((face) => ({ weight: face.weight, style: face.style, family: '__ooxml_docx_bundled_carlito' })));
    vi.stubGlobal('fetch', vi.fn(async () => ({ ok: true, arrayBuffer: async () => new Uint8Array([1, 2, 3]).buffer })));
  });
  afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

  it('loads only unresolved requested Calibri style tuples and releases their faces', async () => {
    const loaded = await loadBundledCalibri([
      { family: 'Calibri', weight: 400 },
      { family: 'Calibri', weight: 700, style: 'italic' },
      { family: 'Arial', weight: 400 },
    ], new Set(['calibri:400:normal']));
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(fetch).toHaveBeenCalledWith('/bold-italic.ttf');
    expect(loaded.routes).toMatchObject([{
      requestedFamily: 'Calibri', source: 'substitute',
      weight: 700, style: 'italic',
    }]);
    expect(register).toHaveBeenCalledTimes(1);
    unloadBundledOfficeFonts(loaded.faces);
    expect(unregister).toHaveBeenCalledWith(loaded.faces);
  });

  it('leaves an embedded or exact local tuple untouched', async () => {
    const loaded = await loadBundledCalibri(
      [{ family: 'Calibri', weight: 400 }], new Set(['calibri:400:normal']),
    );
    expect(loaded).toEqual({ faces: [], routes: [] });
    expect(fetch).not.toHaveBeenCalled();
    expect(register).not.toHaveBeenCalled();
  });
});
