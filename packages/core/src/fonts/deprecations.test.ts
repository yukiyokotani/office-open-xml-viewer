import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  USE_GOOGLE_FONTS_DEPRECATION_MESSAGE,
  normalizeFontOptions,
  resetFontDeprecationWarningsForTests,
} from './deprecations.js';
import { googleFonts } from './presets/google-fonts.js';
import { fontCssApi } from './source.js';

afterEach(() => {
  resetFontDeprecationWarningsForTests();
  vi.restoreAllMocks();
});

describe('normalizeFontOptions (useGoogleFonts alias, design §9.1)', () => {
  it('maps useGoogleFonts: true to exactly fonts: [googleFonts()] and drops the key', () => {
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const normalized = normalizeFontOptions({ useGoogleFonts: true, password: 'x' });
    expect(normalized).not.toHaveProperty('useGoogleFonts');
    expect(normalized.password).toBe('x');
    expect(normalized.fonts).toEqual([googleFonts()]);
  });

  it('warns with the exact message once per realm, also for false, never for an unset key', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    expect(normalizeFontOptions({ useGoogleFonts: undefined })).toEqual({});
    expect(warn).not.toHaveBeenCalled();
    expect(normalizeFontOptions({ useGoogleFonts: false })).toEqual({});
    normalizeFontOptions({ useGoogleFonts: true });
    normalizeFontOptions({ useGoogleFonts: true });
    expect(warn.mock.calls).toEqual([[USE_GOOGLE_FONTS_DEPRECATION_MESSAGE]]);
    expect(USE_GOOGLE_FONTS_DEPRECATION_MESSAGE).toBe(
      '[ooxml] useGoogleFonts is deprecated and will be removed in 0.91.0. Replace it with fonts: [googleFonts()] — import { googleFonts } from \'@silurus/ooxml\' (also exported from the docx/xlsx/pptx entries). A custom CSS origin is googleFonts({ cssOrigin }); other sources are fontCssApi(), fontCss(), fontFiles() and fontResolver(). See docs/migration-0.90.md.',
    );
    resetFontDeprecationWarningsForTests();
    normalizeFontOptions({ useGoogleFonts: false });
    expect(warn).toHaveBeenCalledTimes(2);
  });

  it('rejects useGoogleFonts: true combined with fonts, but lets false accompany fonts', () => {
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const source = fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}' });
    expect(() => normalizeFontOptions({ useGoogleFonts: true, fonts: [] })).toThrow(
      new TypeError('useGoogleFonts cannot be combined with fonts; add googleFonts() to fonts instead'),
    );
    expect(normalizeFontOptions({ useGoogleFonts: false, fonts: [source] }).fonts).toEqual([source]);
  });

  it('validates fonts and omits an empty list', () => {
    expect(normalizeFontOptions({ fonts: [] })).toEqual({});
    expect(() => normalizeFontOptions({ fonts: [{ kind: 'font-css-api' }] as never })).toThrow(TypeError);
    expect(() => normalizeFontOptions({ fonts: 'google' as never })).toThrow(TypeError);
  });
});
