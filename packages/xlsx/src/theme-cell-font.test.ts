import { afterEach, describe, expect, it, vi } from 'vitest';
import { fontStackFor, resolvedThemeCellFont } from './renderer.js';

afterEach(() => vi.unstubAllGlobals());

const theme = {
  themeJapaneseMajorFont: '游ゴシック Light',
  themeJapaneseMinorFont: '游ゴシック',
};
const font = { name: 'Calibri', scheme: 'minor' as const, bold: true, italic: false };

describe('Mac Japanese Excel theme cell font selection', () => {
  it('uses the Jpan theme face for both Calibri- and Arial-named minor fonts', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', language: 'ja-JP', userAgent: 'Macintosh' });
    expect(resolvedThemeCellFont(font, theme)).toEqual({ family: '游ゴシック', weight: 700, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont({ ...font, name: 'Arial' }, theme)).toEqual({ family: '游ゴシック', weight: 700, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont({ ...font, bold: false, italic: true }, theme))
      .toEqual({ family: '游ゴシック', weight: 400, fallbackAlias: 'YuGothic' });
    const selected = resolvedThemeCellFont(font, theme);
    expect(fontStackFor(selected.family, undefined, '', undefined, false, selected.fallbackAlias))
      .toMatch(/^"游ゴシック", "YuGothic", /);
  });

  it('keeps major and minor theme faces distinct', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', language: 'ja-JP', userAgent: 'Macintosh' });
    expect(resolvedThemeCellFont({ ...font, scheme: 'major' }, theme)).toEqual({ family: '游ゴシック Light', weight: 300, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont(font, { ...theme, themeJapaneseMinorFont: 'メイリオ' })).toEqual({ family: 'メイリオ', weight: 700, fallbackAlias: 'Meiryo' });
    expect(resolvedThemeCellFont({ ...font, bold: false }, { ...theme, themeJapaneseMinorFont: 'メイリオ' }))
      .toEqual({ family: 'メイリオ', weight: 400, fallbackAlias: 'Meiryo' });
  });

  it('preserves the authored name without scheme or Jpan entry', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', language: 'ja-JP', userAgent: 'Macintosh' });
    expect(resolvedThemeCellFont({ ...font, scheme: undefined }, theme).family).toBe('Calibri');
    expect(resolvedThemeCellFont(font, { themeJapaneseMinorFont: undefined }).family).toBe('Calibri');
  });

  it('leaves other browser locales and platforms on their existing route', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', language: 'en-US', userAgent: 'Macintosh' });
    expect(resolvedThemeCellFont(font, theme).family).toBe('Calibri');
    vi.stubGlobal('navigator', { platform: 'Win32', language: 'ja-JP', userAgent: 'Windows' });
    expect(resolvedThemeCellFont(font, theme).family).toBe('Calibri');
    vi.stubGlobal('navigator', { platform: 'MacIntel', maxTouchPoints: 5, language: 'ja-JP', userAgent: 'Macintosh' });
    expect(resolvedThemeCellFont(font, theme).family).toBe('Calibri');
  });
});
