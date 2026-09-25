import { afterEach, describe, expect, it, vi } from 'vitest';
import { cjkLangFromLanguage, resolveCjkFallback } from './cjk-fallback.js';

afterEach(() => vi.unstubAllGlobals());

describe('CJK fallback language policy', () => {
  it.each([
    ['zh', 'sc'], ['zh-CN', 'sc'], ['zh-SG', 'sc'], ['zh-Hans', 'sc'],
    ['zh-TW', 'tc'], ['zh-Hant', 'tc'], ['zh-HK', 'hk'], ['zh-MO', 'hk'],
    ['zh-Hans-HK', 'sc'], ['zh-Hans-TW', 'sc'], ['zh-Hant-CN', 'tc'], ['zh-Hant-HK', 'tc'],
    ['JA-jp', 'jp'], ['ko-KR', 'kr'], ['ko-Hang', 'kr'], ['ja-Hani', 'jp'], ['zh-Latn', null], ['en-US', null], ['', null],
    ['not_a_locale', null], ['en-x-zh-Hant', null], ['zh-u-rg-twzzzz', 'sc'],
  ])('%s maps to %s', (language, expected) => {
    expect(cjkLangFromLanguage(language)).toBe(expected);
  });

  it('uses HTML before ordered browser preferences and supports explicit overrides', () => {
    vi.stubGlobal('document', { documentElement: { lang: 'zh-TW' } });
    vi.stubGlobal('navigator', { languages: ['ja', 'zh-CN'], language: 'ko' });
    expect(resolveCjkFallback()).toBe('tc');
    expect(resolveCjkFallback('auto')).toBe('tc');
    expect(resolveCjkFallback('sc')).toBe('sc');
  });

  it('skips irrelevant/invalid languages and falls back to navigator.language', () => {
    vi.stubGlobal('document', { documentElement: { lang: 'en' } });
    vi.stubGlobal('navigator', { languages: ['en', 'invalid_locale', 'zh-CN', 'ja'], language: 'ko' });
    expect(resolveCjkFallback()).toBe('sc');
    vi.stubGlobal('navigator', { languages: ['en'], language: 'ko' });
    expect(resolveCjkFallback()).toBe('kr');
  });

  it('rejects unsupported options before loading', () => {
    expect(() => resolveCjkFallback('ja' as never)).toThrow(TypeError);
  });

  it('retains a deterministic JP default without browser globals', () => {
    vi.stubGlobal('document', undefined);
    vi.stubGlobal('navigator', undefined);
    expect(resolveCjkFallback()).toBe('jp');
    expect(resolveCjkFallback('hk')).toBe('hk');
  });
});
