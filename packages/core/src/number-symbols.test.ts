import { describe, expect, it, vi } from 'vitest';

describe('host locale adoption', () => {
  it('follows the adopted main-thread locale over this realm\'s navigator.language', async () => {
    vi.resetModules();
    vi.stubGlobal('navigator', { language: 'en-US' });
    try {
      const { adoptMessageHostLocale, hostLocale, hostNumberSymbols } = await import('./number-symbols');
      expect(hostNumberSymbols()).toEqual({ decimal: '.', group: ',' });
      adoptMessageHostLocale({ type: 'parse' });
      expect(hostLocale()).toBe('en-US');
      adoptMessageHostLocale({ type: 'parse', hostLocale: 'de-DE' });
      expect(hostLocale()).toBe('de-DE');
      expect(hostNumberSymbols()).toEqual({ decimal: ',', group: '.' });
      // A missing or non-string value keeps the adopted locale.
      adoptMessageHostLocale({ type: 'render', hostLocale: 42 });
      adoptMessageHostLocale(null);
      expect(hostLocale()).toBe('de-DE');
    } finally {
      vi.unstubAllGlobals();
      vi.resetModules();
    }
  });
});
