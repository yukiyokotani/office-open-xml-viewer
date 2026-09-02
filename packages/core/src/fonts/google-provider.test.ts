import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetCssCacheForTests } from './google-css.js';
import { GoogleFontsProvider } from './google-provider.js';
import { FontProviderSession, providerFontFamily, providerFontSource } from './provider.js';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
  _resetCssCacheForTests();
  vi.restoreAllMocks();
});

describe('GoogleFontsProvider', () => {
  it('resolves Office substitutes as ordinary provider assets', async () => {
    globalThis.fetch = vi.fn(async () => new Response(`
      @font-face {
        font-family: 'Carlito';
        font-style: italic;
        font-weight: 700;
        src: url(https://fonts.gstatic.test/carlito-bold-italic.woff2) format('woff2');
      }
    `)) as typeof fetch;
    const provider = new GoogleFontsProvider();

    const assets = await provider.resolve(['Calibri'], { signal: new AbortController().signal });

    expect(assets).toEqual([{
      family: 'Calibri',
      source: { url: 'https://fonts.gstatic.test/carlito-bold-italic.woff2' },
      descriptors: { style: 'italic', weight: '700' },
    }]);
    expect(provider.registrationFamily('Calibri', '__isolated')).toBe('Carlito');
  });

  it('keeps script fallback faces under their public family', () => {
    const provider = new GoogleFontsProvider();
    expect(provider.registrationFamily('Noto Sans Thai', '__isolated')).toBe('Noto Sans Thai');
    expect(provider.registrationFamily('Unknown', '__isolated')).toBe('__isolated');
  });

  it('carries Google and substitution semantics through the shared route', async () => {
    globalThis.fetch = vi.fn(async (input) => {
      if (String(input).includes('googleapis')) {
        return new Response(`
          @font-face { font-family: 'Carlito'; src: url(https://fonts.gstatic.test/carlito.woff2); }
          @font-face { font-family: 'Ubuntu'; src: url(https://fonts.gstatic.test/ubuntu.woff2); }
        `);
      }
      return new Response(new Uint8Array([1, 2, 3]));
    }) as typeof fetch;
    const session = new FontProviderSession(new GoogleFontsProvider());

    const resolved = await session.ensure(['Calibri', 'Ubuntu'], null);

    expect(providerFontFamily(resolved.routes, 'Calibri')).toBe('Carlito');
    expect(providerFontSource(resolved.routes, 'Calibri')).toBe('substitute');
    expect(providerFontFamily(resolved.routes, 'Ubuntu')).toBe('Ubuntu');
    expect(providerFontSource(resolved.routes, 'Ubuntu')).toBe('google');
  });
});
