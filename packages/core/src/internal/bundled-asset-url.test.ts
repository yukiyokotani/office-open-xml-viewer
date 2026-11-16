import { afterEach, describe, expect, it, vi } from 'vitest';
import { realmBaseUrl, resolveBundledAssetUrl } from './bundled-asset-url.js';

afterEach(() => vi.unstubAllGlobals());

const owner = 'file:///build/package/model-source.js';

describe('bundled model-source assets', () => {
  it('resolves root-relative output against the calling page, not a build-time file URL', () => {
    vi.stubGlobal('location', { href: 'https://app.example.test/reports/current?view=1' });
    expect(resolveBundledAssetUrl('/_next/static/media/source.mjs?hash=1', owner))
      .toBe('https://app.example.test/_next/static/media/source.mjs?hash=1');
    expect(resolveBundledAssetUrl('//cdn.example.test/source.wasm', owner))
      .toBe('https://cdn.example.test/source.wasm');
  });

  it('preserves module-relative output, absolute CDN assets and explicit file assets', () => {
    vi.stubGlobal('location', { href: 'https://app.example.test/reports/current' });
    expect(resolveBundledAssetUrl('./reader.mjs', 'https://cdn.example.test/pkg/source.js'))
      .toBe('https://cdn.example.test/pkg/reader.mjs');
    expect(resolveBundledAssetUrl('https://cdn.example.test/reader.wasm', owner))
      .toBe('https://cdn.example.test/reader.wasm');
    expect(resolveBundledAssetUrl('file:///opt/reader.wasm', owner))
      .toBe('file:///opt/reader.wasm');
  });

  it('retains a valid HTTP owner origin for root-relative assets on a CDN', () => {
    vi.stubGlobal('location', { href: 'https://app.example.test/reports/current' });
    expect(resolveBundledAssetUrl('/assets/source.mjs', 'https://cdn.example.test/pkg/source.js'))
      .toBe('https://cdn.example.test/assets/source.mjs');
  });

  it('uses the owner when a blob worker location cannot resolve a root-relative asset', () => {
    vi.stubGlobal('location', { href: 'blob:https://app.example.test/worker-id' });
    expect(resolveBundledAssetUrl('/assets/source.mjs', 'https://cdn.example.test/pkg/source.js'))
      .toBe('https://cdn.example.test/assets/source.mjs');
    expect(resolveBundledAssetUrl('/assets/source.mjs', owner))
      .toBe('file:///assets/source.mjs');
  });

  it('retains Node file resolution when no calling-realm location exists', () => {
    vi.stubGlobal('location', undefined);
    expect(resolveBundledAssetUrl('/opt/reader.wasm', owner)).toBe('file:///opt/reader.wasm');
    expect(resolveBundledAssetUrl('./reader.mjs', owner)).toBe('file:///build/package/reader.mjs');
  });
});

describe('realm base URL (#2212)', () => {
  it('uses the document base URL, which honours <base href>, over the window location', () => {
    vi.stubGlobal('location', { href: 'vscode-webview://panel/index.html?id=1' });
    vi.stubGlobal('document', { baseURI: 'https://ext.example.test/dist/' });
    expect(realmBaseUrl()).toBe('https://ext.example.test/dist/');
    expect(resolveBundledAssetUrl('/assets/source.mjs', owner))
      .toBe('https://ext.example.test/assets/source.mjs');
  });

  it('uses the worker location in a realm without a document', () => {
    vi.stubGlobal('document', undefined);
    vi.stubGlobal('location', { href: 'https://app.example.test/assets/worker.js' });
    expect(realmBaseUrl()).toBe('https://app.example.test/assets/worker.js');
  });
});
