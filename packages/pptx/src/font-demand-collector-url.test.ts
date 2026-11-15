import { afterEach, expect, it, vi } from 'vitest';
import { fontDemandCollectorUrl } from './font-demand-collector-url.js';

afterEach(() => { vi.unstubAllGlobals(); });

it('resolves the collector sidecar beside the module', () => {
  expect(new URL(fontDemandCollectorUrl()).pathname).toMatch(/\/font-preload-demand\.ts$/);
});

it('answers no URL when the bundle has no module URL to resolve against', () => {
  // An esbuild IIFE leaves import.meta.url undefined: `new URL` throws.
  vi.stubGlobal('URL', class { constructor() { throw new TypeError('Invalid URL'); } });
  expect(fontDemandCollectorUrl()).toBe('');
});
