import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const config = readFileSync(new URL('../astro.config.mjs', import.meta.url), 'utf8');

describe('official-site client module initialization', () => {
  it('preserves native top-level await ordering for WASM-backed viewers', () => {
    expect(config).toContain("build: { target: 'esnext' }");
    expect(config).not.toContain("from 'vite-plugin-top-level-await'");
    expect(config).not.toContain('topLevelAwait()');
  });
});
