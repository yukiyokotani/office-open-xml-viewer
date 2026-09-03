import { runInNewContext } from 'node:vm';
import { readFileSync } from 'node:fs';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  toMarkdown: vi.fn((bytes: Uint8Array) => [...bytes].join(',')),
}));

vi.mock('../wasm/docx/ooxml_markdown_docx.js', () => ({
  initSync: vi.fn(),
  to_markdown: mocks.toMarkdown,
}));

import { initFromBytes, toMarkdown } from './docx.js';

describe('markdown byte normalization', () => {
  beforeEach(() => {
    mocks.toMarkdown.mockClear();
    initFromBytes(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
  });

  it('accepts an ArrayBuffer created in another realm', () => {
    const foreign = runInNewContext(`(() => {
      const bytes = new Uint8Array([1, 2, 3]);
      return bytes.buffer;
    })()`);
    expect(toMarkdown(foreign as ArrayBuffer)).toBe('1,2,3');
  });

  it('preserves a cross-realm view offset and length', () => {
    const foreign = runInNewContext('new Uint8Array([9, 1, 2, 8]).subarray(1, 3)');
    expect(toMarkdown(foreign as Uint8Array)).toBe('1,2');
  });
});

describe('format entry points', () => {
  it('publishes only independent format subpaths', () => {
    const manifest = JSON.parse(
      readFileSync(new URL('../package.json', import.meta.url), 'utf8'),
    ) as { exports: Record<string, unknown>; dependencies?: Record<string, string> };
    expect(Object.keys(manifest.exports)).toEqual([
      './pptx',
      './pptx/wasm-binary',
      './docx',
      './docx/wasm-binary',
      './xlsx',
      './xlsx/wasm-binary',
    ]);
    expect(manifest.exports['.']).toBeUndefined();
    expect(manifest.dependencies).toBeUndefined();
  });
});
