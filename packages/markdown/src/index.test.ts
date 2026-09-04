import { runInNewContext } from 'node:vm';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  docxToMarkdown: vi.fn((bytes: Uint8Array) => [...bytes].join(',')),
}));

vi.mock('@silurus/ooxml-pptx/wasm', () => ({ initSync: vi.fn() }));
vi.mock('@silurus/ooxml-xlsx/wasm', () => ({ initSync: vi.fn() }));
vi.mock('@silurus/ooxml-docx/wasm', () => ({
  initSync: vi.fn(),
  docx_to_markdown: mocks.docxToMarkdown,
}));

import { docxToMarkdown, initDocxFromBytes } from './index.js';

describe('markdown byte normalization', () => {
  beforeEach(() => {
    mocks.docxToMarkdown.mockClear();
    initDocxFromBytes(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
  });

  it('accepts an ArrayBuffer created in another realm', () => {
    const foreign = runInNewContext(`(() => {
      const bytes = new Uint8Array([1, 2, 3]);
      return bytes.buffer;
    })()`);
    expect(docxToMarkdown(foreign as ArrayBuffer)).toBe('1,2,3');
  });

  it('preserves a cross-realm view offset and length', () => {
    const foreign = runInNewContext('new Uint8Array([9, 1, 2, 8]).subarray(1, 3)');
    expect(docxToMarkdown(foreign as Uint8Array)).toBe('1,2');
  });
});
