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

import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
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

describe('markdown projection errors', () => {
  beforeEach(() => {
    mocks.docxToMarkdown.mockReset();
    initDocxFromBytes(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
  });

  it('reconstructs a resource-limit envelope as OoxmlResourceLimitError', () => {
    const usage = {
      archiveEntryCount: 5,
      declaredInflatedBytes: 0,
      distinctInflatedBytes: 0,
      operationInflatedBytes: 0,
    };
    const details = {
      stage: 'container',
      violation: {
        format: 'docx',
        operation: 'docx_to_markdown',
        resource: 'archive',
        metric: 'entry-count',
        limit: 4,
        observed: 5,
        configurable: true,
        usage,
      },
    };
    mocks.docxToMarkdown.mockImplementation(() => {
      throw `OOXML_RESOURCE_LIMIT:${JSON.stringify({ code: 'ooxml-resource-limit', details })}`;
    });
    let error: unknown;
    try {
      docxToMarkdown(new Uint8Array([1]));
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error).toMatchObject({ code: 'ooxml-resource-limit', details });
  });

  it('rethrows any other failure unchanged', () => {
    const failure = new RangeError('unrelated');
    mocks.docxToMarkdown.mockImplementation(() => { throw failure; });
    expect(() => docxToMarkdown(new Uint8Array([1]))).toThrow(failure);
  });
});
