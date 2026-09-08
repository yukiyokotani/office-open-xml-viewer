import { describe, expect, it, vi } from 'vitest';

const initializeGlue = vi.fn();
vi.mock('./wasm-direct-doc/legacy_office_converter.js', () => {
  initializeGlue();
  return { default: vi.fn() };
});

import { createLegacyDocSource } from './direct-doc.js';

describe('createLegacyDocSource', () => {
  it('returns the dedicated DOC asset identity without loading native glue', () => {
    const descriptor = createLegacyDocSource();
    expect(descriptor).toMatchObject({
      protocol: 'ooxml-legacy-doc-source/v1',
      builtin: 'doc',
    });
    expect(new URL(descriptor.wasmUrl).href).toBe(descriptor.wasmUrl);
    expect(new URL(descriptor.wasmUrl).pathname).toContain(
      '/wasm-direct-doc/legacy_office_converter_bg.wasm',
    );
    expect(initializeGlue).not.toHaveBeenCalled();
    expect(Object.isFrozen(descriptor)).toBe(true);
  });

  it('validates and freezes an explicit asset URL without initializing WASM', () => {
    expect(createLegacyDocSource({ wasmUrl: 'https://example.test/direct-doc.wasm' })).toEqual({
      protocol: 'ooxml-legacy-doc-source/v1',
      builtin: 'doc',
      wasmUrl: 'https://example.test/direct-doc.wasm',
    });
    expect(() => createLegacyDocSource({ wasmUrl: './relative.wasm' })).toThrow(TypeError);
    expect(initializeGlue).not.toHaveBeenCalled();
  });
});
