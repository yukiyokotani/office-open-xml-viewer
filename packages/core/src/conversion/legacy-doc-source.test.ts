import { describe, expect, it } from 'vitest';
import {
  validateLegacyDocSourceDescriptor,
  type LegacyDocDirectSourceDescriptor,
} from './legacy-doc-source.js';

const valid: LegacyDocDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-doc-source/v1',
  builtin: 'doc',
  wasmUrl: 'https://cdn.example.test/legacy_doc_bg.wasm',
};

describe('legacy DOC direct source descriptor', () => {
  it.each(['http://example.test/doc.wasm', 'https://example.test/doc.wasm',
    'file:///opt/assets/doc.wasm', 'blob:https://example.test/1234',
    'data:application/wasm;base64,AGFzbQ=='])('accepts supported absolute URL %s', (wasmUrl) => {
    expect(validateLegacyDocSourceDescriptor({ ...valid, wasmUrl }).wasmUrl).toBe(wasmUrl);
  });

  it.each([null, [], {}, { ...valid, protocol: 'ooxml-legacy-doc-source/v2' },
    { ...valid, builtin: 'xls' }, { ...valid, wasmUrl: '' },
    { ...valid, wasmUrl: './doc.wasm' }, { ...valid, wasmUrl: 'javascript:alert(1)' },
    { ...valid, moduleUrl: 'https://example.test/module.js' }])('rejects malformed input %#', (value) => {
    expect(() => validateLegacyDocSourceDescriptor(value)).toThrow(TypeError);
  });

  it('returns a frozen detached exact copy', () => {
    const input = { ...valid };
    const result = validateLegacyDocSourceDescriptor(input);
    input.wasmUrl = 'https://example.test/changed.wasm';
    expect(result).not.toBe(input);
    expect(result.wasmUrl).toBe(valid.wasmUrl);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Reflect.ownKeys(result)).toEqual(['protocol', 'builtin', 'wasmUrl']);
    expect(structuredClone(result)).toEqual(result);
  });
});
