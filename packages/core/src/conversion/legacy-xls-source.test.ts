import { describe, expect, it } from 'vitest';
import {
  validateLegacyXlsSourceDescriptor,
  type LegacyXlsDirectSourceDescriptor,
} from './legacy-xls-source.js';

const valid: LegacyXlsDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-xls-source/v1',
  builtin: 'xls',
  wasmUrl: 'https://cdn.example.test/legacy_xls_bg.wasm',
};

describe('legacy XLS direct source descriptor', () => {
  it.each([
    'http://example.test/xls.wasm',
    'https://example.test/xls.wasm',
    'file:///opt/assets/xls.wasm',
    'blob:https://example.test/1234',
    'data:application/wasm;base64,AGFzbQ==',
  ])('accepts supported absolute URL %s', (wasmUrl) => {
    expect(validateLegacyXlsSourceDescriptor({ ...valid, wasmUrl }).wasmUrl).toBe(wasmUrl);
  });

  it.each([
    null,
    [],
    {},
    { ...valid, protocol: 'ooxml-legacy-xls-source/v2' },
    { ...valid, builtin: 'doc' },
    { ...valid, builtin: 'ppt' },
    { ...valid, protocol: 'ooxml-legacy-ppt-source/v1', builtin: 'ppt' },
    { ...valid, wasmUrl: '' },
    { ...valid, wasmUrl: './xls.wasm' },
    { ...valid, wasmUrl: ' javascript:alert(1)' },
    { ...valid, wasmUrl: 'javascript:alert(1)' },
    { ...valid, moduleUrl: 'https://example.test/module.js' },
    { ...valid, factory: () => undefined },
  ])('rejects malformed or extensible input %#', (value) => {
    expect(() => validateLegacyXlsSourceDescriptor(value)).toThrow(TypeError);
  });

  it('returns a frozen detached exact copy that survives structured clone', () => {
    const input = { ...valid };
    const result = validateLegacyXlsSourceDescriptor(input);
    input.wasmUrl = 'https://example.test/changed.wasm';
    expect(result).not.toBe(input);
    expect(result.wasmUrl).toBe(valid.wasmUrl);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Reflect.ownKeys(result)).toEqual(['protocol', 'builtin', 'wasmUrl']);
    expect(structuredClone(result)).toEqual(result);
  });

  it('rejects accessors without evaluating mutable descriptor code', () => {
    let reads = 0;
    const input = {
      protocol: valid.protocol,
      builtin: valid.builtin,
      get wasmUrl() {
        reads += 1;
        return reads === 1 ? valid.wasmUrl : (() => undefined);
      },
    };
    expect(() => validateLegacyXlsSourceDescriptor(input)).toThrow(/data properties/);
    expect(reads).toBe(0);
  });
});
