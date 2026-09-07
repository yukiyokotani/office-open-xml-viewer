import { describe, expect, it } from 'vitest';
import {
  validateLegacyPptSourceDescriptor,
  type LegacyPptDirectSourceDescriptor,
} from './legacy-ppt-source.js';

const valid: LegacyPptDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-ppt-source/v1',
  builtin: 'ppt',
  wasmUrl: 'https://cdn.example.test/legacy_ppt_bg.wasm',
};

describe('legacy PPT direct source descriptor', () => {
  it.each([
    'http://example.test/ppt.wasm',
    'https://example.test/ppt.wasm',
    'file:///opt/assets/ppt.wasm',
    'blob:https://example.test/1234',
    'data:application/wasm;base64,AGFzbQ==',
  ])('accepts supported absolute URL %s', (wasmUrl) => {
    expect(validateLegacyPptSourceDescriptor({ ...valid, wasmUrl }).wasmUrl).toBe(wasmUrl);
  });

  it.each([
    null,
    [],
    {},
    { ...valid, protocol: 'ooxml-legacy-ppt-source/v2' },
    { ...valid, builtin: 'doc' },
    { ...valid, wasmUrl: '' },
    { ...valid, wasmUrl: './ppt.wasm' },
    { ...valid, wasmUrl: ' javascript:alert(1)' },
    { ...valid, wasmUrl: 'javascript:alert(1)' },
    { ...valid, moduleUrl: 'https://example.test/module.js' },
    { ...valid, factory: () => undefined },
  ])('rejects malformed or extensible input %#', (value) => {
    expect(() => validateLegacyPptSourceDescriptor(value)).toThrow(TypeError);
  });

  it('returns a frozen detached exact copy that survives structured clone', () => {
    const input = { ...valid };
    const result = validateLegacyPptSourceDescriptor(input);
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
    expect(() => validateLegacyPptSourceDescriptor(input)).toThrow(/data properties/);
    expect(reads).toBe(0);
  });
});
