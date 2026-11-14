import { describe, expect, it } from 'vitest';
import { fontContentKey } from './source-cache.js';

describe('fontContentKey', () => {
  it('keys bytes by content, once per buffer, and never answers for a detached one', () => {
    const bytes = Uint8Array.from({ length: 4096 }, (_, index) => (index * 31) & 0xff);
    const a = bytes.slice().buffer;
    const b = bytes.slice().buffer;
    const changed = bytes.slice();
    changed[4095] ^= 1;
    expect(fontContentKey(a)).toBe(fontContentKey(b));
    expect(fontContentKey(a)).toBe(fontContentKey(a));
    expect(fontContentKey(changed.buffer)).not.toBe(fontContentKey(a));
    // A transferred buffer no longer has its content: never its old key.
    structuredClone(a, { transfer: [a] });
    expect(a.byteLength).toBe(0);
    expect(() => fontContentKey(a)).toThrow(TypeError);
  });
});
