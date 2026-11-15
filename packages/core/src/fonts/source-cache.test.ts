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

  it('stays linear in the bytes for many buffers of one length that differ only at the end', () => {
    // Pairwise byte comparison would read about 16 GiB here.
    const size = 128 * 1024;
    const buffers = Array.from({ length: 512 }, (_, index) => {
      const bytes = new Uint8Array(size).fill(7);
      new DataView(bytes.buffer).setUint32(size - 4, index);
      return bytes.buffer;
    });
    const started = performance.now();
    const keys = buffers.map((buffer) => fontContentKey(buffer));
    expect(performance.now() - started).toBeLessThan(1500);
    expect(new Set(keys).size).toBe(buffers.length);
    expect(fontContentKey(buffers[10]!.slice(0))).toBe(keys[10]);
    expect(fontContentKey(new Uint8Array(size).fill(7).buffer)).not.toBe(keys[0]);
  });

  it('stays linear for many distinct small buffers of one length', () => {
    const buffers = Array.from({ length: 16384 }, (_, index) => {
      const bytes = new Uint8Array(1024);
      new DataView(bytes.buffer).setUint32(0, index);
      return bytes.buffer;
    });
    const started = performance.now();
    const keys = buffers.map((buffer) => fontContentKey(buffer));
    expect(performance.now() - started).toBeLessThan(1500);
    expect(new Set(keys).size).toBe(buffers.length);
    expect(fontContentKey(buffers[16000]!.slice(0))).toBe(keys[16000]);
  });
});
