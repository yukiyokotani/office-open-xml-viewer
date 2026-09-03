import { describe, it, expect } from 'vitest';
import { Buffer } from 'node:buffer';
import { assertNotCfbContainer, toArrayBuffer } from './cfb-guard';
import { OoxmlError } from './ooxml-error';
import { buildCfbFixture } from '../testing/cfb-fixture';

/** Build a tiny CFB whose single directory sector contains the given entries,
 *  as a `Uint8Array` (the shared fixture returns an `ArrayBuffer`). Reuses the
 *  shared `buildCfbFixture` builder so this suite and cfb-sniff.test.ts /
 *  the docx/pptx/xlsx load-guard suites all construct CFB bytes the same way,
 *  rather than maintaining a third parallel inline implementation. */
function cfbWith(names: string[]): Uint8Array {
  return new Uint8Array(buildCfbFixture(names));
}

describe('assertNotCfbContainer', () => {
  it('does nothing for a non-CFB (ZIP) buffer', () => {
    const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]);
    expect(() => assertNotCfbContainer(zip)).not.toThrow();
  });

  it('throws OoxmlError code "encrypted" for an EncryptionInfo CFB', () => {
    const cfb = cfbWith(['Root Entry', 'EncryptionInfo']);
    expect(() => assertNotCfbContainer(cfb)).toThrow(OoxmlError);
    try {
      assertNotCfbContainer(cfb);
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(OoxmlError);
      expect((e as OoxmlError).code).toBe('encrypted');
      expect((e as OoxmlError).message).toMatch(/password-protected/i);
    }
  });

  it('throws OoxmlError code "legacy-binary-format" for a WordDocument CFB', () => {
    const cfb = cfbWith(['Root Entry', 'WordDocument']);
    try {
      assertNotCfbContainer(cfb);
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(OoxmlError);
      expect((e as OoxmlError).code).toBe('legacy-binary-format');
      expect((e as OoxmlError).message).toMatch(/legacy binary/i);
    }
  });

  it('throws OoxmlError code "not-ooxml" for an unrecognised CFB', () => {
    const cfb = cfbWith(['Root Entry', 'Mystery']);
    try {
      assertNotCfbContainer(cfb);
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(OoxmlError);
      expect((e as OoxmlError).code).toBe('not-ooxml');
    }
  });

  it('accepts an ArrayBuffer as well as a Uint8Array', () => {
    const cfb = cfbWith(['Root Entry', 'EncryptionInfo']);
    // Copy into a fresh, definitely-ArrayBuffer to exercise the ArrayBuffer arm.
    const ab = new ArrayBuffer(cfb.byteLength);
    new Uint8Array(ab).set(cfb);
    expect(() => assertNotCfbContainer(ab)).toThrow(OoxmlError);
  });
});

describe('toArrayBuffer', () => {
  it('copies exactly a Node Buffer subview instead of exposing its pooled backing store', () => {
    const source = Buffer.from([9, 1, 2, 3, 9]);
    const view = source.subarray(1, 4);

    const result = toArrayBuffer(view);

    expect(result.byteLength).toBe(3);
    expect(new Uint8Array(result)).toEqual(new Uint8Array([1, 2, 3]));
  });
});
