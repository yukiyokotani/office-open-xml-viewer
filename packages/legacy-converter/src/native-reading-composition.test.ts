import { describe, expect, it } from 'vitest';
import { legacyDocSource, type LegacyDocSourceOptions } from './legacy-doc.js';
import { createLegacyDocSourceEngine, type LegacyDocGlue, type LegacyDocNativeDocument } from './direct-doc-engine.js';
import { readLegacyDocSourceModuleConfig, readLegacySourceModuleConfig } from './source-module-config.js';

const wasmUrl = 'https://example.test/composed-doc.wasm';
const reading = {
  nativeContourPolicy: 'relocateForReading',
  nativeWordBreakingPolicy: 'simplifyForReading',
  nativePictureBulletPolicy: 'storedSizeForReading',
} as const;

describe('independent DOC reading policy composition', () => {
  it('transports all three choices in one closed DOC descriptor', () => {
    const source = legacyDocSource({ wasmUrl, maxInputBytes: 4096, ...reading });
    const descriptor = source.beginLoad().module;
    expect(descriptor.target).toBe('docx');
    expect(descriptor.config).toEqual({ wasmUrl, maxInputBytes: 4096, ...reading });
    expect(readLegacyDocSourceModuleConfig(structuredClone(descriptor.config)))
      .toEqual({ wasmUrl, maxInputBytes: 4096, ...reading });
    expect(() => readLegacySourceModuleConfig(descriptor.config, 'legacy PPT'))
      .toThrow('unknown field');
    expect(legacyDocSource({ wasmUrl }).beginLoad().module.config)
      .toEqual({ wasmUrl, maxInputBytes: 256 * 1024 * 1024 });
  });

  it('keeps concurrent mixed and single-capability opens independent of the strict constructor', async () => {
    const calls: unknown[][] = [];
    class Archive implements LegacyDocNativeDocument {
      constructor(...args: [Uint8Array, number?, boolean?, boolean?, boolean?]) { calls.push(args); }
      free(): void {}
      open_document_cursor(): void {}
      pull_document_chunk(): Uint8Array { return new Uint8Array(); }
      document_chunk_done(): boolean { return true; }
      acknowledge_document_chunk(): void {}
      cancel_document_cursor(): void {}
      close_document_session(): void {}
      assert_healthy(): void {}
      extract_image(): Uint8Array { return new Uint8Array(); }
    }
    const glue: LegacyDocGlue = { default: async () => undefined, LegacyDocDocument: Archive };
    const engine = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array(), 4096);
    const requests = [reading, { nativeWordBreakingPolicy: 'simplifyForReading' },
      { nativePictureBulletPolicy: 'storedSizeForReading' }, {}] as const;
    const opened = await Promise.all(requests.map((request, index) =>
      engine.open(new Uint8Array([index]), wasmUrl, undefined, request)));
    try {
      calls.sort((a, b) => (a[0] as Uint8Array)[0] - (b[0] as Uint8Array)[0]);
      expect(calls).toEqual([
        [new Uint8Array([0]), 4096, true, true, true],
        [new Uint8Array([1]), 4096, false, true, false],
        [new Uint8Array([2]), 4096, false, false, true],
        [new Uint8Array([3]), 4096],
      ]);
    } finally { for (const archive of opened) archive.closeArchive(); }
  });

  it('rejects one invalid member even when the other reading choices are valid', () => {
    const invalid = { wasmUrl, ...reading, nativePictureBulletPolicy: null };
    expect(() => legacyDocSource(invalid as unknown as LegacyDocSourceOptions)).toThrow(TypeError);
    expect(() => readLegacyDocSourceModuleConfig(invalid)).toThrow(TypeError);
    // A failed mixed selection cannot contaminate a subsequent strict descriptor.
    expect(readLegacyDocSourceModuleConfig(legacyDocSource({ wasmUrl }).beginLoad().module.config))
      .toEqual({ wasmUrl, maxInputBytes: 256 * 1024 * 1024,
        nativeContourPolicy: 'strict', nativeWordBreakingPolicy: 'strict', nativePictureBulletPolicy: 'strict' });
  });
});
