import { describe, expect, it } from 'vitest';
import { legacyXlsHostServices } from '@silurus/ooxml-core/internal/legacy-xls-source';
import { createLegacyXlsSource } from './direct-xls.js';
import { measureLegacyXlsNormalFontInDocument } from './xls-font-worker.js';

describe('createLegacyXlsSource', () => {
  it('returns a frozen validated descriptor without initializing WASM', () => {
    const descriptor = createLegacyXlsSource({ wasmUrl: 'https://example.test/direct-xls.wasm' });
    expect(descriptor).toEqual({
      protocol: 'ooxml-legacy-xls-source/v1',
      builtin: 'xls',
      wasmUrl: 'https://example.test/direct-xls.wasm',
    });
    expect(Object.isFrozen(descriptor)).toBe(true);
  });

  it('binds the measurement policy the spreadsheet host calls', () => {
    const descriptor = createLegacyXlsSource({ wasmUrl: 'https://example.test/direct-xls.wasm' });
    const services = legacyXlsHostServices(descriptor);
    expect(services).toBeDefined();
    const measure = () => 7;
    // The caller's measurement wins; otherwise the source's default applies
    // where a document exists, and nothing is measured without one.
    expect(services!.resolve(measure)).toBe(measure);
    expect(services!.resolve(undefined)).toBe(
      typeof document === 'undefined' ? undefined : measureLegacyXlsNormalFontInDocument,
    );
    // The descriptor itself stays a plain structured-clone-safe value.
    expect(structuredClone(descriptor)).toEqual(descriptor);
  });
});
