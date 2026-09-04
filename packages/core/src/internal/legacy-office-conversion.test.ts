import { describe, expect, it, vi } from 'vitest';
import type { LegacyOfficeConverter } from '../conversion/legacy-office.js';
import { bindLegacyOfficeConversionSignal } from './legacy-office-conversion.js';

describe('bindLegacyOfficeConversionSignal', () => {
  const converter: LegacyOfficeConverter = {
    convert: vi.fn(async () => ({ bytes: new Uint8Array() })),
  };

  it('does not create conversion options when the feature is omitted', () => {
    const bound = bindLegacyOfficeConversionSignal(
      undefined,
      'docx',
      new AbortController().signal,
    );
    expect(bound.options).toBeUndefined();
    expect(() => bound.cleanup()).not.toThrow();
  });

  it('combines caller and lifecycle cancellation without mutating caller options', () => {
    const caller = new AbortController();
    const lifecycle = new AbortController();
    const source = { doc: { converter, signal: caller.signal } };
    const bound = bindLegacyOfficeConversionSignal(source, 'docx', lifecycle.signal);

    expect(bound.options).not.toBe(source);
    expect(bound.options?.doc?.signal).not.toBe(caller.signal);
    expect(bound.options?.doc?.signal?.aborted).toBe(false);
    lifecycle.abort();
    expect(bound.options?.doc?.signal?.aborted).toBe(true);
    expect(source.doc.signal.aborted).toBe(false);
    bound.cleanup();
  });

  it('does not bind a lifecycle signal to an unselected format', () => {
    const source = { xls: { converter } };
    const bound = bindLegacyOfficeConversionSignal(
      source,
      'docx',
      new AbortController().signal,
    );
    expect(bound.options).toBe(source);
    expect(bound.options?.doc).toBeUndefined();
  });

  it('removes combined listeners when a load settles', () => {
    const caller = new AbortController();
    const lifecycle = new AbortController();
    const bound = bindLegacyOfficeConversionSignal(
      { doc: { converter, signal: caller.signal } },
      'docx',
      lifecycle.signal,
    );

    bound.cleanup();
    caller.abort();
    expect(bound.options?.doc?.signal?.aborted).toBe(false);
  });
});
