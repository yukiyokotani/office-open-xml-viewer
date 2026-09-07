import { describe, expect, it, vi } from 'vitest';
import type { LegacyOfficeConverter } from '../conversion/legacy-office.js';
import { buildCfbFixture, buildStoredZip } from '../testing';
import {
  bindLegacyOfficeConversionSignal,
  resolvePptPresentationInput,
} from './legacy-office-conversion.js';

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

const source = {
  protocol: 'ooxml-legacy-ppt-source/v1' as const,
  builtin: 'ppt' as const,
  wasmUrl: 'https://example.test/ppt.wasm',
};

function pptx(): Uint8Array {
  return buildStoredZip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>',
    'ppt/presentation.xml': '<p:presentation/>',
  });
}

describe('resolvePptPresentationInput', () => {
  it('selects direct native PPT without invoking a converter', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'PowerPoint Document']);
    const result = await resolvePptPresentationInput(bytes, {
      ppt: { source },
    });
    expect(result).toMatchObject({ kind: 'legacy-ppt', source });
  });

  it('keeps ordinary PPTX on the OOXML path even when a source is configured', async () => {
    const bytes = pptx();
    const result = await resolvePptPresentationInput(bytes, { ppt: { source } });
    expect(result).toEqual({ kind: 'ooxml', bytes });
  });

  it('rejects ambiguity, wrong families, limits, and abort before engine initialization', async () => {
    const ppt = buildCfbFixture(['Root Entry', 'PowerPoint Document']);
    await expect(resolvePptPresentationInput(ppt, {
      ppt: { source, converter: { convert: vi.fn() } } as never,
    })).rejects.toThrow(/mutually exclusive/);
    await expect(resolvePptPresentationInput(
      buildCfbFixture(['Root Entry', 'Workbook']), { ppt: { source } },
    )).rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(resolvePptPresentationInput(ppt, {
      ppt: { source, maxInputBytes: 1 },
    })).rejects.toMatchObject({ reason: 'source-too-large' });
    const controller = new AbortController();
    controller.abort();
    await expect(resolvePptPresentationInput(ppt, {
      ppt: { source, signal: controller.signal },
    })).rejects.toMatchObject({ reason: 'aborted' });
  });
});
