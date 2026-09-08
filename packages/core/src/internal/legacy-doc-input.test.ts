import { describe, expect, it, vi } from 'vitest';
import { buildCfbFixture, buildStoredZip } from '../testing';
import { normalizeOfficeInput } from '../conversion/legacy-office.js';
import { resolveDocDocumentInput } from './legacy-office-conversion.js';

const source = { protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc',
  wasmUrl: 'https://example.test/doc.wasm' } as const;

describe('direct DOC input selection', () => {
  it('selects only explicitly enabled DOC and keeps converter selection separate', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'WordDocument']);
    const direct = await resolveDocDocumentInput(bytes, { doc: { source } });
    expect(direct).toMatchObject({ kind: 'legacy-doc', source, bytes });
    expect('converter' in direct).toBe(false);

    const convert = vi.fn(async () => { throw new Error('observable converter invocation'); });
    await expect(resolveDocDocumentInput(bytes, { doc: { converter: { convert } } }))
      .rejects.toMatchObject({ reason: 'failed' });
    expect(convert).toHaveBeenCalledTimes(1);
    await expect(resolveDocDocumentInput(bytes)).rejects.toThrow();
    await expect(resolveDocDocumentInput(bytes, { xls: { converter: { convert } } })).rejects.toThrow();
    expect(convert).toHaveBeenCalledTimes(1);
  });

  it('leaves OOXML on the existing path without native admission', async () => {
    const bytes = buildStoredZip({ '[Content_Types].xml': '<Types/>',
      'word/document.xml': '<document/>' });
    expect(await resolveDocDocumentInput(bytes, { doc: { source } }))
      .toEqual({ kind: 'ooxml', bytes });
  });

  it('rejects ambiguity and refuses direct sources in byte-only normalization', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'WordDocument']);
    await expect(resolveDocDocumentInput(bytes, {
      doc: { source, converter: { convert: vi.fn() } } as never,
    })).rejects.toThrow(/mutually exclusive/);
    await expect(normalizeOfficeInput(bytes, 'docx', { doc: { source } }))
      .rejects.toThrow(/document session API/);
  });

  it('checks family, admission limit and cancellation before returning a source', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'WordDocument']);
    await expect(resolveDocDocumentInput(buildCfbFixture(['Root Entry', 'Workbook']), {
      doc: { source },
    })).rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(resolveDocDocumentInput(bytes, { doc: { source, maxInputBytes: 1 } }))
      .rejects.toMatchObject({ reason: 'source-too-large' });
    await expect(resolveDocDocumentInput(bytes, { doc: { source, maxInputBytes: NaN } }))
      .rejects.toThrow(RangeError);
    const controller = new AbortController(); controller.abort();
    await expect(resolveDocDocumentInput(bytes, { doc: { source, signal: controller.signal } }))
      .rejects.toMatchObject({ reason: 'aborted' });
  });
});
