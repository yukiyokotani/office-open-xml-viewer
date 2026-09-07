import { describe, expect, it, vi } from 'vitest';
import { buildCfbFixture, buildStoredZip } from '../testing';
import { resolveXlsWorkbookInput } from './legacy-office-conversion.js';
import { normalizeOfficeInput } from '../conversion/legacy-office.js';

const source = {
  protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
  wasmUrl: 'https://example.test/xls.wasm',
} as const;

describe('direct XLS input selection', () => {
  it('selects only explicitly enabled XLS without converter or decoder loading', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'Workbook']);
    expect(await resolveXlsWorkbookInput(bytes, { xls: { source } }))
      .toMatchObject({ kind: 'legacy-xls', source, bytes });
    await expect(resolveXlsWorkbookInput(bytes)).rejects.toThrow();
    await expect(resolveXlsWorkbookInput(bytes, { ppt: { converter: { convert: vi.fn() } } }))
      .rejects.toThrow();
  });

  it('leaves OOXML on the existing path', async () => {
    const bytes = buildStoredZip({
      '[Content_Types].xml': '<Types/>', 'xl/workbook.xml': '<workbook/>',
    });
    expect(await resolveXlsWorkbookInput(bytes, { xls: { source } }))
      .toEqual({ kind: 'ooxml', bytes });
  });

  it('rejects ambiguous configuration and refuses the byte-normalization API', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'Workbook']);
    await expect(resolveXlsWorkbookInput(bytes, {
      xls: { source, converter: { convert: vi.fn() } } as never,
    })).rejects.toThrow(/mutually exclusive/);
    await expect(normalizeOfficeInput(bytes, 'xlsx', { xls: { source } }))
      .rejects.toThrow(/workbook session API/);
  });

  it('checks source family, admission limit and cancellation before initialization', async () => {
    const bytes = buildCfbFixture(['Root Entry', 'Workbook']);
    await expect(resolveXlsWorkbookInput(buildCfbFixture(['Root Entry', 'WordDocument']), {
      xls: { source },
    })).rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(resolveXlsWorkbookInput(bytes, { xls: { source, maxInputBytes: 1 } }))
      .rejects.toMatchObject({ reason: 'source-too-large' });
    await expect(resolveXlsWorkbookInput(bytes, { xls: { source, maxInputBytes: NaN } }))
      .rejects.toThrow(RangeError);
    const controller = new AbortController();
    controller.abort();
    await expect(resolveXlsWorkbookInput(bytes, { xls: { source, signal: controller.signal } }))
      .rejects.toMatchObject({ reason: 'aborted' });
  });
});
