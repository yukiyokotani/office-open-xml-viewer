import { describe, expect, it, vi } from 'vitest';
import { OoxmlError, type LegacyOfficeConverter } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { materializeDocxDocument } from './docx.ts';
import { openXlsxWorkbook } from './xlsx.ts';
import { openPptxPresentation } from './pptx.ts';

describe('Node legacy Office normalization', () => {
  it.each([
    ['doc', 'docx', () => materializeDocxDocument(
      buildCfbFixture(['Root Entry', 'WordDocument']),
    )],
    ['xls', 'xlsx', () => openXlsxWorkbook(
      buildCfbFixture(['Root Entry', 'Workbook']),
    )],
    ['ppt', 'pptx', () => openPptxPresentation(
      buildCfbFixture(['Root Entry', 'PowerPoint Document']),
    )],
  ] as const)('retains the typed legacy rejection for %s without an opt-in converter', async (
    _from,
    _to,
    open,
  ) => {
    await expect(open()).rejects.toBeInstanceOf(OoxmlError);
    await expect(open()).rejects.toMatchObject({ code: 'legacy-binary-format' });
  });

  it.each([
    ['doc', 'docx', (convert: LegacyOfficeConverter['convert']) => materializeDocxDocument(
      buildCfbFixture(['Root Entry', 'WordDocument']),
      { legacyConversion: { doc: { converter: { convert } } } },
    )],
    ['xls', 'xlsx', (convert: LegacyOfficeConverter['convert']) => openXlsxWorkbook(
      buildCfbFixture(['Root Entry', 'Workbook']),
      { legacyConversion: { xls: { converter: { convert } } } },
    )],
    ['ppt', 'pptx', (convert: LegacyOfficeConverter['convert']) => openPptxPresentation(
      buildCfbFixture(['Root Entry', 'PowerPoint Document']),
      { legacyConversion: { ppt: { converter: { convert } } } },
    )],
  ] as const)('runs the opted-in %s -> %s converter before loading parser WASM', async (
    from,
    to,
    open,
  ) => {
    const convert = vi.fn<LegacyOfficeConverter['convert']>(
      async () => ({ bytes: new Uint8Array([0x50, 0x4b]) }),
    );

    await expect(open(convert)).rejects.toMatchObject({
      code: 'legacy-office-conversion',
      reason: 'invalid-output',
      from,
      to,
    });
    expect(convert).toHaveBeenCalledWith(expect.objectContaining({ from, to }));
  });
});
