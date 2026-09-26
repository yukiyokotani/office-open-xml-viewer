import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { buildPptFixture, buildXlsFixture } from '../test-fixtures.js';
import { testPptSource, testXlsSource } from '../test-sources.js';
import {
  materializeDocxDocument,
  materializePptxPresentation,
  materializeXlsxWorkbook,
  openPptxPresentation,
  openXlsxWorkbook,
} from './node-facade.js';

const cfb = (...streams: string[]) => new Uint8Array(buildCfbFixture(['Root Entry', ...streams]));
const legacyRejection = { code: 'legacy-binary-format' };

describe('Node openers and legacy model sources', () => {
  it('keep the typed legacy rejection for every binary family without model sources', async () => {
    for (const open of [
      () => materializeDocxDocument(cfb('WordDocument')),
      () => openXlsxWorkbook(cfb('Workbook')),
      () => openPptxPresentation(cfb('PowerPoint Document')),
    ]) {
      await expect(open()).rejects.toBeInstanceOf(OoxmlError);
      await expect(open()).rejects.toMatchObject(legacyRejection);
    }
  });

  it('leave foreign-family and ambiguous containers to the OOXML rejection', async () => {
    const xls = { modelSources: [testXlsSource()] };
    const ppt = { modelSources: [testPptSource()] };
    await expect(openXlsxWorkbook(cfb('PowerPoint Document'), xls)).rejects.toMatchObject(legacyRejection);
    await expect(openXlsxWorkbook(cfb('Workbook', 'WordDocument'), xls)).rejects.toMatchObject(legacyRejection);
    await expect(openPptxPresentation(cfb('WordDocument'), ppt)).rejects.toMatchObject(legacyRejection);
    // An encrypted package is never claimed; the OOXML path reports it.
    await expect(openXlsxWorkbook(cfb('EncryptionInfo', 'EncryptedPackage', 'Workbook'), xls))
      .rejects.toMatchObject({ code: 'encrypted' });
    // A source for another target is a configuration error, not a fallback.
    await expect(materializeDocxDocument(cfb('WordDocument'), xls as never)).rejects.toThrow(TypeError);
  });

  it('route each own family to its direct reader', async () => {
    const workbook = await materializeXlsxWorkbook(buildXlsFixture(), { modelSources: [testXlsSource()] });
    expect(JSON.stringify(workbook.workbookIndex.workbook)).toContain('表計算');
    expect(workbook.worksheets).toHaveLength(1);
    const cells = JSON.stringify([workbook.workbookIndex.sharedStrings, workbook.worksheets[0]]);
    expect(cells).toContain('42.5');
    expect(cells).toContain('日本語');

    // The direct PPT reader does not project a slide with outline text but no
    // drawing; reaching its fail-closed rejection proves the routing.
    await expect(materializePptxPresentation(buildPptFixture(), { modelSources: [testPptSource()] }))
      .rejects.toThrow(/UNSUPPORTED:.*no drawing/);
  });

  it('reject an oversize claimed input before opening it', async () => {
    const bytes = buildXlsFixture();
    await expect(materializeXlsxWorkbook(bytes, { modelSources: [testXlsSource({ maxInputBytes: bytes.byteLength - 1 })] }))
      .rejects.toThrow(RangeError);
  });
});
