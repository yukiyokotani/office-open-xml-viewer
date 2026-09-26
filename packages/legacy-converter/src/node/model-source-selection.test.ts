import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { buildPptFixture } from '../test-fixtures.js';
import { testPptSource } from '../test-sources.js';
import {
  materializeDocxDocument,
  materializePptxPresentation,
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
    const ppt = { modelSources: [testPptSource()] };
    await expect(openPptxPresentation(cfb('WordDocument'), ppt)).rejects.toMatchObject(legacyRejection);
    await expect(openPptxPresentation(cfb('PowerPoint Document', 'Workbook'), ppt)).rejects.toMatchObject(legacyRejection);
    // An encrypted package is never claimed; the OOXML path reports it.
    await expect(openPptxPresentation(cfb('EncryptionInfo', 'EncryptedPackage', 'PowerPoint Document'), ppt))
      .rejects.toMatchObject({ code: 'encrypted' });
    // A source for another target is a configuration error, not a fallback.
    await expect(materializeDocxDocument(cfb('WordDocument'), ppt as never)).rejects.toThrow(TypeError);
  });

  it('route its own family to the direct reader', async () => {
    // The direct PPT reader does not project a slide with outline text but no
    // drawing; reaching its fail-closed rejection proves the routing.
    await expect(materializePptxPresentation(buildPptFixture(), { modelSources: [testPptSource()] }))
      .rejects.toThrow(/UNSUPPORTED:.*no drawing/);
  });

  it('reject an oversize claimed input before opening it', async () => {
    const bytes = buildPptFixture();
    await expect(materializePptxPresentation(bytes, { modelSources: [testPptSource({ maxInputBytes: bytes.byteLength - 1 })] }))
      .rejects.toThrow(RangeError);
  });
});
