import { describe, it, expect } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

/**
 * `XlsxWorkbook.load` rejects a password-protected / legacy-binary (CFB) file
 * with a typed OoxmlError *before* constructing the parser worker — so passing
 * synthetic CFB bytes here never opens a real Worker. See
 * `assertNotCfbContainer` in core.
 */
describe('XlsxWorkbook.load — CFB guard', () => {
  it('rejects an encrypted OOXML container with code "encrypted"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'EncryptionInfo', 'EncryptedPackage']);
    await expect(XlsxWorkbook.load(cfb)).rejects.toBeInstanceOf(OoxmlError);
    await expect(XlsxWorkbook.load(cfb)).rejects.toMatchObject({ code: 'encrypted' });
  });

  it('rejects a legacy .xls (Workbook stream) with code "legacy-binary-format"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'Workbook']);
    await expect(XlsxWorkbook.load(cfb)).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('rejects the older Excel "Book" stream with code "legacy-binary-format"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'Book']);
    await expect(XlsxWorkbook.load(cfb)).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('rejects an unrecognised CFB with code "not-ooxml"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'Whatever']);
    await expect(XlsxWorkbook.load(cfb)).rejects.toMatchObject({ code: 'not-ooxml' });
  });
});
