import { describe, it, expect } from 'vitest';
import { DocxDocument } from './document';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

/**
 * `DocxDocument.load` rejects a password-protected / legacy-binary (CFB) file
 * with a typed OoxmlError *before* constructing the parser worker — so passing
 * synthetic CFB bytes here never opens a real Worker (which is why this runs in
 * a plain node test). See `assertNotCfbContainer` in core.
 */
describe('DocxDocument.load — CFB guard', () => {
  it('rejects an encrypted OOXML container with code "encrypted"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'EncryptionInfo', 'EncryptedPackage']);
    await expect(DocxDocument.load(cfb)).rejects.toBeInstanceOf(OoxmlError);
    await expect(DocxDocument.load(cfb)).rejects.toMatchObject({ code: 'encrypted' });
  });

  it('rejects a legacy .doc (WordDocument stream) with code "legacy-binary-format"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'WordDocument', '1Table']);
    await expect(DocxDocument.load(cfb)).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('rejects an unrecognised CFB with code "not-ooxml"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'Whatever']);
    await expect(DocxDocument.load(cfb)).rejects.toMatchObject({ code: 'not-ooxml' });
  });
});
