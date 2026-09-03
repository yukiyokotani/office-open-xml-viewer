import { describe, it, expect, vi } from 'vitest';
import { PptxPresentation } from './presentation';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

/**
 * `PptxPresentation.load` rejects a password-protected / legacy-binary (CFB)
 * file with a typed OoxmlError *before* constructing the parser worker — so
 * passing synthetic CFB bytes here never opens a real Worker. See
 * `assertNotCfbContainer` in core.
 */
describe('PptxPresentation.load — CFB guard', () => {
  it('rejects an encrypted OOXML container with code "encrypted"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'EncryptionInfo', 'EncryptedPackage']);
    await expect(PptxPresentation.load(cfb)).rejects.toBeInstanceOf(OoxmlError);
    await expect(PptxPresentation.load(cfb)).rejects.toMatchObject({ code: 'encrypted' });
  });

  it('rejects a legacy .ppt (PowerPoint Document stream) with code "legacy-binary-format"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'PowerPoint Document']);
    await expect(PptxPresentation.load(cfb)).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('routes an opted-in legacy .ppt through the PPT -> PPTX converter contract', async () => {
    const convert = vi.fn(async () => ({ bytes: new Uint8Array([0x50, 0x4b]) }));
    const cfb = buildCfbFixture(['Root Entry', 'PowerPoint Document']);

    await expect(PptxPresentation.load(cfb, {
      legacyConversion: { converter: { convert } },
    })).rejects.toMatchObject({
      code: 'legacy-office-conversion',
      reason: 'invalid-output',
      from: 'ppt',
      to: 'pptx',
    });
    expect(convert).toHaveBeenCalledWith(expect.objectContaining({ from: 'ppt', to: 'pptx' }));
  });

  it('rejects an unrecognised CFB with code "not-ooxml"', async () => {
    const cfb = buildCfbFixture(['Root Entry', 'Whatever']);
    await expect(PptxPresentation.load(cfb)).rejects.toMatchObject({ code: 'not-ooxml' });
  });
});
