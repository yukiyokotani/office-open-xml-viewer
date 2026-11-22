import { describe, expect, it } from 'vitest';
import { documentCarriesRunLanguage } from './run-language-capability.js';
import type { DocxDocumentModel } from './types.js';

const model = (value: unknown) => value as DocxDocumentModel;

describe('documentCarriesRunLanguage', () => {
  it('scans a very wide untagged model without exhausting the stack', () => {
    const body = Array.from({ length: 150_000 }, () => ({ type: 'paragraph', runs: [{ type: 'text', text: 'a' }] }));
    expect(documentCarriesRunLanguage(model({ body }))).toBe(false);
    body[149_999]!.runs.push({ type: 'text', text: 'b', lang: 'sr-rs' } as never);
    expect(documentCarriesRunLanguage(model({ body }))).toBe(true);
  });

  it('finds a language authored only on a ruby guide, and ignores x-none', () => {
    const ruby = (language: unknown, guideLanguage: string | null, axes: Record<string, string | null> = {}) =>
      model({ body: [{ runs: [{
        type: 'text', text: '漢',
        typographyInput: { ruby: { language, guideRuns: [{
          text: 'かん', language: guideLanguage, eastAsiaLanguage: null, bidiLanguage: null, ...axes,
        }] } },
      }] }] });
    expect(documentCarriesRunLanguage(ruby({ status: 'valid', value: 'ja-JP' }, null))).toBe(true);
    expect(documentCarriesRunLanguage(ruby({ status: 'absent', value: null }, 'ja-jp'))).toBe(true);
    expect(documentCarriesRunLanguage(ruby({ status: 'absent', value: null }, null))).toBe(false);
    // A guide's East Asian or bidi language alone is a run language too.
    expect(documentCarriesRunLanguage(ruby({ status: 'absent', value: null }, null, { eastAsiaLanguage: 'ja-jp' })))
      .toBe(true);
    expect(documentCarriesRunLanguage(ruby({ status: 'absent', value: null }, null, { bidiLanguage: 'ar-sa' })))
      .toBe(true);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ lang: 'x-none' }] }] }))).toBe(false);
  });
});
