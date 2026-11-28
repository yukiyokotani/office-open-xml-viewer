import { describe, expect, it } from 'vitest';
import { documentRenderCapabilities } from './vertical-render-capability.js';
import type { DocxDocumentModel } from './types.js';

const model = (value: unknown) => value as DocxDocumentModel;
const documentCarriesRunLanguage = (document: DocxDocumentModel) =>
  documentRenderCapabilities(document).carriesRunLanguage;

describe('run-language capability', () => {
  it('scans a very wide untagged model without exhausting the stack', () => {
    const body = Array.from({ length: 150_000 }, () => ({ type: 'paragraph', runs: [{ type: 'text', text: 'a' }] }));
    expect(documentCarriesRunLanguage(model({ body }))).toBe(false);
    body[149_999]!.runs.push({ type: 'text', text: 'b', lang: 'sr-rs' } as never);
    expect(documentCarriesRunLanguage(model({ body }))).toBe(true);
  });

  it('ignores x-none and ruby guide tags (guides are shaped untagged)', () => {
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ lang: 'x-none' }] }] }))).toBe(false);
    // A tag shaping would not admit (extlang, ill-formed) names no language here either.
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ lang: 'zh-cmn-Hans-CN' }, { langEastAsia: 'en-a' }] }] })))
      .toBe(false);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', text: '漢',
      typographyInput: { ruby: { language: { status: 'valid', value: 'ja-JP' }, guideRuns: [{ language: 'ja-jp' }] } },
    }] }] }))).toBe(false);
  });
});
