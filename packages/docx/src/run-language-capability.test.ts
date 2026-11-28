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
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', lang: 'x-none' }] }] }))).toBe(false);
    // A tag shaping would not admit (extlang, ill-formed) names no language here either.
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', lang: 'zh-cmn-Hans-CN' }, { type: 'text', langEastAsia: 'en-a' }] }] })))
      .toBe(false);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', text: '漢',
      typographyInput: { ruby: { language: { status: 'valid', value: 'ja-JP' }, guideRuns: [{ language: 'ja-jp' }] } },
    }] }] }))).toBe(false);
  });

  it('does not read a property named lang outside a run as a run language', () => {
    expect(documentCarriesRunLanguage(model({ body: [], fontFamilyClasses: { lang: 'roman' } }))).toBe(false);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', text: 'a', langBidi: 'ar-sa' }] }] })))
      .toBe(true);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [], paragraphMarkFontFacts: { lang: 'sr-rs' } }] })))
      .toBe(true);
    expect(documentCarriesRunLanguage(model({ body: [{ runs: [], numbering: { fontFacts: { langEastAsia: 'ja-jp' } } }] })))
      .toBe(true);
  });

  it('applies the shaping rules to script-class precedence and automatic axes', () => {
    const only = (run: Record<string, unknown>) => documentCarriesRunLanguage(model({ body: [{ runs: [{ type: 'text', ...run }] }] }));
    expect(only({ langScriptClass: '', langEastAsia: 'ja-jp', lang: 'en-us' })).toBe(false);
    expect(only({ langScriptClass: 'ja-jp' })).toBe(true);
    expect(only({ langEastAsia: 'ja-jp', langEastAsiaAutomatic: true })).toBe(false);
    expect(only({ langBidi: 'ar-sa', langBidiAutomatic: true })).toBe(false);
    expect(only({ langEastAsia: 'ja-jp', langEastAsiaAutomatic: true, lang: 'en-us' })).toBe(true);
    expect(only({ langEastAsia: 'x-none', langBidi: 'zh-cmn-hans' })).toBe(false);
    const facts = (value: Record<string, unknown>) =>
      documentCarriesRunLanguage(model({ body: [{ runs: [], paragraphMarkFontFacts: value }] }));
    expect(facts({ langScriptClass: '', langEastAsia: 'ja-jp' })).toBe(false);
    expect(facts({ langBidi: 'ar-sa', langBidiAutomatic: true })).toBe(false);
    expect(facts({ langBidi: 'ar-sa' })).toBe(true);
  });

  it('reads language properties only on runs and run facts', () => {
    let reads = 0;
    const counted = () => Object.defineProperty({ type: 'paragraph', runs: [{ type: 'image' }] }, 'lang', {
      enumerable: true, get: () => { reads += 1; return 'sr-rs'; },
    });
    const body = Array.from({ length: 100 }, counted);
    expect(documentCarriesRunLanguage(model({ body, fontFamilyClasses: { lang: 'roman' } }))).toBe(false);
    // Only the walk's own enumeration touches the value, never the predicate's keyed reads.
    expect(reads).toBe(100);
  });
});
