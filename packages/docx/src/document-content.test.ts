import { describe, expect, it } from 'vitest';
import { docxRenderedTextUsages } from './document-content.js';
import { docxFontDemandCandidates } from './font-demand-plan.js';
import { docxOfficeFontFallbackRequests } from './font-plan.js';
import type { InternalFieldRun } from './parser-model.js';
import type { DocxDocumentModel } from './types.js';

describe('docx rendered text inventory', () => {
  // The families demand would request: authored, then script fallbacks.
  const docxFontPreloadNames = (doc: DocxDocumentModel) => {
    const candidates = docxFontDemandCandidates(doc, []);
    return [...candidates.authored, ...candidates.fallback];
  };
  const documentWithRun = (run: object) => ({
    body: [{ type: 'paragraph', runs: [run] }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  }) as unknown as DocxDocumentModel;

  it('demands a face authored only on a ruby guide run', () => {
    const fonts = { ascii: 'GuideOnly', highAnsi: 'GuideOnly', eastAsia: 'GuideOnly', complexScript: null };
    const doc = documentWithRun({
      type: 'text', text: '漢', fontSize: 10, fontFamily: 'BaseFace',
      ruby: { text: 'かん', fontSizePt: 5 },
      __typographyAcquisition: { ruby: { guideRuns: [{
        text: 'かん', fontFamily: 'GuideOnly', fonts, fontSizePt: 5, bold: false, italic: false, color: null, language: null,
      }] } },
    });
    for (const collection of ['font-demand', 'native-preflight'] as const) {
      expect([...docxRenderedTextUsages(doc, collection)].flatMap((usage) => usage.fontFamilies)).toContain('GuideOnly');
    }
    expect(docxFontPreloadNames(doc)).toContain('GuideOnly');
  });

  it('demands a ruby guide run\'s complex-script face with its complex-script bold and italic', () => {
    // §17.3.2.2 / §17.3.2.17: the guide's complex-script text is shaped with
    // bCs and iCs, so that is the variant its face is demanded in.
    const fonts = { ascii: null, highAnsi: null, eastAsia: null, complexScript: 'GuideComplex' };
    const doc = documentWithRun({
      type: 'text', text: '漢', fontSize: 10, fontFamily: 'BaseFace',
      ruby: { text: 'سلام', fontSizePt: 5 },
      __typographyAcquisition: { ruby: { guideRuns: [{
        text: 'سلام', fontFamily: null, fonts, fontSizePt: 5, bold: false, italic: false,
        boldCs: true, italicCs: true, rtl: true, cs: false, fontHint: null, color: null, language: null,
      }] } },
    });
    for (const collection of ['font-demand', 'native-preflight'] as const) {
      const usages = [...docxRenderedTextUsages(doc, collection)].filter((usage) => usage.fontFamilies.includes('GuideComplex'));
      expect(usages.map((usage) => [usage.bold, usage.italic]), collection).toEqual([[true, true]]);
    }
  });

  it('demands a ruby guide run\'s face inside a text box story in every collection', () => {
    // The flattened text-box blocks carry no guide runs; the story does.
    const fonts = { ascii: null, highAnsi: null, eastAsia: null, complexScript: 'BoxGuide' };
    const ruby = {
      type: 'text', text: '漢', fontSize: 10, fontFamily: 'BaseFace',
      ruby: { text: 'سلام', fontSizePt: 5 },
      __typographyAcquisition: { ruby: { guideRuns: [{
        text: 'سلام', fontFamily: null, fonts, fontSizePt: 5, bold: false, italic: false,
        boldCs: true, italicCs: true, rtl: true, cs: false, fontHint: null, color: null, language: null,
      }] } },
    };
    const doc = documentWithRun({
      type: 'shape', inline: true,
      textBlocks: [{ text: '漢', fontFamily: 'BaseFace', runs: [{ text: '漢', fontFamily: 'BaseFace' }] }],
      textBoxContent: [{ type: 'paragraph', runs: [ruby] }],
    });
    for (const collection of ['font-demand', 'native-preflight'] as const) {
      const usages = [...docxRenderedTextUsages(doc, collection)].filter((usage) => usage.fontFamilies.includes('BoxGuide'));
      expect(usages.map((usage) => [usage.bold, usage.italic]), collection).toEqual([[true, true]]);
    }
  });

  it('demands the complex-script face a ruby guide run inherits from its base, in the guide\'s variant', () => {
    // The guide run names no faces: it is shaped in the base run's
    // complex-script face, with its own bCs.
    const none = { ascii: null, highAnsi: null, eastAsia: null, complexScript: null };
    const doc = documentWithRun({
      type: 'text', text: '漢', fontSize: 10, fontFamily: 'BaseFace', fontFamilyCs: 'BaseComplex',
      ruby: { text: 'سلام', fontSizePt: 5 },
      __typographyAcquisition: { ruby: { guideRuns: [{
        text: 'سلام', fontFamily: null, fonts: none, fontSizePt: 5, bold: false, italic: false,
        boldCs: true, italicCs: false, rtl: true, cs: false, fontHint: null, color: null, language: null,
      }] } },
    });
    for (const collection of ['font-demand', 'native-preflight'] as const) {
      const usages = [...docxRenderedTextUsages(doc, collection)]
        .filter((usage) => usage.rubyGuide && usage.fontFamilies.includes('BaseComplex'));
      expect(usages.map((usage) => [usage.bold, usage.italic]), collection).toEqual([[true, false]]);
    }
  });

  it('preloads an inline drawing run face even when its text box has no text', () => {
    const doc = documentWithRun({ type: 'shape', inline: true, textBlocks: [],
      __inlineObjectFontFacts: { fontSize: 30, fontFamily: 'Courier New', bold: true } });
    expect([...docxRenderedTextUsages(doc, 'font-demand')]
      .flatMap((usage) => usage.fontFamilies)).toContain('Courier New');
    expect(docxOfficeFontFallbackRequests(doc, {}))
      .toContainEqual({ family: 'Courier New', weight: 700, style: 'normal', format: 'docx' });
  });

  it.each(['rtl', 'cs'] as const)('collects the independent CS tuple of a compatibility %s run', (flag) => {
    const facts = {
      text: '123', fontSizePt: 10, [flag]: true,
      fontFamily: 'Arial', bold: true, italic: true,
      fontFamilyCs: 'Cambria', boldCs: false, italicCs: false,
    };
    const shapeDoc = documentWithRun({
      type: 'shape', textBlocks: [{ text: '123', fontSizePt: 10, runs: [facts] }],
    });
    const bodyDoc = documentWithRun({ type: 'text', ...facts });

    expect(docxFontPreloadNames(shapeDoc)).toContain('Cambria');
    expect(docxFontPreloadNames(shapeDoc)).toEqual(docxFontPreloadNames(bodyDoc));
    for (const options of [{}, { fontDemand: true }]) {
      const requests = docxOfficeFontFallbackRequests(shapeDoc, options);
      expect(requests.filter((request) => request.family === 'Cambria')).toEqual([
        { family: 'Cambria', weight: 400, style: 'normal', format: 'docx' },
        { family: 'Cambria', presenceOnly: true },
      ]);
      expect(requests).toContainEqual({ family: 'Arial', weight: 700, style: 'italic', format: 'docx' });
      expect(requests).toEqual(docxOfficeFontFallbackRequests(bodyDoc, options));
    }
  });

  it.each([
    { companion: 'family', family: 'Cambria', csFacts: { boldCs: false, italicCs: false }, boldCs: false, italicCs: false },
    { companion: 'style', family: 'Arial', csFacts: { fontFamilyCs: 'Cambria' }, boldCs: true, italicCs: true },
  ])('collects an omitted CS $companion from inherited compatibility Latin facts', ({ family, csFacts, boldCs, italicCs }) => {
    const shapeDoc = documentWithRun({
      type: 'shape', textBlocks: [{
        text: '123', fontSizePt: 10, fontFamily: family, bold: true, italic: true,
        runs: [{ text: '123', fontSizePt: 10, rtl: true, ...csFacts }],
      }],
    });
    const bodyDoc = documentWithRun({
      type: 'text', text: '123', rtl: true,
      fontFamily: family, bold: true, italic: true,
      fontFamilyCs: 'Cambria', boldCs, italicCs,
    });

    expect(docxFontPreloadNames(shapeDoc)).toEqual(docxFontPreloadNames(bodyDoc));
    for (const options of [{}, { fontDemand: true }]) {
      expect(docxOfficeFontFallbackRequests(shapeDoc, options))
        .toEqual(docxOfficeFontFallbackRequests(bodyDoc, options));
    }
  });

  it('keeps the compatibility inventory unchanged without rtl/cs', () => {
    const doc = documentWithRun({
      type: 'shape', textBlocks: [{ text: '123', fontSizePt: 10, runs: [{
        text: '123', fontSizePt: 10, rtl: false, cs: false,
        fontFamily: 'Arial', bold: true, italic: true,
        fontFamilyCs: 'Cambria', boldCs: false, italicCs: false,
      }] }],
    });

    expect(docxFontPreloadNames(doc)).not.toContain('Cambria');
    for (const options of [{}, { fontDemand: true }]) {
      const requests = docxOfficeFontFallbackRequests(doc, options);
      expect(requests.some((request) => request.family === 'Cambria')).toBe(false);
      expect(requests).toContainEqual({ family: 'Arial', weight: 700, style: 'italic', format: 'docx' });
    }
  });

  it('inventories field results on both non-CS/EA and CS formatting tuples', () => {
    const field: InternalFieldRun & { type: 'field' } = {
      type: 'field', fieldType: 'other', instruction: 'REF x', fallbackText: 'result',
      bold: true, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Latin Face', background: null,
      vertAlign: null, fontFamilyHighAnsi: 'HANSI Face', fontFamilyEastAsia: 'EA Face', fontFamilyCs: 'CS Face',
      boldCs: false, italicCs: true,
    };
    const doc = {
      body: [{ type: 'paragraph', runs: [field] }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;

    expect([...docxRenderedTextUsages(doc)].filter((usage) => usage.text === 'result')).toEqual([
      expect.objectContaining({
        text: 'result',
        fontFamilies: ['Latin Face', 'HANSI Face', 'EA Face'],
        bold: true, italic: false,
      }),
      expect.objectContaining({
        text: 'result', fontFamilies: ['CS Face'], bold: false, italic: true,
      }),
    ]);
    expect(docxOfficeFontFallbackRequests(doc)).toContainEqual({
      family: 'CS Face', weight: 400, style: 'italic', format: 'docx',
    });
  });

  it('inventories an ordinary text face authored only on the hAnsi axis', () => {
    const doc = {
      body: [{ type: 'paragraph', runs: [{
        type: 'text', text: 'é', fontFamily: null, fontFamilyHighAnsi: 'HANSI Only',
        bold: false, italic: false,
      }] }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;

    expect([...docxRenderedTextUsages(doc)].find((usage) => usage.text === 'é')?.fontFamilies)
      .toContain('HANSI Only');
  });

  it('requests the complex-script tuple painted by an ordinary text run', () => {
    const doc = {
      body: [{ type: 'paragraph', runs: [{
        type: 'text', text: 'مرحبا', fontFamily: 'Latin Face', fontFamilyCs: 'CS Face',
        bold: false, italic: true, boldCs: true, italicCs: false,
      }] }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;

    expect([...docxRenderedTextUsages(doc)].filter((usage) => usage.text === 'مرحبا'))
      .toEqual([
        expect.objectContaining({
          fontFamilies: expect.arrayContaining(['Latin Face']), bold: false, italic: true,
        }),
        expect.objectContaining({ fontFamilies: ['CS Face'], bold: true, italic: false }),
      ]);
    expect(docxOfficeFontFallbackRequests(doc)).toContainEqual({
      family: 'CS Face', weight: 700, style: 'normal', format: 'docx',
    });
  });
});
