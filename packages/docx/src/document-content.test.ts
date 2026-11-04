import { describe, expect, it } from 'vitest';
import { docxRenderedTextUsages } from './document-content.js';
import { docxFontPreloadNames, docxOfficeFontFallbackRequests } from './google-fonts.js';
import type { InternalFieldRun } from './parser-model.js';
import type { DocxDocumentModel } from './types.js';

describe('docx rendered text inventory', () => {
  const documentWithRun = (run: object) => ({
    body: [{ type: 'paragraph', runs: [run] }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  }) as unknown as DocxDocumentModel;

  it('preloads an inline drawing run face even when its text box has no text', () => {
    const doc = documentWithRun({ type: 'shape', inline: true, textBlocks: [],
      __inlineObjectFontFacts: { fontSize: 30, fontFamily: 'Courier New', bold: true } });
    expect([...docxRenderedTextUsages(doc, 'google-preload')]
      .flatMap((usage) => usage.fontFamilies)).toContain('Courier New');
    expect(docxOfficeFontFallbackRequests(doc, {}))
      .toContainEqual({ family: 'Courier New', weight: 700, style: 'normal' });
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
    for (const options of [{}, { useGoogleFonts: true }]) {
      const requests = docxOfficeFontFallbackRequests(shapeDoc, options);
      expect(requests.filter((request) => request.family === 'Cambria')).toEqual([
        { family: 'Cambria', weight: 400, style: 'normal' },
        { family: 'Cambria', presenceOnly: true },
      ]);
      expect(requests).toContainEqual({ family: 'Arial', weight: 700, style: 'italic' });
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
    for (const options of [{}, { useGoogleFonts: true }]) {
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
    for (const options of [{}, { useGoogleFonts: true }]) {
      const requests = docxOfficeFontFallbackRequests(doc, options);
      expect(requests.some((request) => request.family === 'Cambria')).toBe(false);
      expect(requests).toContainEqual({ family: 'Arial', weight: 700, style: 'italic' });
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
      family: 'CS Face', weight: 400, style: 'italic',
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
      family: 'CS Face', weight: 700, style: 'normal',
    });
  });
});
