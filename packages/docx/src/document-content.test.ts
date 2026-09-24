import { describe, expect, it } from 'vitest';
import { docxRenderedTextUsages } from './document-content.js';
import { docxOfficeFontFallbackRequests } from './google-fonts.js';
import type { InternalFieldRun } from './parser-model.js';
import type { DocxDocumentModel } from './types.js';

describe('docx rendered text inventory', () => {
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
