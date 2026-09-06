import { describe, expect, it } from 'vitest';
import {
  docxRenderedTextUsages,
  docxResolvedFontMetricCandidates,
} from './document-content.js';
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
      {
        text: 'result',
        fontFamilies: ['Latin Face', 'HANSI Face', 'EA Face'],
        latinFontFamilies: ['Latin Face', 'HANSI Face'],
        eastAsianFontFamilies: ['EA Face'],
        bold: true,
        italic: false,
      },
      { text: 'result', fontFamilies: ['CS Face'], bold: false, italic: true },
    ]);
  });

  it('enrolls only regular faces that win a rendered Latin or East-Asian slot', () => {
    const doc = {
      body: [{ type: 'paragraph', runs: [
        {
          type: 'text', text: 'Latin', fontFamily: 'EA Latin Face',
          fontFamilyEastAsia: 'Unused EA Default', bold: false, italic: false,
        },
        {
          type: 'text', text: '国 語𠀀', fontFamily: 'Latin Face',
          fontFamilyEastAsia: 'CJK Route', bold: false, italic: false,
        },
        {
          type: 'text', text: '国', fontFamilyEastAsia: 'Bold CJK',
          bold: true, italic: false,
        },
        {
          type: 'text', text: '𠀀', fontFamilyEastAsia: 'Supplementary CJK',
          bold: false, italic: false,
        },
        {
          type: 'text', text: '×', fontFamily: 'Punctuation Face',
          bold: false, italic: false,
        },
      ] }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;

    expect(docxResolvedFontMetricCandidates(doc, {
      'EA Latin Face': '80',
      'Latin Face': '80',
      'Punctuation Face': '80',
      'Unused EA Default': '80',
    })).toEqual([
      { family: 'EA Latin Face', probeText: '国', appliesToLatin: true },
      { family: 'CJK Route', probeText: '国', appliesToLatin: false },
      { family: 'Supplementary CJK', probeText: '𠀀', appliesToLatin: false },
    ]);
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
});
