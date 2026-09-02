import { describe, expect, it } from 'vitest';
import { docxFontPreloadNames, docxFontProviderNames } from './font-plan.js';
import type { DocxDocumentModel } from './types.js';

function docWith(text: string, major = 'Calibri', minor = 'Calibri'): DocxDocumentModel {
  return {
    section: {} as DocxDocumentModel['section'],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    majorFont: major,
    minorFont: minor,
    body: [{
      type: 'paragraph',
      runs: [{ type: 'text', text } as never],
    } as never],
  } as DocxDocumentModel;
}

describe('DOCX font plan', () => {
  it('keeps a Latin document to its authored theme fonts', () => {
    const names = docxFontPreloadNames(docWith('Hello, world.'));
    expect(names).toEqual(['Calibri', 'Calibri']);
    expect(names).not.toContain('Noto Sans JP');
  });

  it('adds script fallbacks only to the Google preload plan', () => {
    const doc = docWith('日本語', 'Calibri', 'Ubuntu');
    expect(docxFontPreloadNames(doc)).toEqual(expect.arrayContaining([
      'Calibri', 'Ubuntu', 'Noto Sans JP', 'Noto Serif JP',
    ]));
    expect(docxFontProviderNames(doc)).toEqual(expect.arrayContaining(['Calibri', 'Ubuntu']));
    expect(docxFontProviderNames(doc)).not.toContain('Noto Sans JP');
  });

  it('uses the theme CJK family as the language hint', () => {
    const names = docxFontPreloadNames(docWith('漢字', 'Malgun Gothic', 'Malgun Gothic'));
    expect(names).toContain('Noto Sans KR');
    expect(names).not.toContain('Noto Sans JP');
  });
});
