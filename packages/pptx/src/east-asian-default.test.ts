import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import {
  complexScriptDefaultFace,
  eastAsianDefaultFaces,
  isEastAsianFace,
  isSerifLatinFace,
} from './east-asian-default.js';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

// Issue #1627 controls D (PowerPoint reference PDF engine): a run whose East
// Asian or complex-script slot has no face.
describe('PowerPoint application default faces', () => {
  it('classifies Latin faces by their installed PANOSE serif style', () => {
    expect(isSerifLatinFace('Perpetua')).toBe(true);
    expect(isSerifLatinFace('Georgia')).toBe(true);
    expect(isSerifLatinFace('Garamond')).toBe(true);
    for (const face of ['Corbel', 'Arial', 'Tahoma', 'Gill Sans MT', 'No Such Face']) {
      expect(isSerifLatinFace(face)).toBe(false);
    }
    expect(isEastAsianFace('Yu Gothic')).toBe(true);
    expect(isEastAsianFace('Tahoma')).toBe(false);
  });

  it('uses the first tier face that covers the whole East Asian text', () => {
    expect(eastAsianDefaultFaces('Perpetua', '日本語かな')[0]).toBe('MS Mincho');
    expect(eastAsianDefaultFaces('Perpetua', '繁體中文')[0]).toBe('MS Mincho');
    expect(eastAsianDefaultFaces('Perpetua', '简体中文')[0]).toBe('PMingLiU');
    expect(eastAsianDefaultFaces('Perpetua', '한국어')[0]).toBe('Batang');
    expect(eastAsianDefaultFaces('Corbel', '日本語かな')[0]).toBe('MS Gothic');
    expect(eastAsianDefaultFaces('Corbel', '简体中文')[0]).toBe('Microsoft JhengHei');
    expect(eastAsianDefaultFaces('Corbel', '한국어')[0]).toBe('Malgun Gothic');
    expect(eastAsianDefaultFaces('No Such Face', '日本語かな')[0]).toBe('MS Gothic');
  });

  it('draws East Asian text in a Latin face that is itself East Asian', () => {
    expect(eastAsianDefaultFaces('Yu Gothic', '日本語かな'))
      .toEqual(['Yu Gothic', 'Microsoft JhengHei', 'Malgun Gothic']);
    expect(eastAsianDefaultFaces('SimSun', '한국어')).toEqual(['SimSun', 'PMingLiU', 'Batang']);
  });

  it('picks the complex-script default per character script', () => {
    expect(complexScriptDefaultFace('ע')).toBe('Arial');
    expect(complexScriptDefaultFace('ع')).toBe('Arial');
    expect(complexScriptDefaultFace('ภ')).toBe('Angsana New');
    expect(complexScriptDefaultFace('ह')).toBe('Mangal');
    expect(complexScriptDefaultFace('த')).toBe('Latha');
  });
});

function measuringContext(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true, runs,
  } as Paragraph;
}

function run(text: string, extra: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Corbel', ...extra,
  };
}

describe('layoutParagraph East Asian / complex-script faces', () => {
  const fonts = (runs: TextRunData[]) => layoutParagraph(measuringContext(), paragraph(runs), 2000, 20, '000000', 1, 0)
    .flatMap((line) => line.segments.map((s) => [s.text, s.font] as const));

  it('offers an empty East Asian slot to the selected face, then the application defaults', () => {
    // #1689: with no cs face the Latin face is selected and draws what it
    // covers; Corbel maps no CJK, so the glyphs pass the symbol fallback
    // (which maps no CJK either) to the #1627 sans tier.
    const [latin, ea] = fonts([run('Hxg 日本語')]);
    expect(latin[1]).toMatch(/^\d+px "Corbel"/u);
    expect(ea[0]).toBe('日本語');
    expect(ea[1]).toMatch(/^\d+px "Corbel", "Calibri", "Cambria Math", "MS Gothic", "Microsoft JhengHei", "Malgun Gothic"/u);
  });

  it('keeps a resolved East Asian face and splits complex scripts by script default', () => {
    const segments = fonts([run('日本 עב ไท', { fontFamilyEa: 'Yu Mincho' })]);
    expect(segments.find(([text]) => text.includes('日本'))?.[1]).toMatch(/^\d+px "Yu Mincho"/u);
    expect(segments.find(([text]) => text.includes('עב'))?.[1]).toMatch(/^\d+px "Arial"/u);
    expect(segments.find(([text]) => text.includes('ไท'))?.[1]).toMatch(/^\d+px "Angsana New"/u);
  });
});
