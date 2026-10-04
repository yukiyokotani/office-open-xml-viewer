import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph, naturalWidthExceedsBbox, paragraphInputRuns, renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';

const SCALE = 1 / 12700;
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };
const FACES = { fontFamily: 'Corbel', fontFamilyEa: 'Meiryo UI', fontFamilyCs: 'Microsoft Sans Serif' };
function run(text: string, lang?: string, extra: Partial<TextRunData> = {}): TextRunData {
  return { type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', ...FACES, lang, ...extra };
}
function paragraph(runs: Paragraph['runs']): Paragraph {
  return { alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
    spaceLine: null, lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs };
}
function body(runs: Paragraph['runs'], vert = 'horz'): TextBody {
  return { paragraphs: [paragraph(runs)], verticalAnchor: 't', defaultFontSize: 20,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert, autoFit: 'none' };
}
function segments(runs: TextRunData[]) {
  return paragraphInputRuns(paragraph(runs), 20, '#000000', SCALE, false, false, 1, undefined, RC)
    .input.flatMap((item) => item.type === 'text' ? [{ text: item.text, face: item.style.faceFamily, font: item.style.font }] : []);
}
function context() {
  const calls: { text: string; font: string }[] = [];
  let font = '';
  const ctx = {
    canvas: { style: {} }, get font() { return font; }, set font(value: string) { font = value; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px', direction: 'ltr',
    textAlign: 'left', textBaseline: 'alphabetic',
    measureText(text: string) {
      const advance = font.includes('"Meiryo UI"') ? 20 : font.includes('"Microsoft Sans Serif"') ? 15 : 5;
      return { width: [...text].reduce((sum, ch) => sum + (ch === '¥' ? 20 : advance), 0),
        actualBoundingBoxAscent: 16, actualBoundingBoxDescent: 4,
        fontBoundingBoxAscent: 16, fontBoundingBoxDescent: 4 };
    },
    fillText(text: string) { calls.push({ text, font }); },
    save() {}, restore() {}, translate() {}, rotate() {}, scale() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {}, fillRect() {}, setLineDash() {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

describe('PPTX language-dependent font slots through the renderer', () => {
  it('routes measured punctuation and symbol boundaries independently of line-break classes', () => {
    const result = segments([run('A§¨°±´×÷⁇⁈⓾⓿─▟■☙♰⚀❟❨❶B', 'en-US')]);
    expect(result.map(({ text, face }) => [text, face])).toEqual([
      ['A', 'Corbel'], ['§¨°±´×÷⁇⁈⓾⓿', 'Meiryo UI'], ['─', 'Corbel'],
      ['▟■☙', 'Meiryo UI'], ['♰', 'Microsoft Sans Serif'], ['⚀❟❨', 'Corbel'], ['❶', 'Meiryo UI'], ['B', 'Corbel'],
    ]);
  });

  it('uses normative symbol boundaries and measured quote endpoints for arbitrary faces and sizes', () => {
    // Symbol cycles fail face independence; these are specification expectations,
    // not an attempt to infer slots from the substituted PDF faces.
    for (const [latin, ea, cs, size] of [['Calibri', 'Cambria', 'Arial', 10],
      ['Times New Roman', 'Arial', 'Calibri', 32]] as const) {
      const faces = { fontFamily: latin, fontFamilyEa: ea, fontFamilyCs: cs, fontSize: size };
      for (const lang of ['en-US', 'ja-JP']) {
        for (const [ch, face] of [['⓾', ea], ['⓿', ea], ['▟', ea], ['■', ea],
          ['☙', ea], ['♰', cs], ['♱', cs], ['„', lang === 'ja-JP' ? ea : latin], ['‟', latin]]) {
          const runs = [run(ch, lang, faces)];
          expect(segments(runs).map(({ text, face }) => [text, face])).toEqual([[ch, face]]);
          const { ctx, calls } = context();
          renderTextBody(ctx, body(runs), 0, 0, 300, 100, SCALE);
          expect(calls.find((c) => c.text === ch)?.font).toContain(`"${face}"`);
        }
        // Interior scalars without new contradictory evidence retain prior routing.
        expect(segments([run('─', lang, faces)])[0].face).toBe(latin);
      }
    }
  });

  it('uses the East Asian slot for seven curly quotes but keeps U+201F Latin', () => {
    for (const lang of ['ja-JP', 'ko-KR', 'zh-CN', 'zh-TW']) {
      expect(segments([run('‘’‚‛“”„‟', lang)]).map(({ text, face }) => [text, face]))
        .toEqual([['‘’‚‛“”„', 'Meiryo UI'], ['‟', 'Corbel']]);
    }
    expect(segments([run('‘’‚‛“”„‟', 'en-US')])[0].face).toBe('Corbel');
    expect(segments([run('“”', 'en-US', { altLang: 'ja-JP' })])[0].face).toBe('Corbel');
  });

  it('selects cs for all measured European digits', () => {
    for (const lang of ['he-IL', 'ar-SA', 'th-TH', 'hi-IN', 'fa-IR', 'ur-PK', 'yi-001', 'syr-SY', 'ug-CN', 'ar-EG', 'he', 'ur-IN']) {
      expect(segments([run('A0123456789B', lang)]).map(({ text, face }) => [text, face]))
        .toEqual([['A', 'Corbel'], ['0123456789', 'Microsoft Sans Serif'], ['B', 'Corbel']]);
    }
    expect(segments([run('0123456789', 'en-US', { altLang: 'he-IL' })])[0].face).toBe('Corbel');
  });

  it('resolves measured standalone and Latin-neighbour punctuation across run seams', () => {
    for (const lang of ['ar-EG', 'ar-SA', 'fa-IR', 'he', 'he-IL', 'hi-IN',
      'syr-SY', 'th-TH', 'ug-CN', 'ur-IN', 'ur-PK', 'yi-001']) {
      for (const ch of '»×÷‘’‚‛“”„⁇⁈') {
        expect(segments([run(ch, lang)]).map(({ text, face }) => [text, face]))
          .toEqual([[ch, 'Microsoft Sans Serif']]);
        for (const runs of [[run(`A${ch}B`, lang)], [run('A', 'en-US'), run(ch, lang), run('B', 'en-US')]]) {
          expect(segments(runs).map(({ text, face }) => [text, face]))
            .toEqual(runs.length === 1 ? [[`A${ch}B`, 'Corbel']] : [['A', 'Corbel'], [ch, 'Corbel'], ['B', 'Corbel']]);
        }
      }
    }
    // Native neighbours and multi-punctuation sequences remain evidence gaps.
    expect(segments([run('אב×אב', 'he-IL')]).map(({ text, face }) => [text, face]))
      .toEqual([['אב', 'Microsoft Sans Serif'], ['×', 'Corbel'], ['אב', 'Microsoft Sans Serif']]);
    expect(segments([run('×÷', 'fa-IR')])[0].face).toBe('Meiryo UI');
  });

  it('keeps nonseparable Myanmar marks with their base and resumes scalar routing afterward', () => {
    expect(segments([run('\u1000\ua9e0\uaa60', 'my-MM')]).map(({ text, face }) => [text, face]))
      .toEqual([['\u1000', 'Microsoft Sans Serif'], ['\ua9e0\uaa60', 'Meiryo UI']]);
    for (const [lang, mark, seam] of [['en-US', '\ua9e5', false],
      ['my-MM', '\uaa7c', true], ['ja-JP', '\ua9e5', true]] as const) {
      const runs = seam ? [run('\u1000', lang), run(`${mark}B`, lang)] : [run(`\u1000${mark}B`, lang)];
      expect(segments(runs).map(({ text, face }) => [text, face]))
        .toEqual([[`\u1000${mark}`, 'Microsoft Sans Serif'], ['B', 'Corbel']]);
      const { ctx, calls } = context();
      renderTextBody(ctx, body(runs), 0, 0, 300, 100, SCALE);
      expect(calls.find((c) => c.text.includes(mark))?.text).toBe(`\u1000${mark}`);
      expect(calls.find((c) => c.text.includes(mark))?.font).toContain('"Microsoft Sans Serif"');
    }
    expect(segments([run('\u1000\ua9e5', 'fr-FR')])[0].text).toBe('\u1000\ua9e5');
    expect(segments([run('\u1001\ua9e5', 'my-MM')])[0].text).toBe('\u1001\ua9e5');
    for (const mark of '\ua9e5\uaa7b\uaa7c\uaa7d') {
      expect(segments([run(mark, 'my-MM')])[0].face).toBe('Meiryo UI');
    }
    for (const mark of '\uaa7b\uaa7d') {
      expect(segments([run(`\u1000${mark}`, 'my-MM')]).map(({ text, face }) => [text, face]))
        .toEqual([['\u1000', 'Microsoft Sans Serif'], [mark, 'Meiryo UI']]);
    }
    expect(segments([run('\u1000', 'my-MM'), run('\ua9e5', 'en-US')])[0].text).toBe('\u1000\ua9e5');
  });

  it('wraps only at original grapheme boundaries, independent of the recorded font slots', () => {
    const ctx = context().ctx;
    ctx.measureText = (text) => ({ width: [...text].reduce((sum, ch) => sum + (ch === '\u1000' ? 20 : 0), 0),
      actualBoundingBoxAscent: 16, actualBoundingBoxDescent: 4,
      fontBoundingBoxAscent: 16, fontBoundingBoxDescent: 4 }) as TextMetrics;
    for (const [mark, expected] of [['\ua9e5', ['\u1000\ua9e5']], ['\uaa7c', ['\u1000\uaa7c']],
      ['\uaa7b', ['\u1000', '\uaa7b']], ['\uaa7d', ['\u1000', '\uaa7d']]] as const) {
      for (const runs of [[run(`\u1000${mark}`, 'my-MM')], [run('\u1000', 'my-MM'), run(mark, 'my-MM')]]) {
        const result = layoutParagraph(ctx, paragraph(runs), 10, 20, '#000', SCALE, 0);
        expect(result.map((line) => line.segments.map((segment) => segment.text).join(''))).toEqual(expected);
        expect(result[0].segments[0].faceFamily).toBe('Microsoft Sans Serif');
        expect(result.at(-1)?.segments.at(-1)?.faceFamily).toBe(mark === '\ua9e5' || mark === '\uaa7c' ? 'Microsoft Sans Serif' : 'Meiryo UI');
      }
    }
  });

  it('measures and wraps in the selected cs face, including the shape-autofit probe', () => {
    const en = paragraph([run('1 2 3', 'en-US')]);
    const he = paragraph([run('1 2 3', 'he-IL')]);
    expect(layoutParagraph(context().ctx, en, 35, 20, '#000', SCALE, 0)).toHaveLength(1);
    expect(layoutParagraph(context().ctx, he, 35, 20, '#000', SCALE, 0).length).toBeGreaterThan(1);
    expect(naturalWidthExceedsBbox(context().ctx, body(en.runs), 35, 0, 0, SCALE, RC)).toBe(false);
    expect(naturalWidthExceedsBbox(context().ctx, body(he.runs), 35, 0, 0, SCALE, RC)).toBe(true);
  });

  it('includes tracking at font-slot seams in the autofit width', () => {
    expect(naturalWidthExceedsBbox(context().ctx,
      body([run('A§B', 'en-US', { letterSpacing: 3 })]), 35, 0, 0, SCALE, RC)).toBe(true);
  });

  it('draws newly selected ea punctuation in the selected face of an empty ea slot', () => {
    // #1689 emptyea E07/E16: the named cs face draws the neutral symbols; with
    // no cs face (emptyea3 H02) the Latin face does. Language plays no part.
    expect(segments([run('§°±×÷“”', 'ja-JP', { fontFamilyEa: undefined })])[0].face)
      .toBe('Microsoft Sans Serif');
    expect(segments([run('§°±×÷“”', 'ja-JP', {
      fontFamily: 'Perpetua', fontFamilyEa: undefined, fontFamilyCs: undefined,
    })])[0].face).toBe('Perpetua');
  });

  it('maps Japanese backslash before measuring, without changing its Latin slot', () => {
    expect(segments([run('A\\B', 'ja-JP')]).map(({ text, face }) => [text, face])).toEqual([['A¥B', 'Corbel']]);
    expect(segments([run('A\\B', 'en-US')])[0].text).toBe('A\\B');
    expect(naturalWidthExceedsBbox(context().ctx, body([run('A\\B', 'ja-JP')]), 25, 0, 0, SCALE, RC)).toBe(true);
  });

  it('keeps graphemes and formatting across a language-changing run seam', () => {
    expect(segments([run('“', 'ja-JP'), run('\u0301B', 'en-US')]).map(({ text, face }) => [text, face]))
      .toEqual([['“\u0301', 'Meiryo UI'], ['B', 'Corbel']]);
  });

  it('uses the symbol slot throughout U+F0xx and maps known glyphs before layout', () => {
    const result = segments([run('\uf000\uf0b7', 'ja-JP', { fontFamilySym: 'Symbol' })]);
    expect(result[0].face).toBe('Symbol');
    expect(result.map((s) => s.text).join('')).toBe('\uf000•');
    expect(result.at(-1)?.face).toBe('sans-serif');
  });

  it.each(['wordArtVert', 'wordArtVertRtl', 'eaVert', 'vert', 'vert270'])(
    'paints %s text with the same language-selected faces and Japanese glyph mapping', (vert) => {
      const { ctx, calls } = context();
      renderTextBody(ctx, body([run('“§\\', 'ja-JP'), run('123', 'he-IL')], vert), 0, 0, 300, 500, SCALE);
      const fontFor = (ch: string) => calls.find((c) => c.text.includes(ch))?.font;
      expect(fontFor('“')).toContain('"Meiryo UI"');
      expect(fontFor('§')).toContain('"Meiryo UI"');
      expect(fontFor('¥')).toContain('"Corbel"');
      expect(fontFor('1')).toContain('"Microsoft Sans Serif"');
    },
  );
});
