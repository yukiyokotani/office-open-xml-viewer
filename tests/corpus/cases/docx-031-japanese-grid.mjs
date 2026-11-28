import { FONT } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

// The document grid of Japanese Word documents (w:docGrid) and punctuation
// compression (w:characterSpacingControl compressPunctuation), the defaults
// of most Japanese templates. Three sections share one page:
//   A. type="lines", linePitch 18 pt: single-spaced lines snap to 18 pt.
//   B. type="linesAndChars", charSpace 0: characters also sit on a pitch of
//      the default size, so a line holds a whole number of characters.
//   C. type="linesAndChars", charSpace −1161: the pitch narrows by
//      1161/4096 pt, so more characters fit on a line.
// Each section ends with the same paragraph with w:snapToGrid off, set
// without the grid for comparison. Brackets, commas and stops in a row show
// compression; Latin words and digits show the automatic spacing
// (autoSpaceDE / autoSpaceDN) between them and Japanese.
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const TEXT = '約物が続く箇所、「（例）」や「。」」のような並びは、詰めの設定によって半角分ずつ詰められる。'
  + 'A4判は210×297 mm、B5判は176×250 mmであり、英数字と和文の間には自動で四分の空きが入る。'
  + '一行に入る文字数は、文字グリッドがあれば字送りで決まり、なければ字形の幅で決まる。';
const para = (snap = true) => ({
  kind: 'raw',
  xml: `<w:p><w:pPr>${snap ? '' : '<w:snapToGrid w:val="0"/>'}<w:spacing w:before="240" w:line="240" w:lineRule="auto"/></w:pPr>`
    + `${run(snap ? TEXT : `【グリッドなし】${TEXT}`, JA)}</w:p>`,
});
const label = (text) => ({ kind: 'raw', xml: `<w:p><w:pPr><w:spacing w:before="240" w:line="240" w:lineRule="auto"/></w:pPr>${run(text, '<w:b/><w:sz w:val="16"/>')}</w:p>` });
const grid = (type, charSpace) => `<w:docGrid w:type="${type}" w:linePitch="360"${charSpace === undefined ? '' : ` w:charSpace="${charSpace}"`}/>`;

export default {
  id: 'docx-031',
  slug: 'japanese-grid',
  format: 'docx',
  title: 'Japanese document grid',
  covers: [
    'word/document.xml|w:sectPr>w:docGrid',
    'word/document.xml|w:docGrid@w:type=lines',
    'word/document.xml|w:docGrid@w:type=linesAndChars',
    'word/document.xml|w:docGrid@w:charSpace',
    'word/document.xml|w:pPr>w:snapToGrid',
    'word/settings.xml|w:settings>w:characterSpacingControl',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-031',
      title: 'Japanese document grid',
      verify: [
        'Verify: in A-C grid lines are 18 pt apart; B and C set',
        'characters on a fixed pitch, C narrower; each section ends',
        'with the same text off the grid. Punctuation pairs close up.',
      ],
    }),
    label('A. docGrid lines, linePitch 18 pt'),
    para(), para(false),
    { kind: 'p', lines: [' '] },
    { kind: 'section', type: 'continuous', docGrid: grid('linesAndChars', 0) },
    label('B. docGrid linesAndChars, charSpace 0'),
    para(), para(false),
    { kind: 'p', lines: [' '] },
    { kind: 'section', type: 'continuous', docGrid: grid('linesAndChars', -1161) },
    label('C. docGrid linesAndChars, charSpace −1161'),
    para(), para(false),
  ], {
    firstSection: { docGrid: grid('lines') },
    settings: '<w:characterSpacingControl w:val="compressPunctuation"/>',
  }),
};
