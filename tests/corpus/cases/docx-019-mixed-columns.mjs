import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';
import { A_SERIES } from '../lib/iso216.mjs';

// A paper laid out the way journals and conference templates are: a
// one-column abstract, a two-column body, a full-width table between two
// two-column parts, a column break, and a two-column part that runs onto the
// second page and is balanced there above a closing one-column note. Every
// change of layout is a continuous section break, so up to four sections share
// a page.
const COLS = { num: 2, space: DOCX.gutterAfter(2) };
const B_SERIES = [
  ['B0', 1000, 1414], ['B1', 707, 1000], ['B2', 500, 707], ['B3', 353, 500], ['B4', 250, 353], ['B5', 176, 250],
  ['B6', 125, 176], ['B7', 88, 125], ['B8', 62, 88], ['B9', 44, 62], ['B10', 31, 44],
];
const C_SERIES = [
  ['C0', 917, 1297], ['C1', 648, 917], ['C2', 458, 648], ['C3', 324, 458], ['C4', 229, 324], ['C5', 162, 229],
  ['C6', 114, 162], ['C7', 81, 114], ['C8', 57, 81], ['C9', 40, 57], ['C10', 28, 40],
];

const ratio = (w, h) => (h / w).toFixed(3);
const sizeText = (series, i) => {
  const [name, w, h] = series[i];
  const next = series[i + 1];
  return `${name} measures ${w} by ${h} millimetres; the long side divided by the short side is ${ratio(w, h)}, `
    + `close to the square root of two, 1.414. `
    + (next ? `Folding it across the long side gives ${next[0]}, ${next[1]} by ${next[2]} millimetres, `
      + 'the same shape at half the area, because each size is the previous one cut in half.'
      : `It is the smallest size the standard names in this series.`);
};
const para = (text, before = 0) => ({ kind: 'p', text, before });
const heading = (text) => ({ kind: 'p', bold: true, lines: [text], before: 1 });

const TABLE_GRID = [DOCX.colStart(2) - DOCX.colStart(1), DOCX.colStart(3) - DOCX.colStart(2), DOCX.colStart(4) - DOCX.colStart(3), DOCX.span(4, 4)];
const cell = (text, bold = false) => ({ text, ...(bold ? { bold: true } : {}) });
const TABLE = {
  kind: 'table',
  before: 1,
  grid: TABLE_GRID,
  borders: { top: 8, bottom: 8, insideH: 4 },
  rows: [
    { lines: 1, cells: [cell('Series', true), cell('Size 0 (mm)', true), cell('Size 4 (mm)', true), cell('Use', true)] },
    { lines: 1, cells: [cell('A'), cell('841 × 1189'), cell('210 × 297'), cell('Letters, printed pages')] },
    { lines: 1, cells: [cell('B'), cell('1000 × 1414'), cell('250 × 353'), cell('Posters, books')] },
    { lines: 1, cells: [cell('C'), cell('917 × 1297'), cell('229 × 324'), cell('Envelopes for A sizes')] },
  ],
};

export default {
  id: 'docx-019',
  slug: 'mixed-columns',
  format: 'docx',
  title: 'Mixed columns',
  covers: [
    'word/document.xml|w:sectPr>w:cols',
    'word/document.xml|w:cols@w:num',
    'word/document.xml|w:type@w:val=continuous',
    'word/document.xml|w:pPr>w:sectPr',
    'word/document.xml|w:br@w:type=column',
    'word/document.xml|w:body>w:tbl',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-019',
      title: 'Mixed columns',
      verify: [
        'Verify: page 1 holds a one-column abstract, two columns',
        '(the right one starting at a column break), a full-width',
        'table, then two columns that continue on page 2 and end',
        'balanced above a one-column note.',
      ],
    }),
    {
      kind: 'p',
      before: 1,
      runs: [
        run('Abstract. ', '<w:b/>'),
        run('This specimen sets one text in five sections. The abstract spans the full measure; the body '
          + 'below is set in two columns; a table spans the full measure again between two-column parts; '
          + 'the second part runs onto the next page and ends balanced, its two columns of nearly equal '
          + 'length, before a closing note in one column. Each change of layout is made by a continuous '
          + 'section break, which ends a section without starting a new page.'),
      ],
    },
    { kind: 'section', type: 'continuous', cols: COLS },
    heading('1 The A series'),
    ...A_SERIES.slice(0, 4).map((_, i) => para(sizeText(A_SERIES, i), i ? 1 : 0)),
    { kind: 'p', columnBreakBefore: true, text: 'This paragraph follows a column break (w:br w:type="column"), so it '
      + 'starts at the top of the right column whatever the length of the left one.' },
    ...A_SERIES.slice(4, 8).map((_, i) => para(sizeText(A_SERIES, i + 4), 1)),
    { kind: 'section', type: 'continuous' },
    { kind: 'p', before: 1, style: 'caption', bold: true, lines: ['Table 1. The three ISO 216 and ISO 269 series, sizes 0 and 4'] },
    TABLE,
    { kind: 'p', style: 'caption', lines: ['The table spans the full measure in a one-column section between two-column parts.'], before: 1 },
    { kind: 'section', type: 'continuous', cols: COLS },
    heading('2 The B and C series'),
    ...B_SERIES.slice(0, -1).map((_, i) => para(sizeText(B_SERIES, i), i ? 1 : 0)),
    ...C_SERIES.slice(0, -1).map((_, i) => para(sizeText(C_SERIES, i), 1)),
    heading('3 Conclusion'),
    para('Every size in the three series has sides in the ratio of the square root of two, so halving a sheet '
      + 'keeps its shape. The B sizes lie between neighbouring A sizes, the C sizes between A and B, which is '
      + 'why a C4 envelope holds an unfolded A4 sheet.'),
    para('These two columns end on the second page. The continuous section break that follows balances them, '
      + 'so the left and right columns finish on nearly the same line instead of the left column filling first.', 1),
    { kind: 'section', type: 'continuous' },
    { kind: 'p', before: 1, style: 'caption', text: 'Closing note in one column: a final section of one column after the '
      + 'balanced two-column part. The page continues below it as one column.' },
  ]),
};
