import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { esc } from '../lib/xml.mjs';

// Table rows across page breaks, the case where most form-like documents
// fail. Three pages:
//   A. A table whose header row repeats (w:tblHeader). Its tenth body row
//      has no w:cantSplit and is taller than the space left on page 1, so
//      Word splits it inside the row: the first lines stay at the page foot,
//      the rest continue on page 2 below the repeated header. Its first cell
//      is vertically merged (w:vMerge) with the next row, so the merge also
//      crosses the break.
//   B. A table of w:cantSplit rows of fixed height. The eighth row does not
//      fit in what is left of page 2, so the whole row moves to page 3; the
//      note after the table follows it.
// Every line in a cell names its row and line, so the split point is
// readable on the page. Cell paragraphs are exact 12 pt lines; rows of A
// other than the split one and all rows of B have exact heights.
const { LINE, colStart, span } = DOCX;
const W1 = colStart(2) - colStart(1);
const W2 = colStart(3) - colStart(2);
const W3 = span(3, 4);
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const TBL_PR = `<w:tblPr><w:tblW w:w="${W1 + W2 + W3}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W2}"/><w:gridCol w:w="${W3}"/></w:tblGrid>`;

const lines = (texts, bold = false) => `<w:p><w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}${texts
  .map((t, i) => `${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${esc(t)}</w:t>`).join('')}</w:r></w:p>`;
const tc = (w, body, extra = '') => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr>${body}</w:tc>`;
const tr = (trPr, cells) => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
const exact = (n) => `<w:trHeight w:val="${n * LINE}" w:hRule="exact"/>`;
const numbered = (row, n, extra = () => '') => Array.from({ length: n }, (_, i) => `Row ${row}, line ${i + 1} of ${n}${extra(i)}`);

// A: header + 9 rows of 4 lines fill lines 13-48 of page 1 (the table
// starts on line 12). Row 10 (15 lines, splittable) starts on line 49.
const A_HEADER = tr(`<w:cantSplit/>${exact(1)}<w:tblHeader/>`,
  tc(W1, lines(['Size'], true)) + tc(W2, lines(['Millimetres'], true)) + tc(W3, lines(['Lines of the row (repeats on page 2)'], true)));
const A_ROWS = A_SERIES.slice(0, 9).map(([n, w, h], i) => tr(exact(4),
  tc(W1, lines([n])) + tc(W2, lines([`${w} × ${h}`]))
  + tc(W3, lines(numbered(i + 1, 4)))));
const A_SPLIT = tr('',
  tc(W1, lines(['A9 and A10', '(merged with', 'the next row)']), '<w:vMerge w:val="restart"/>')
  + tc(W2, lines(['37 × 52']))
  + tc(W3, lines(numbered(10, 15, (i) => (i === 0 ? ', no w:cantSplit' : i === 14 ? ', end of row 10' : '')))));
const A_AFTER = [
  tr(exact(4), tc(W1, lines(['']), '<w:vMerge/>') + tc(W2, lines(['26 × 37']))
    + tc(W3, lines(numbered(11, 4, (i) => (i === 0 ? ', first cell merged' : ''))))),
  tr(exact(4), tc(W1, lines(['A0'])) + tc(W2, lines(['841 × 1189']))
    + tc(W3, lines(numbered(12, 4, (i) => (i === 3 ? ', end of table A' : ''))))),
];
const TABLE_A = `<w:tbl>${TBL_PR}${A_HEADER}${A_ROWS.join('')}${A_SPLIT}${A_AFTER.join('')}</w:tbl>`;

// B: rows of 6 lines with w:cantSplit, no header row. On page 2 table A ends
// on line 12 and caption B sits on line 14, so B rows 1-7 fill lines 15-56;
// row 8 needs 6 lines with 4 left and moves to page 3 whole.
const B_ROWS = Array.from({ length: 9 }, (_, i) => tr(`<w:cantSplit/>${exact(6)}`,
  tc(W1, lines([`B${i + 1}`])) + tc(W2, lines(['w:cantSplit']))
  + tc(W3, lines(numbered(`B${i + 1}`, 6, (k) => (i === 7 && k === 0 ? ', moved whole to page 3' : ''))))));
const TABLE_B = `<w:tbl>${TBL_PR}${B_ROWS.join('')}</w:tbl>`;

const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], ...extra });

export default {
  id: 'docx-035',
  slug: 'table-pagination',
  format: 'docx',
  title: 'Table rows across pages',
  covers: [
    'word/document.xml|w:trPr>w:tblHeader',
    'word/document.xml|w:trPr>w:cantSplit',
    'word/document.xml|w:tcPr>w:vMerge',
    'word/document.xml|w:vMerge@w:val=restart',
    'word/document.xml|w:trHeight@w:hRule=exact',
    'word/document.xml|w:r>w:br',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-035',
      title: 'Table rows across pages',
      verify: [
        'Verify: A row 10 splits at the foot of page 1 and continues',
        'on page 2 under the repeated header; its merged first cell',
        'spans the break. B row 8 moves whole to page 3.',
      ],
    }),
    caption('A. Header row repeats; row 10 may split (no w:cantSplit)', { at: DOCX.rowLine(2) }),
    { kind: 'raw', xml: TABLE_A },
    caption('B. Rows of fixed height with w:cantSplit', { before: 1 }),
    { kind: 'raw', xml: TABLE_B },
    { kind: 'p', before: 1, lines: ['The paragraph after table B follows its last row on page 3.'] },
  ]),
};
