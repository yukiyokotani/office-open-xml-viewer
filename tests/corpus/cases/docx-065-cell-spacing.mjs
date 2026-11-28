import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { p, r } from '../lib/wml.mjs';

// Tables with w:tblCellSpacing of 2 pt, automatic layout (no w:tblLayout) and
// single 0.5 pt borders, so the table and every cell draw their own rules.
// Every case has two cells whose preferred widths are one and two grid
// columns, unless stated.
//   A. Table width auto: the table grows by the spacing at every outer and
//      shared boundary; the cells keep their widths.
//   B. Table width the sum of the cell widths (dxa): the table keeps that
//      width, the cells share what the gaps leave.
//   C. Table width 50 % (pct): the same, from half the column.
//   D. A 100 %-wide spaced table nested in the cell of a spaced table.
//   E. A negative table indent with two half-column cells: the automatic
//      table may reach the end of the column.
//   F. A first column merged down 64 rows (w:vMerge) across the page break,
//      shaded on the restart cell only: one box, filled on both pages.
const { span } = DOCX;
const SPACING = '<w:tblCellSpacing w:w="40" w:type="dxa"/>';
const BORDERS = '<w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')
  + '</w:tblBorders>';
const SHD = `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>`;
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, height: NaN, ...extra });

const tc = (w, content, tcPr = '') => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${tcPr}</w:tcPr>${content}</w:tc>`;
const table = ({ tblW = '<w:tblW w:w="0" w:type="auto"/>', ind = '', grid, rows }) => '<w:tbl><w:tblPr>'
  + `${tblW}${SPACING}${ind}${BORDERS}</w:tblPr>`
  + `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  + rows.map((cells) => `<w:tr>${cells.join('')}</w:tr>`).join('')
  + '</w:tbl>';

const W1 = span(1, 1);
const W2 = span(2, 3);
const HALF = span(1, 2);
const pair = (name, extra = {}) => table({
  grid: [W1, W2],
  rows: [[tc(W1, p('', r(`${name} cell 1`))), tc(W2, p('', r(`${name} cell 2`)))]],
  ...extra,
});

const nested = table({
  tblW: '<w:tblW w:w="5000" w:type="pct"/>',
  grid: [span(1, 1), span(2, 2)],
  rows: [[tc(span(1, 1), p('', r('D inner 1'))), tc(span(2, 2), p('', r('D inner 2')))]],
});
const outer = table({
  grid: [span(1, 3)],
  rows: [[tc(span(1, 3), `${p('', r('D outer cell'))}${nested}${p('', r('D after the inner table'))}`)]],
});

const ROWS = 64;
const merged = table({
  grid: [W1, W2],
  rows: Array.from({ length: ROWS }, (_, i) => [
    tc(W1, p('', i === 0 ? r('F merged, shaded') : ''), i === 0 ? `<w:vMerge w:val="restart"/>${SHD}` : '<w:vMerge/>'),
    tc(W2, p('', r(`F row ${String(i + 1).padStart(2, '0')}`))),
  ]),
});

export default {
  id: 'docx-065',
  slug: 'cell-spacing',
  format: 'docx',
  title: 'Cell spacing',
  covers: [
    'word/document.xml|w:tblPr>w:tblCellSpacing',
    'word/document.xml|w:tblW@w:type=auto',
    'word/document.xml|w:tblW@w:type=dxa',
    'word/document.xml|w:tblW@w:type=pct',
    'word/document.xml|w:tblPr>w:tblInd',
    'word/document.xml|w:tcPr>w:vMerge',
    'word/document.xml|w:tcPr>w:shd',
    'word/document.xml|w:tc>w:tbl',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-065',
      title: 'Cell spacing',
      verify: [
        'Verify: table and cells draw separate rules 2 pt apart;',
        'A is wider than its cells, B and C keep their width,',
        'F is one shaded box on both pages. Compare the PDF.',
      ],
    }),
    caption('A. Table width auto', { at: DOCX.rowLine(2), before: 0 }),
    raw(pair('A')),
    caption('B. Table width = the cell widths (dxa)'),
    raw(pair('B', { tblW: `<w:tblW w:w="${W1 + W2}" w:type="dxa"/>` })),
    caption('C. Table width 50 % (pct)'),
    raw(pair('C', { tblW: '<w:tblW w:w="2500" w:type="pct"/>' })),
    caption('D. A 100 % spaced table nested in a cell'),
    raw(outer),
    caption('E. Negative indent, two half-column cells'),
    raw(table({
      ind: '<w:tblInd w:w="-800" w:type="dxa"/>',
      grid: [HALF, HALF],
      rows: [[tc(HALF, p('', r('E cell 1'))), tc(HALF, p('', r('E cell 2')))]],
    })),
    caption('F. First column merged across the page break'),
    raw(merged),
    { kind: 'p', lines: ['F: after the table.'], height: NaN },
  ]),
};
