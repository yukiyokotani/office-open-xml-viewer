import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const { colStart: start, colEnd: end, GUTTER } = DOCX;

// ISO 216 A series, width × height in millimetres.
const A_SERIES = [
  [841, 1189], [594, 841], [420, 594], [297, 420], [210, 297], [148, 210],
  [105, 148], [74, 105], [52, 74], [37, 52], [26, 37],
];

// Table columns coincide with the grid columns: each cell spans its column and
// the following gutter, and the right cell margin is one gutter, so text runs
// from the column's left edge to its right edge. The last column has no gutter.
const GRID = [start(2) - start(1), start(3) - start(2), start(4) - start(3), end(4) - start(4)];
const RULE = 4; // 0.5 pt, in eighths of a point
const HAIRLINE = 2; // 0.25 pt

const cell = (text, options = {}) => ({ text, ...options });
const lastColumn = { marRight: 0 };

export default {
  id: 'docx-003',
  slug: 'table-iso-216',
  format: 'docx',
  title: 'Table',
  covers: [
    'word/document.xml|w:tbl>w:tblGrid',
    'word/document.xml|w:tblGrid>w:gridCol',
    'word/document.xml|w:tblLayout@w:type=fixed',
    'word/document.xml|w:tblBorders>w:top',
    'word/document.xml|w:tblBorders>w:bottom',
    'word/document.xml|w:tcBorders>w:bottom',
    'word/document.xml|w:tblCellMar>w:right',
    'word/document.xml|w:tcMar>w:right',
    'word/document.xml|w:trHeight@w:hRule=exact',
    'word/document.xml|w:vAlign@w:val=bottom',
    'word/document.xml|w:jc@w:val=right',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-003',
      title: 'Table',
      verify: [
        'Verify: column edges at 42.5, 85 and 127.5 mm;',
        'numbers flush right on the column edge; rules',
        '0.5 pt, hairline under the heading 0.25 pt.',
      ],
    }),
    {
      kind: 'table',
      at: at(2),
      grid: GRID,
      cellMar: { left: 0, right: GUTTER },
      borders: { top: RULE, bottom: RULE },
      rows: [
        {
          lines: 2,
          cells: [
            cell('ISO 216', { style: 'caption', bold: true, vAlign: 'bottom', borders: { bottom: HAIRLINE } }),
            cell('Width mm', { style: 'caption', bold: true, jc: 'right', vAlign: 'bottom', borders: { bottom: HAIRLINE } }),
            cell('Height mm', { style: 'caption', bold: true, jc: 'right', vAlign: 'bottom', borders: { bottom: HAIRLINE } }),
            cell('Height ÷ width', { style: 'caption', bold: true, jc: 'right', vAlign: 'bottom', borders: { bottom: HAIRLINE }, ...lastColumn }),
          ],
        },
        ...A_SERIES.map(([w, h], n) => ({
          lines: 1,
          cells: [
            cell(`A${n}`),
            cell(String(w), { jc: 'right' }),
            cell(String(h), { jc: 'right' }),
            cell((h / w).toFixed(3), { jc: 'right', ...lastColumn }),
          ],
        })),
      ],
    },
    {
      kind: 'p',
      at: at(3) + 4,
      style: 'caption',
      cols: [1, 3],
      lines: [
        'Width and height in millimetres, as defined by ISO 216.',
        'The ratio is height divided by width, rounded to three decimals;',
        'the ideal ratio of the series is the square root of 2, 1.414.',
      ],
    },
  ]),
};
