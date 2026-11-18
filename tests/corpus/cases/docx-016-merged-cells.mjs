import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const { colStart: start, colEnd: end, GUTTER } = DOCX;
// The table starts on column 2 (w:tblInd) and spans columns 2-4.
const GRID = [start(3) - start(2), start(4) - start(3), end(4) - start(4)];
const shd = (fill) => `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
const ACCENT = shd(COLOR.accent);
const right = { jc: 'right' };
const last = { marRight: 0 };

export default {
  id: 'docx-016',
  slug: 'merged-cells',
  format: 'docx',
  title: 'Merged cells and shading',
  covers: [
    'word/document.xml|w:tcPr>w:gridSpan',
    'word/document.xml|w:tcPr>w:vMerge',
    'word/document.xml|w:vMerge@w:val=restart',
    'word/document.xml|w:tcPr>w:shd',
    'word/document.xml|w:tblPr>w:tblInd',
    'word/document.xml|w:vAlign@w:val=center',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-016',
      title: 'Merged cells',
      verify: [
        'Verify: the table starts on column 2; the heading spans two',
        'columns; the label spans three rows, centred; cyan cells filled.',
      ],
    }),
    {
      kind: 'table',
      at: at(2),
      ind: start(2),
      cellMar: { left: 0, right: GUTTER },
      borders: { top: 4, bottom: 4 },
      grid: GRID,
      rows: [
        { lines: 1, cells: [{ text: 'Series', style: 'caption', bold: true, borders: { bottom: 2 } }, { text: 'Width and height, mm', style: 'caption', bold: true, gridSpan: 2, borders: { bottom: 2 }, ...right, ...last }] },
        { lines: 1, cells: [{ text: 'A', vMerge: 'restart', vAlign: 'center' }, { text: '210', ...right }, { text: '297', ...right, ...last }] },
        { lines: 1, cells: [{ vMerge: 'continue', vAlign: 'center' }, { text: '148', ...right, shd: ACCENT }, { text: '210', ...right, shd: ACCENT, ...last }] },
        { lines: 1, cells: [{ vMerge: 'continue', vAlign: 'center' }, { text: '105', ...right }, { text: '148', ...right, ...last }] },
        { lines: 1, cells: [{ text: 'B' }, { text: '176', ...right }, { text: '250', ...right, ...last }] },
      ],
    },
    {
      kind: 'p', before: 1, style: 'caption', cols: [2, 4],
      lines: [
        `w:tblInd ${start(2)}: the table edge is the left edge of grid column 2.`,
        'Heading: w:gridSpan 2. Series A: w:vMerge restart, then two continue cells.',
        `A5 row: w:shd fill ${COLOR.accent} on both number cells.`,
      ],
    },
  ]),
};
