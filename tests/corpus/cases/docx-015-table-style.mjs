import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const { colStart: start, colEnd: end, GUTTER } = DOCX;
const GRID = [start(2) - start(1), start(3) - start(2), start(4) - start(3), end(4) - start(4)];
const rule = (name, sz) => `<w:${name} w:val="single" w:sz="${sz}" w:space="0" w:color="${COLOR.ink}"/>`;

// The table's look lives in a table style; the table only names it and says
// which conditional parts apply (w:tblLook).
const STYLE = '<w:style w:type="table" w:styleId="CorpusTable"><w:name w:val="Corpus Table"/><w:uiPriority w:val="59"/>'
  + `<w:tblPr><w:tblBorders>${rule('top', 4)}${rule('bottom', 4)}</w:tblBorders>`
  + `<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="${GUTTER}" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  + `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:sz w:val="${DOCX.TYPE.caption.sz}"/></w:rPr><w:tcPr><w:tcBorders>${rule('bottom', 2)}</w:tcBorders></w:tcPr></w:tblStylePr>`
  + '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>'
  + '</w:style>';
const LOOK = '<w:tblLook w:val="00A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>';

// ISO 216 B series, width × height in millimetres.
const B = [[1000, 1414], [707, 1000], [500, 707], [353, 500], [250, 353], [176, 250], [125, 176]];
const right = { jc: 'right' };

export default {
  id: 'docx-015',
  slug: 'table-style',
  format: 'docx',
  title: 'Table style',
  covers: [
    'word/styles.xml|w:style@w:type=table',
    'word/styles.xml|w:style>w:tblPr',
    'word/styles.xml|w:tblPr>w:tblBorders',
    'word/styles.xml|w:tblPr>w:tblCellMar',
    'word/styles.xml|w:style>w:tblStylePr',
    'word/styles.xml|w:tblStylePr@w:type=firstRow',
    'word/styles.xml|w:tblStylePr@w:type=firstCol',
    'word/styles.xml|w:style>w:uiPriority',
    'word/document.xml|w:tblPr>w:tblStyle',
    'word/document.xml|w:tblPr>w:tblLook',
    'word/document.xml|w:tblLook@w:firstRow',
    'word/document.xml|w:tblLook@w:firstColumn',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-015',
      title: 'Table style',
      verify: [
        'Verify: rules, the heading hairline, cell margins, the bold',
        'heading row and bold first column come from the table style.',
      ],
    }),
    {
      kind: 'table',
      at: at(2),
      style: 'CorpusTable',
      cellMar: null,
      look: LOOK,
      grid: GRID,
      rows: [
        { lines: 1, cells: [{ text: 'ISO 216' }, { text: 'Width mm', ...right }, { text: 'Height mm', ...right }, { text: 'Area cm²', ...right, marRight: 0 }] },
        ...B.map(([w, h], n) => ({
          lines: 1,
          cells: [{ text: `B${n}` }, { text: String(w), ...right }, { text: String(h), ...right }, { text: (w * h / 100).toFixed(0), ...right, marRight: 0 }],
        })),
      ],
    },
    {
      kind: 'p', before: 1, style: 'caption', cols: [1, 3],
      lines: [
        'The table names the style CorpusTable (w:tblStyle) and applies its first-row and',
        'first-column formatting through w:tblLook; it has no direct borders or margins.',
      ],
    },
  ], { styles: STYLE }),
};
