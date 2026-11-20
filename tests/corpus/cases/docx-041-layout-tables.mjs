import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { A_NS } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Tables used as page layout, the way résumé and leaflet templates build it.
//   A. A table 100 % wide (w:tblW pct) whose grid alternates content columns
//      and narrow spacer columns. Content cells give their width in percent,
//      spacer cells in twips; rows are at least four lines tall. A band
//      behind the table, a full-width rectangle anchored in the first cell
//      to the page horizontally and to the paragraph vertically (behindDoc),
//      marks the first row.
//   B. A bar chart made of a table nested in a cell: each bar is a shaded
//      cell spanning 3, 2 or 1 grid columns, on a row 6 pt high.
//   C. A narrow cell with vertical text (w:textDirection tbRlV) whose two
//      paragraphs are distributed (w:jc distribute) along the row; the two
//      lines are wider together than the cell. The row stays near its
//      authored height of four lines instead of growing.
const { LINE, colStart, span, emu, GUTTER, PAGE_W } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const para = (text, { bold = false, rPr = '', pPr = '' } = {}) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${bold || rPr ? `<w:rPr>${bold ? '<w:b/>' : ''}${rPr}</w:rPr>` : ''}${t(text)}</w:r></w:p>`;
const lines = (texts) => `<w:p><w:r>${texts.map((x, i) => `${i ? '<w:br/>' : ''}${t(x)}`).join('')}</w:r></w:p>`;
const tcPr = (w, type, extra = '') => `<w:tcPr><w:tcW w:w="${w}" w:type="${type}"/>${extra}</w:tcPr>`;

// A: columns [content 1 | spacer | content 2 | spacer | content 3-4].
const GRID_A = [span(1, 1), DOCX.gutterAfter(1), span(2, 2), DOCX.gutterAfter(2), span(3, 4)];
const pctOf = (w) => Math.round((w / DOCX.LIVE_W) * 5000);
const BAND = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
  + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
  + `<wp:extent cx="${emu(PAGE_W)}" cy="${emu(LINE)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
  + '<wp:docPr id="1" name="Band"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(PAGE_W)}" cy="${emu(LINE)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  + `<a:solidFill><a:srgbClr val="${COLOR.accent}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
const spacer = (i) => `<w:tc>${tcPr(GRID_A[i], 'dxa')}<w:p/></w:tc>`;
const content = (i, body) => `<w:tc>${tcPr(pctOf(GRID_A[i]), 'pct')}${body}</w:tc>`;
const rowA = (cells) => `<w:tr><w:trPr><w:trHeight w:val="${4 * LINE}" w:hRule="atLeast"/></w:trPr>${cells}</w:tr>`;
const HISTORY = [
  ['1975', 'ISO 216', ['Paper sizes A, B and C, the', 'series based on 1 : 1.414 sheets.']],
  ['1922', 'DIN 476', ['The German standard the', 'A series was first defined in.']],
];

// B: nested table, 3 grid columns inside content columns 2-4.
const GRID_B = [colStart(3) - colStart(2), colStart(4) - colStart(3), span(4, 4)];
const SHADE = `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>`;
const barRow = (n) => {
  const filled = GRID_B.slice(0, n).reduce((a, b) => a + b, 0);
  const rest = GRID_B.slice(n).reduce((a, b) => a + b, 0);
  return `<w:tr><w:trPr><w:trHeight w:val="${LINE / 2}" w:hRule="exact"/></w:trPr>`
    + `<w:tc>${tcPr(filled, 'dxa', `${n > 1 ? `<w:gridSpan w:val="${n}"/>` : ''}${SHADE}`)}<w:p/></w:tc>`
    + (rest ? `<w:tc>${tcPr(rest, 'dxa', 3 - n > 1 ? `<w:gridSpan w:val="${3 - n}"/>` : '')}<w:p/></w:tc>` : '') + '</w:tr>';
};
const labelRow = (text) => `<w:tr><w:trPr><w:trHeight w:val="${LINE}" w:hRule="exact"/></w:trPr><w:tc>${tcPr(span(2, 4), 'dxa', '<w:gridSpan w:val="3"/>')}${para(text)}</w:tc></w:tr>`;
const gapRow = `<w:tr><w:trPr><w:trHeight w:val="${LINE / 2}" w:hRule="exact"/></w:trPr><w:tc>${tcPr(span(2, 4), 'dxa', '<w:gridSpan w:val="3"/>')}<w:p/></w:tc></w:tr>`;
const NESTED = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${span(2, 4)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  + `<w:tblGrid>${GRID_B.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  + [['Tables, 3 of 3 columns', 3], ['Sections, 2 of 3 columns', 2], ['Fields, 1 of 3 columns', 1]]
    .map(([label, n]) => labelRow(label) + barRow(n) + gapRow).join('')
  + '</w:tbl>';

const TABLE_A = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/>'
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid>${GRID_A.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  + rowA(content(0, `<w:p>${BAND}<w:r><w:rPr><w:b/></w:rPr>${t('Period')}</w:r></w:p>`) + spacer(1)
    + content(2, para('Standard', { bold: true })) + spacer(3) + content(4, para('Scope (band behind this row)', { bold: true })))
  + HISTORY.map(([year, name, scope]) => rowA(content(0, para(year)) + spacer(1) + content(2, para(name)) + spacer(3) + content(4, lines(scope)))).join('')
  + rowA(content(0, para('Use', { bold: true })) + spacer(1)
    + `<w:tc>${tcPr(pctOf(span(2, 4)), 'pct', '<w:gridSpan w:val="3"/>')}${NESTED}<w:p/></w:tc>`)
  + '</w:tbl>';

// C: narrow vertical cell, two distributed paragraphs.
const DIST = `<w:jc w:val="distribute"/>`;
const TABLE_C = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${GUTTER + span(2, 4)}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideV']
    .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('')}</w:tblBorders>`
  + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${GUTTER}"/><w:gridCol w:w="${span(2, 4)}"/></w:tblGrid>`
  + `<w:tr><w:trPr><w:trHeight w:val="${4 * LINE}" w:hRule="atLeast"/></w:trPr>`
  + `<w:tc>${tcPr(GUTTER, 'dxa', '<w:textDirection w:val="tbRlV"/>')}${para('職・氏名', { rPr: JA, pPr: DIST })}${para('代表者', { rPr: JA, pPr: DIST })}</w:tc>`
  + `<w:tc>${tcPr(span(2, 4), 'dxa')}${lines(['The left cell is one gutter wide, narrower than its two', 'stacked vertical lines; the row stays four lines tall.'])}</w:tc></w:tr></w:tbl>`;

const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

export default {
  id: 'docx-041',
  slug: 'layout-tables',
  format: 'docx',
  title: 'Tables as page layout',
  covers: [
    'word/document.xml|w:tblW@w:type=pct',
    'word/document.xml|w:tcW@w:type=pct',
    'word/document.xml|w:trHeight@w:hRule=atLeast',
    'word/document.xml|w:tc>w:tbl',
    'word/document.xml|w:tcPr>w:shd',
    'word/document.xml|w:tcPr>w:textDirection',
    'word/document.xml|w:textDirection@w:val=tbRlV',
    'word/document.xml|w:jc@w:val=distribute',
    'word/document.xml|wp:anchor@behindDoc',
    'word/document.xml|wp:anchor@layoutInCell',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-041',
      title: 'Tables as page layout',
      verify: [
        'Verify: A columns keep their widths and the band spans the',
        'page behind row 1; B bars span 3, 2 and 1 columns; C row',
        'stays four lines tall with the vertical text distributed.',
      ],
    }),
    caption('A. 100 % table with spacer columns; B. bars in a nested table', { at: DOCX.rowLine(2), before: 0 }),
    { kind: 'raw', xml: TABLE_A, height: NaN },
    caption('C. Vertical distributed text in a narrow cell'),
    { kind: 'raw', xml: TABLE_C, height: NaN },
  ], {
    ns: [A_NS,
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'],
  }),
};
