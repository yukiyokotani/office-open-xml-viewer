import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS, wpsShape } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Where Word puts a floating object, and how text flows around it.
//   A. Top-and-bottom wrapping, one parameter at a time: object height 2 or
//      4 lines; vertical offset 0 or 1 line below the anchor paragraph; then
//      square wrapping and no wrapping with the same object. Each anchor
//      paragraph names its case; the paragraph after it shows the flow.
//   B. An object anchored in a table cell (layoutInCell) and aligned to the
//      right and the bottom of the margin with square wrapping: Word places
//      it against the cell, at the top right beside the paragraph, not at
//      the foot of the page.
//   C. A fixed-size outlined text box anchored two lines above its
//      paragraph (negative paragraph offset).
//   D. In a two-column section, an object positioned from the column's left
//      edge in a centred paragraph, followed by a centred caption: the
//      object follows its offset, the caption the column centre.
const { LINE, colStart, span, emu } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const OUTLINE = `<a:ln w="12700"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln>`;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const P = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const raw = (xml) => ({ kind: 'raw', xml, height: NaN });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const W_TB = '<wp:wrapTopAndBottom/>';
const W_SQ = '<wp:wrapSquare wrapText="bothSides"/>';
const W_NO = '<wp:wrapNone/>';

let nextId = 1;
const box = (opts) => wpsShape({ id: nextId++, fill: COLOR.accent, ...opts });

const parts = () => {
  nextId = 1;
  const cases = [
    ['A1 top and bottom, 2 lines, offset 0', W_TB, 2, 0],
    ['A2 top and bottom, 2 lines, offset 1 line', W_TB, 2, 1],
    ['A3 top and bottom, 4 lines, offset 1 line', W_TB, 4, 1],
    ['A4 square, 2 lines, offset 1 line', W_SQ, 2, 1],
    ['A5 none, 2 lines, offset 1 line', W_NO, 2, 1],
  ];
  const A = cases.flatMap(([label, wrap, lines, offset]) => [
    raw(P(box({ x: emu(colStart(3)), y: emu(offset * LINE), cx: emu(span(3, 4)), cy: emu(lines * LINE), wrap }) + r(label))),
    raw(P(r(`The paragraph after ${label.slice(0, 2)}.`))),
  ]);

  const CELL_W = span(1, 4);
  const B = '<w:tbl><w:tblPr>'
    + `<w:tblW w:w="${CELL_W}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('')}</w:tblBorders>`
    + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    + `<w:tblGrid><w:gridCol w:w="${CELL_W}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${CELL_W}" w:type="dxa"/></w:tcPr>`
    + P(r('This cell paragraph anchors the square: margin right, margin bottom, layoutInCell.'))
    + P(box({ h: { from: 'margin', align: 'right' }, v: { from: 'margin', align: 'bottom' }, cx: emu(4 * LINE), cy: emu(4 * LINE), wrap: W_SQ })
      + r('Anchor paragraph; the square sits against this cell.'))
    + P(r('A third line keeps the cell taller than one paragraph.'))
    + '</w:tc></w:tr></w:tbl>';

  const C = P(box({ x: 0, y: emu(-2 * LINE), cx: emu(span(1, 2)), cy: emu(2 * LINE), fill: null, outline: OUTLINE,
    insets: [152400, 0, 152400, 0], wrap: W_NO,
    text: P(r('枠は段落の二行上に置かれる。', JA)) + P(r('固定値の行送りで二行を収める。', JA)) }) + r('This paragraph anchors the box drawn two lines above it.'));

  const D = [
    raw(P(box({ h: { from: 'column', offset: emu(colStart(2) - colStart(1)) }, v: { from: 'paragraph', offset: 0 },
      cx: emu(span(1, 1)), cy: emu(4 * LINE), wrap: '<wp:wrapTopAndBottom/>' }) + r('Centred anchor paragraph'), '<w:jc w:val="center"/>')),
    raw(P(r('Figure 1. Caption centred in the column'), '<w:jc w:val="center"/>')),
    raw(P(r('The left column continues after the figure with ordinary text so the column has content of its own.'))),
  ];

  return docxPackage([
    ...header({
      id: 'docx-043',
      title: 'Floating placement',
      verify: [
        'Verify: A text flows around each case as in Word; B',
        'square sits in the cell, top right; C box is above its',
        'paragraph; D object at its offset, caption centred.',
      ],
    }),
    caption('A. Wrapping, one parameter at a time', { at: DOCX.rowLine(2), before: 0 }),
    ...A,
    caption('B. Anchored in a cell, aligned to margin right and bottom'),
    raw(B),
    caption('C. Text box two lines above its paragraph'),
    { kind: 'p', lines: [''], height: 1 },
    { kind: 'p', lines: [''], height: 1 },
    raw(C),
    caption('D. Two columns: object offset from the column, centred caption'),
    { kind: 'section', type: 'continuous', cols: { num: 2, space: DOCX.gutterAfter(2) } },
    ...D,
  ], {
    ns: [A_NS, WP_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-043',
  slug: 'floating-placement',
  format: 'docx',
  title: 'Floating placement',
  covers: [
    'word/document.xml|wp:anchor>wp:wrapTopAndBottom',
    'word/document.xml|wp:anchor>wp:wrapSquare',
    'word/document.xml|wp:anchor>wp:wrapNone',
    'word/document.xml|wp:positionH>wp:align',
    'word/document.xml|wp:positionV>wp:align',
    'word/document.xml|wp:positionH@relativeFrom=column',
    'word/document.xml|wp:positionV@relativeFrom=paragraph',
    'word/document.xml|w:tc>w:p',
  ],
  parts,
};
