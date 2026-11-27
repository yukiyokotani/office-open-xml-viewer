import { LANG, PPTX } from '../lib/grid.mjs';
import { field, header, pptxPackage, table, textBox } from '../lib/pptx.mjs';
import { esc } from '../lib/xml.mjs';

// a:tr@h is a minimum ([MS-OE376] §2.1.1347): PowerPoint grows a row to the
// text of its tallest cell and moves the rows below it down. Every row here
// has the same 24 pt minimum and the frame extent equals their sum, as
// generators that write a:tr@h without laying out the text do; the extent is
// not where the grown table ends.
const RULE = 12700; // 1 pt
const HAIRLINE = 6350; // 0.5 pt
const ROW_H = 2 * PPTX.LINE; // 24 pt

// Label column three grid columns, text column six, the last column the rest.
// The right cell margin is one gutter, so text ends on a column edge.
const COLUMNS = [3 * PPTX.STEP, 6 * PPTX.STEP, PPTX.w(10, 12)];
const last = (i) => i === COLUMNS.length - 1;

// One paragraph whose lines are separated by a:br, so the line count does
// not depend on the font.
const brokenLines = (lines) => {
  const rPr = `<a:rPr lang="${LANG.latin}" sz="${PPTX.TYPE.text.sz}"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr>`;
  return `<a:p><a:pPr><a:lnSpc><a:spcPts val="${PPTX.TYPE.text.leading}"/></a:lnSpc></a:pPr>`
    + lines.map((text) => `<a:r>${rPr}<a:t>${esc(text)}</a:t></a:r>`).join(`<a:br>${rPr}</a:br>`)
    + '</a:p>';
};

const ROWS = [
  ['Breaks', { xml: brokenLines([
    'This cell holds four lines',
    'separated by three line breaks,',
    'so its row grows from the',
    '24 pt minimum to 96 pt.',
  ]) }, '4 lines'],
  ['Wrap', {
    text: 'This sentence has no line break: it wraps at the edge of the column, and its row grows to every line it takes.',
  }, 'wrapped'],
  ['Fits', { text: 'One line keeps the minimum.' }, '24 pt'],
];

const headCell = (text, i) => ({
  text, style: 'caption', bold: true, anchor: 'b',
  marR: last(i) ? 0 : undefined, borders: { T: RULE, B: HAIRLINE },
});
// A shared edge is written on both cells that meet there.
const bodyCell = (content, i, r) => ({
  ...(typeof content === 'string' ? { text: content } : content),
  marR: last(i) ? 0 : undefined,
  borders: { T: HAIRLINE, B: r === ROWS.length - 1 ? RULE : HAIRLINE },
});

export default {
  id: 'pptx-020',
  slug: 'table-row-growth',
  format: 'pptx',
  title: 'Table row growth',
  covers: [
    'ppt/slides/slide1.xml|a:graphicData>a:tbl',
    'ppt/slides/slide1.xml|a:tr@h',
    'ppt/slides/slide1.xml|a:tc>a:txBody',
    'ppt/slides/slide1.xml|a:p>a:br',
    'ppt/slides/slide1.xml|a:tcPr>a:lnT',
    'ppt/slides/slide1.xml|a:tcPr>a:lnB',
  ],
  parts: () => {
    const rows = [
      { h: ROW_H, cells: ['Row', 'Text', 'Lines'].map(headCell) },
      ...ROWS.map((cells, r) => ({ h: ROW_H, cells: cells.map((c, i) => bodyCell(c, i, r)) })),
    ];
    // The frame extent is the sum of the minima (96 pt), not the grown height.
    const box = { ...field(1, 12, 2, 2), cy: rows.length * ROW_H };
    return pptxPackage([
      header({
        id: 'pptx-020',
        title: 'Table row growth',
        verify: [
          'Verify: each row is as tall as its text; the Breaks row is 96 pt, the Wrap row',
          'wraps at the text column edge; every rule sits below the text above it, none crosses it.',
        ],
      }),
      table(3, 'Rows', box, COLUMNS, rows),
      textBox(4, 'Caption', field(1, 6, 4, 4), [
        { style: 'caption', text: 'Every a:tr@h is 24 pt and the frame extent is their sum, 96 pt.' },
        { style: 'caption', text: 'a:tr@h is a minimum: PowerPoint grows a row to its tallest cell' },
        { style: 'caption', text: 'and moves the rows below it down ([MS-OE376] 2.1.1347).' },
      ]),
    ]);
  },
};
