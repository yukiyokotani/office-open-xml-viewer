import { PPTX } from '../lib/grid.mjs';
import { field, header, pptxPackage, table, textBox } from '../lib/pptx.mjs';

// Length units in English Metric Units; points and millimetres follow from
// 1 in = 914400 EMU = 72 pt = 25.4 mm.
const UNITS = [
  ['inch', 914400], ['centimetre', 360000], ['pica', 152400],
  ['millimetre', 36000], ['point', 12700], ['pixel at 96 dpi', 9525],
];
const RULE = 12700; // 1 pt
const HAIRLINE = 6350; // 0.5 pt
const ROW_H = 2 * PPTX.LINE; // 24 pt

// Table columns span three grid columns each; the right cell margin is one
// gutter, so text ends on a column edge. The last column has no gutter.
const COLUMNS = [3 * PPTX.STEP, 3 * PPTX.STEP, 3 * PPTX.STEP, PPTX.w(10, 12)];
const last = (i) => i === COLUMNS.length - 1;

const headCell = (text, i) => ({
  text, style: 'caption', bold: true, algn: i ? 'r' : undefined, anchor: 'b',
  marR: last(i) ? 0 : undefined, borders: { T: RULE, B: HAIRLINE },
});
// A shared edge is written on both cells that meet there.
const bodyCell = (text, i, r) => ({
  text, algn: i ? 'r' : undefined, marR: last(i) ? 0 : undefined,
  borders: { T: r === 0 ? HAIRLINE : undefined, B: r === UNITS.length - 1 ? RULE : undefined },
});

export default {
  id: 'pptx-003',
  slug: 'table-units',
  format: 'pptx',
  title: 'Table',
  covers: [
    'ppt/slides/slide1.xml|a:graphicData>a:tbl',
    'ppt/slides/slide1.xml|a:tblGrid>a:gridCol',
    'ppt/slides/slide1.xml|a:tr@h',
    'ppt/slides/slide1.xml|a:tcPr@marL',
    'ppt/slides/slide1.xml|a:tcPr@marR',
    'ppt/slides/slide1.xml|a:tcPr@anchor=b',
    'ppt/slides/slide1.xml|a:tcPr>a:lnT',
    'ppt/slides/slide1.xml|a:tcPr>a:lnB',
    'ppt/slides/slide1.xml|a:lnL>a:noFill',
    'ppt/slides/slide1.xml|a:pPr@algn=r',
  ],
  parts: () => {
    const rows = [
      { h: ROW_H, cells: ['Unit', 'EMU', 'Points', 'Millimetres'].map(headCell) },
      ...UNITS.map(([name, emu], r) => ({
        h: ROW_H,
        cells: [name, String(emu), (emu / 12700).toFixed(4), (emu / 36000).toFixed(4)]
          .map((text, i) => bodyCell(text, i, r)),
      })),
    ];
    const box = { ...field(1, 12, 2, 2), cy: rows.length * ROW_H };
    return pptxPackage([
      header({
        id: 'pptx-003',
        title: 'Table',
        verify: [
          'Verify: table columns start on the edges of columns 1, 4, 7 and 10;',
          'numbers flush right on a column edge; rules 1 pt, hairline 0.5 pt.',
        ],
      }),
      table(3, 'Units', box, COLUMNS, rows),
      textBox(4, 'Caption', field(1, 6, 4, 4), [
        { style: 'caption', text: 'ECMA-376 measures DrawingML lengths in English Metric Units:' },
        { style: 'caption', text: '1 in = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.' },
        { style: 'caption', text: 'Points and millimetres are rounded to four decimals.' },
      ]),
    ]);
  },
};
