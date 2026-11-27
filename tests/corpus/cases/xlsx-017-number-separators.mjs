import { XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;

// ECMA-376 §18.8.30 stores `.` and `,` in a format code whatever the
// author's locale: they stand for the decimal and group separators of the
// reader's regional settings. Every row paints at least one of them; the last
// one also has a dot inside a quoted literal, which stays a dot.
// [quantity, value, format code; null for General]
const ROWS = [
  ['Inches per metre', 39.37007874, null],
  ['Golden ratio', 1.6180339887, null],
  ['Metres per mile', 1609.344, '#,##0.000'],
  ['Seconds in 365 days', 31536000, '#,##0'],
  ['A4 as a share of A0', 0.0625, '0.00%'],
  ['Avogadro constant, per mol', 6.02214076e23, '0.000E+00'],
  ['Litres per US gallon', 3.785411784, '0.00" l."'],
];

export default {
  id: 'xlsx-017',
  slug: 'number-separators',
  format: 'xlsx',
  title: 'Number separators',
  covers: [
    'xl/styles.xml|x:numFmts>x:numFmt',
    'xl/styles.xml|x:numFmt@formatCode',
    'xl/worksheets/sheet1.xml|x:c@s',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-017',
      title: 'Number separators',
      verify: 'Verify: . and , of a code (and General) paint the host decimal and group separators; the dot of " l." stays a dot.',
      // Column E is a gutter between the flush-right numbers and the codes.
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE], ['E', MARGIN_COL], ['F', 2 * MODULE]],
    });
    const head = FIRST_ROW;
    s.set(`B${head}`, 'Quantity', { font: 'captionBold' });
    s.set(`D${head}`, 'Displayed', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`F${head}`, 'Format code', { font: 'captionBold' });
    ROWS.forEach(([label, value, code], i) => {
      const r = head + 1 + i;
      s.set(`B${r}`, label);
      s.set(`D${r}`, value, code ? { numFmt: code } : {});
      s.set(`F${r}`, code ?? 'General');
    });
    return s.build();
  },
};
