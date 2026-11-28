import { XLSX } from '../lib/grid.mjs';
import { BUILTIN_NUMFMT, sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;

// ECMA-376 §18.8.30 defines `.` and `,` in a format code as the decimal point
// and thousands separator tokens; [MS-OI29500] §2.1.712 states the grammar for
// en-US and has Office adjust it for international settings. The expected
// behavior, tracked by #2346, is that the viewer paints those tokens, and
// General, with the host locale's separators. This sheet is en-US coverage
// only: the VRT host is en-US, where the separators are the invariant ones
// whether or not the viewer localizes. The General, grouped-decimal and
// quoted-dot rows are the cases that a comma-decimal host tells apart (the
// dot inside " l." is literal text and stays a dot); a check under such a
// host belongs with the #2346 fix, not with this specimen.
// [quantity, value, number format: built-in id, format code, or null for General]
const ROWS = [
  ['Inches per metre', 39.37007874, null],
  ['Millimetres per inch', 25.4, '0.000'],
  ['Twips per inch', 1440, 3],
  ['EMU per inch', 914400, 3],
  ['Square root of 2', Math.SQRT2, '0.00000'],
  ['Area of A0, m²', 0.999949, '0.0000'],
  ['A4 as a share of A0', 0.0625, 10],
  ['Metres per mile', 1609.344, '#,##0.000'],
  ['Speed of light, m/s', 299792458, '0.000E+00'],
  ['Litres per US gallon', 3.785411784, '0.00" l."'],
  ['1970-01-01 as a date serial', 25569, 'yyyy-mm-dd'],
  ['Half a day', 0.5, 'hh:mm:ss'],
  ['Three quarters', 0.75, 12],
  ['Absolute zero, °C', -273.15, '#,##0.00;(#,##0.00)'],
];

export default {
  id: 'xlsx-001',
  slug: 'number-formats',
  format: 'xlsx',
  title: 'Number formats',
  covers: [
    'xl/styles.xml|x:numFmts>x:numFmt',
    'xl/styles.xml|x:numFmt@numFmtId',
    'xl/styles.xml|x:numFmt@formatCode',
    'xl/styles.xml|x:xf@numFmtId',
    'xl/styles.xml|x:xf@applyNumberFormat',
    'xl/worksheets/sheet1.xml|x:c@s',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-001',
      title: 'Number formats',
      verify: 'Verify: column D shows the value of column C through the format code in column F; numbers flush right.',
      // Column E is a gutter: it separates the flush-right numbers of D from the
      // flush-left codes of F.
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE], ['E', MARGIN_COL], ['F', 2 * MODULE]],
    });
    const head = FIRST_ROW;
    s.set(`B${head}`, 'Quantity', { font: 'captionBold' });
    s.set(`C${head}`, 'Stored value', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${head}`, 'Displayed', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`F${head}`, 'Format code', { font: 'captionBold' });
    ROWS.forEach(([label, value, numFmt], i) => {
      const r = head + 1 + i;
      const code = numFmt === null ? 'General'
        : typeof numFmt === 'number' ? `${BUILTIN_NUMFMT[numFmt]}  (built-in ${numFmt})` : numFmt;
      s.set(`B${r}`, label);
      s.set(`C${r}`, value);
      s.set(`D${r}`, value, numFmt === null ? {} : { numFmt });
      s.set(`F${r}`, code);
    });
    return s.build();
  },
};
