import { XLSX } from '../lib/grid.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const SHEET = 'xlsx-012';
const first = FIRST_ROW + 1;
const SERIES = A_SERIES.slice(0, 5);
const last = first + SERIES.length - 1;
const sumWidths = SERIES.reduce((a, [, w]) => a + w, 0);

// [label, formula, cached value]
const FORMULAS = [
  ['=SQRT(2)', 'SQRT(2)', Math.sqrt(2)],
  ['=PI()', 'PI()', Math.PI],
  ['=SUM(Widths)', 'SUM(Widths)', sumWidths],
  ['=10*Inch', '10*Inch', 254],
  [`=C${first}*D${first}/10^6 (A0, m²)`, `C${first}*D${first}/10^6`, (SERIES[0][1] * SERIES[0][2]) / 1e6],
  ['="A"&4 (text result)', '"A"&4', 'A4'],
];

export default {
  id: 'xlsx-012',
  slug: 'formulas-names',
  format: 'xlsx',
  title: 'Formulas and defined names',
  covers: [
    'xl/worksheets/sheet1.xml|x:c>x:f',
    'xl/worksheets/sheet1.xml|x:c@t=str',
    'xl/workbook.xml|x:workbook>x:definedNames',
    'xl/workbook.xml|x:definedNames>x:definedName',
    'xl/workbook.xml|x:definedName@name',
  ],
  parts() {
    const s = sheet({
      id: SHEET,
      title: 'Formulas and defined names',
      verify: 'Verify: column G shows each formula’s cached value; Widths names C7:C11 and Inch is the constant 25.4.',
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MODULE], ['E', MARGIN_COL], ['F', 2 * MODULE], ['G', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Width mm', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${FIRST_ROW}`, 'Height mm', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`F${FIRST_ROW}`, 'Formula', { font: 'captionBold' });
    s.set(`G${FIRST_ROW}`, 'Value', { font: 'captionBold', align: { horizontal: 'right' } });
    SERIES.forEach(([name, w, h], i) => s.set(`B${first + i}`, name).set(`C${first + i}`, w).set(`D${first + i}`, h));
    FORMULAS.forEach(([label, f, v], i) => {
      s.set(`F${first + i}`, label);
      s.set(`G${first + i}`, { f, v }, typeof v === 'number' && !Number.isInteger(v) ? { numFmt: '0.0000' } : { align: { horizontal: 'right' } });
    });
    s.definedName('Inch', '25.4');
    s.definedName('Widths', `'${SHEET}'!$C$${first}:$C$${last}`);
    return s.build();
  },
};
