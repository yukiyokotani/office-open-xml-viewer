import { XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
// The A4 page margins of the DOCX grid, in inches as SpreadsheetML stores them.
const MM = { left: 25, right: 20, top: 20, bottom: 23, header: 10, footer: 10 };
const inch = (mm) => mm / 25.4;
const SETTINGS = [
  ['pageMargins left', `${MM.left} mm`, inch(MM.left)],
  ['pageMargins right', `${MM.right} mm`, inch(MM.right)],
  ['pageMargins top', `${MM.top} mm`, inch(MM.top)],
  ['pageMargins bottom', `${MM.bottom} mm`, inch(MM.bottom)],
  ['pageMargins header, footer', `${MM.header} mm`, inch(MM.header)],
];

export default {
  id: 'xlsx-011',
  slug: 'page-setup',
  format: 'xlsx',
  title: 'Page setup',
  covers: [
    'xl/worksheets/sheet1.xml|x:worksheet>x:pageMargins',
    'xl/worksheets/sheet1.xml|x:pageMargins@left',
    'xl/worksheets/sheet1.xml|x:worksheet>x:pageSetup',
    'xl/worksheets/sheet1.xml|x:pageSetup@paperSize',
    'xl/worksheets/sheet1.xml|x:pageSetup@orientation=portrait',
    'xl/worksheets/sheet1.xml|x:worksheet>x:printOptions',
    'xl/worksheets/sheet1.xml|x:printOptions@gridLines',
    'xl/worksheets/sheet1.xml|x:worksheet>x:headerFooter',
    'xl/worksheets/sheet1.xml|x:headerFooter>x:oddHeader',
    'xl/worksheets/sheet1.xml|x:headerFooter>x:oddFooter',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-011',
      title: 'Page setup',
      verify: 'Verify (print or page view): A4 portrait, the margins listed below, gridlines printed, header and footer as stated.',
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE]],
      gridLines: true,
    });
    s.set(`B${FIRST_ROW}`, 'Setting', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Value', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${FIRST_ROW}`, 'Inches', { font: 'captionBold', align: { horizontal: 'right' } });
    SETTINGS.forEach(([label, value, inches], i) => {
      const r = FIRST_ROW + 1 + i;
      s.set(`B${r}`, label).set(`C${r}`, value, { align: { horizontal: 'right' } }).set(`D${r}`, inches, { numFmt: '0.0000' });
    });
    const r = FIRST_ROW + 1 + SETTINGS.length;
    s.set(`B${r}`, 'pageSetup paperSize').set(`C${r}`, '9 (A4)', { align: { horizontal: 'right' } });
    s.set(`B${r + 1}`, 'pageSetup orientation').set(`C${r + 1}`, 'portrait', { align: { horizontal: 'right' } });
    s.set(`B${r + 2}`, 'printOptions gridLines').set(`C${r + 2}`, 'on', { align: { horizontal: 'right' } });
    s.set(`B${r + 4}`, 'Header: “xlsx-011 Page setup” left, the sheet name right. Footer: “Page n of m” centred.', { font: 'caption' });
    s.after('<printOptions gridLines="1"/>'
      + `<pageMargins left="${inch(MM.left)}" right="${inch(MM.right)}" top="${inch(MM.top)}" bottom="${inch(MM.bottom)}" header="${inch(MM.header)}" footer="${inch(MM.footer)}"/>`
      + '<pageSetup paperSize="9" orientation="portrait"/>'
      + '<headerFooter><oddHeader>&amp;L&amp;8xlsx-011 Page setup&amp;R&amp;8&amp;A</oddHeader><oddFooter>&amp;C&amp;8Page &amp;P of &amp;N</oddFooter></headerFooter>');
    return s.build();
  },
};
