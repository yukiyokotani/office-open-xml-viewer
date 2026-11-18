import { COLOR, XLSX } from '../lib/grid.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const first = FIRST_ROW + 1;
const last = first + A_SERIES.length - 1;

export default {
  id: 'xlsx-006',
  slug: 'conditional-formatting',
  format: 'xlsx',
  title: 'Conditional formatting',
  covers: [
    'xl/worksheets/sheet1.xml|x:worksheet>x:conditionalFormatting',
    'xl/worksheets/sheet1.xml|x:conditionalFormatting@sqref',
    'xl/worksheets/sheet1.xml|x:conditionalFormatting>x:cfRule',
    'xl/worksheets/sheet1.xml|x:cfRule@type=cellIs',
    'xl/worksheets/sheet1.xml|x:cfRule@operator=greaterThan',
    'xl/worksheets/sheet1.xml|x:cfRule@rank',
    'xl/worksheets/sheet1.xml|x:cfRule@type=expression',
    'xl/worksheets/sheet1.xml|x:cfRule>x:formula',
    'xl/worksheets/sheet1.xml|x:cfRule@dxfId',
    'xl/styles.xml|x:dxfs>x:dxf',
    'xl/styles.xml|x:dxf>x:fill',
    'xl/styles.xml|x:dxf>x:font',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-006',
      title: 'Conditional formatting',
      verify: 'Verify: widths over 300 white on cyan; the three largest heights bold; the A4 label cyan.',
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MODULE], ['E', MARGIN_COL], ['F', 3 * MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Width mm', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${FIRST_ROW}`, 'Height mm', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`F${FIRST_ROW}`, 'Rule', { font: 'captionBold' });
    A_SERIES.forEach(([name, w, h], i) => {
      s.set(`B${first + i}`, name).set(`C${first + i}`, w).set(`D${first + i}`, h);
    });
    s.set(`F${first}`, `C: cellIs greaterThan 300, dxf fill ${COLOR.accent}, font ${COLOR.paper}`);
    s.set(`F${first + 1}`, 'D: top10 rank 3, dxf bold');
    s.set(`F${first + 2}`, `B: expression $C${first}=210, dxf font ${COLOR.accent}`);
    const filled = s.styles.dxf(`<font><color rgb="FF${COLOR.paper}"/></font><fill><patternFill><bgColor rgb="FF${COLOR.accent}"/></patternFill></fill>`);
    const bold = s.styles.dxf('<font><b/></font>');
    const accent = s.styles.dxf(`<font><color rgb="FF${COLOR.accent}"/></font>`);
    s.conditional(`<conditionalFormatting sqref="C${first}:C${last}"><cfRule type="cellIs" dxfId="${filled}" priority="1" operator="greaterThan"><formula>300</formula></cfRule></conditionalFormatting>`);
    s.conditional(`<conditionalFormatting sqref="D${first}:D${last}"><cfRule type="top10" dxfId="${bold}" priority="2" rank="3"/></conditionalFormatting>`);
    s.conditional(`<conditionalFormatting sqref="B${first}:B${last}"><cfRule type="expression" dxfId="${accent}" priority="3"><formula>$C${first}=210</formula></cfRule></conditionalFormatting>`);
    return s.build();
  },
};
