import { COLOR, XLSX } from '../lib/grid.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const first = FIRST_ROW + 1;
const last = first + A_SERIES.length - 1;

export default {
  id: 'xlsx-007',
  slug: 'data-bars-colour-scale',
  format: 'xlsx',
  title: 'Data bars and colour scale',
  covers: [
    'xl/worksheets/sheet1.xml|x:cfRule@type=dataBar',
    'xl/worksheets/sheet1.xml|x:cfRule>x:dataBar',
    'xl/worksheets/sheet1.xml|x:dataBar>x:cfvo',
    'xl/worksheets/sheet1.xml|x:cfvo@type=min',
    'xl/worksheets/sheet1.xml|x:cfvo@type=max',
    'xl/worksheets/sheet1.xml|x:dataBar>x:color',
    'xl/worksheets/sheet1.xml|x:cfRule@type=colorScale',
    'xl/worksheets/sheet1.xml|x:cfRule>x:colorScale',
    'xl/worksheets/sheet1.xml|x:colorScale>x:color',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-007',
      title: 'Data bars and colour scale',
      verify: 'Verify: bar length grows with width (A0 longest); height fill runs from white (A10) to cyan (A0).',
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', 2 * MODULE], ['D', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Width mm, data bar', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${FIRST_ROW}`, 'Height mm', { font: 'captionBold', align: { horizontal: 'right' } });
    A_SERIES.forEach(([name, w, h], i) => {
      s.set(`B${first + i}`, name).set(`C${first + i}`, w).set(`D${first + i}`, h);
    });
    s.conditional(`<conditionalFormatting sqref="C${first}:C${last}"><cfRule type="dataBar" priority="1"><dataBar>`
      + `<cfvo type="min"/><cfvo type="max"/><color rgb="FF${COLOR.accent}"/></dataBar></cfRule></conditionalFormatting>`);
    s.conditional(`<conditionalFormatting sqref="D${first}:D${last}"><cfRule type="colorScale" priority="2"><colorScale>`
      + `<cfvo type="min"/><cfvo type="max"/><color rgb="FF${COLOR.paper}"/><color rgb="FF${COLOR.accent}"/></colorScale></cfRule></conditionalFormatting>`);
    s.set(`B${last + 2}`, 'Data bar: cfvo min to max, colour accent. Colour scale: two stops, min white, max accent.', { font: 'caption' });
    return s.build();
  },
};
