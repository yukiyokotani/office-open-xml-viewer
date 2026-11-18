import { COLOR, XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;

export default {
  id: 'xlsx-013',
  slug: 'rich-text',
  format: 'xlsx',
  title: 'Rich text and font styles',
  covers: [
    'xl/sharedStrings.xml|x:si>x:r',
    'xl/sharedStrings.xml|x:r>x:rPr',
    'xl/sharedStrings.xml|x:rPr>x:b',
    'xl/sharedStrings.xml|x:rPr>x:i',
    'xl/sharedStrings.xml|x:rPr>x:u',
    'xl/sharedStrings.xml|x:rPr>x:color',
    'xl/sharedStrings.xml|x:rPr>x:rFont',
    'xl/styles.xml|x:font>x:i',
    'xl/styles.xml|x:font>x:u',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-013',
      title: 'Rich text and font styles',
      verify: 'Verify: row 7 changes style run by run inside one cell; rows 9 and 10 are italic and underlined as whole cells.',
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', 4 * MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'Source', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Sample', { font: 'captionBold' });
    s.set(`B${FIRST_ROW + 1}`, 'shared string runs (x:r)');
    s.set(`C${FIRST_ROW + 1}`, {
      rich: [['One cell, four runs: '], ['bold', { b: true }], [', '], ['italic', { i: true }], [', '], ['underlined', { u: true }],
        [' and '], ['orange', { color: COLOR.accent }], ['.']],
    });
    s.set(`B${FIRST_ROW + 3}`, 'cell font i');
    s.set(`C${FIRST_ROW + 3}`, 'The whole cell is italic.', { italic: true });
    s.set(`B${FIRST_ROW + 4}`, 'cell font u');
    s.set(`C${FIRST_ROW + 4}`, 'The whole cell is underlined.', { underline: true });
    return s.build();
  },
};
