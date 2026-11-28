import { XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW, ROW_PT } = XLSX;
const label = (s, r, text) => s.set(`B${r}`, text);

export default {
  id: 'xlsx-004',
  slug: 'alignment',
  format: 'xlsx',
  title: 'Alignment',
  covers: [
    'xl/styles.xml|x:xf>x:alignment',
    'xl/styles.xml|x:alignment@horizontal=center',
    'xl/styles.xml|x:alignment@vertical=top',
    'xl/styles.xml|x:alignment@vertical=center',
    'xl/styles.xml|x:alignment@wrapText',
    'xl/styles.xml|x:alignment@indent',
    'xl/styles.xml|x:alignment@textRotation',
    'xl/worksheets/sheet1.xml|x:row@ht',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-004',
      title: 'Alignment',
      verify: 'Verify: each sample sits where its own text says, inside the cell outlined by the gridlines.',
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE], ['E', MODULE]],
      gridLines: true,
    });
    let r = FIRST_ROW;
    s.set(`B${r}`, 'Property and values', { font: 'captionBold' });
    r += 1;
    label(s, r, 'horizontal');
    s.set(`C${r}`, 'left', { align: { horizontal: 'left' } });
    s.set(`D${r}`, 'center', { align: { horizontal: 'center' } });
    s.set(`E${r}`, 'right', { align: { horizontal: 'right' } });
    r += 2;
    label(s, r, 'vertical, row 36 pt');
    s.height(r, 3 * ROW_PT);
    s.set(`C${r}`, 'top', { align: { vertical: 'top' } });
    s.set(`D${r}`, 'center', { align: { vertical: 'center' } });
    s.set(`E${r}`, 'bottom (default)');
    r += 2;
    label(s, r, 'wrapText, row 48 pt');
    s.height(r, 4 * ROW_PT);
    s.set(`C${r}`, 'Wrapped text breaks at the cell edge.', { align: { vertical: 'top', wrapText: 1 } });
    r += 2;
    label(s, r, 'indent 1 and 2');
    s.set(`C${r}`, 'indent 1', { align: { indent: 1 } });
    s.set(`D${r}`, 'indent 2', { align: { indent: 2 } });
    r += 2;
    label(s, r, 'textRotation 90, 45, 255');
    s.height(r, 6 * ROW_PT);
    s.set(`C${r}`, 'up 90°', { align: { textRotation: 90 } });
    s.set(`D${r}`, 'up 45°', { align: { textRotation: 45 } });
    s.set(`E${r}`, 'stack', { align: { textRotation: 255 } });
    return s.build();
  },
};
