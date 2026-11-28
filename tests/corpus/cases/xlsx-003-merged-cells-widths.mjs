import { XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;

// Half, one, one and a half and two column modules.
const WIDTHS = [['B', MODULE / 2], ['C', MODULE], ['D', 1.5 * MODULE], ['E', 2 * MODULE]];

export default {
  id: 'xlsx-003',
  slug: 'merged-cells-widths',
  format: 'xlsx',
  title: 'Merged cells and column widths',
  covers: [
    'xl/worksheets/sheet1.xml|x:cols>x:col',
    'xl/worksheets/sheet1.xml|x:col@width',
    'xl/worksheets/sheet1.xml|x:col@customWidth',
    'xl/worksheets/sheet1.xml|x:sheetFormatPr@defaultColWidth',
    'xl/worksheets/sheet1.xml|x:mergeCells>x:mergeCell',
    'xl/worksheets/sheet1.xml|x:mergeCell@ref',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-003',
      title: 'Merged cells and column widths',
      verify: 'Verify: columns B to E are 6, 12, 18 and 24 characters wide; each merged range spans the cells it names.',
      cols: [['A', MARGIN_COL], ...WIDTHS],
      gridLines: true,
    });
    let r = FIRST_ROW;
    s.set(`B${r}`, 'Column width in characters', { font: 'captionBold' });
    r += 1;
    for (const [col, width] of WIDTHS) s.set(`${col}${r}`, width);

    r += 2;
    s.set(`B${r}`, 'Merged across columns: range and width in characters', { font: 'captionBold' });
    const spans = [['C', 18], ['D', 36], ['E', 60]];
    for (const [last, sum] of spans) {
      r += 1;
      s.set(`B${r}`, `B${r}:${last}${r} · ${sum}`).merge(`B${r}:${last}${r}`);
    }

    r += 2;
    s.set(`B${r}`, 'Merged across columns and rows: text at the default bottom left', { font: 'captionBold' });
    const top = r + 1;
    const range = `C${top}:D${top + 2}`;
    s.set(`C${top}`, `${range} · 30`).merge(range);
    return s.build();
  },
};
