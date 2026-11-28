import { XLSX } from '../lib/grid.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const first = FIRST_ROW + 1;
// Rows of A2 and A3, and column D, are hidden.
const HIDDEN_ROWS = [first + 2, first + 3];

export default {
  id: 'xlsx-008',
  slug: 'frozen-hidden',
  format: 'xlsx',
  title: 'Frozen panes and hidden cells',
  covers: [
    'xl/worksheets/sheet1.xml|x:sheetView>x:pane',
    'xl/worksheets/sheet1.xml|x:pane@state=frozen',
    'xl/worksheets/sheet1.xml|x:pane@xSplit',
    'xl/worksheets/sheet1.xml|x:pane@ySplit',
    'xl/worksheets/sheet1.xml|x:pane@topLeftCell',
    'xl/worksheets/sheet1.xml|x:sheetView>x:selection',
    'xl/worksheets/sheet1.xml|x:row@hidden',
    'xl/worksheets/sheet1.xml|x:col@hidden',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-008',
      title: 'Frozen panes and hidden cells',
      verify: `Verify: A2, A3 (rows ${HIDDEN_ROWS.join(', ')}) and column D are hidden; rows 1-${FIRST_ROW} and columns A-B stay in place when scrolling.`,
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MODULE, true], ['E', MODULE]],
      gridLines: true,
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Width mm', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`D${FIRST_ROW}`, 'Hidden', { font: 'captionBold', align: { horizontal: 'right' } });
    s.set(`E${FIRST_ROW}`, 'Height mm', { font: 'captionBold', align: { horizontal: 'right' } });
    A_SERIES.forEach(([name, w, h], i) => {
      s.set(`B${first + i}`, name).set(`C${first + i}`, w).set(`D${first + i}`, w * h).set(`E${first + i}`, h);
    });
    for (const r of HIDDEN_ROWS) s.hideRow(r);
    s.pane(`<pane xSplit="2" ySplit="${FIRST_ROW}" topLeftCell="C${first}" activePane="bottomRight" state="frozen"/>`
      + '<selection pane="topRight"/><selection pane="bottomLeft"/>'
      + `<selection pane="bottomRight" activeCell="C${first}" sqref="C${first}"/>`);
    return s.build();
  },
};
