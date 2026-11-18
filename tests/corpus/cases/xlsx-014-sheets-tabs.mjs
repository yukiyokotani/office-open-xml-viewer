import { COLOR, XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const COLS = [['A', MARGIN_COL], ['B', 4 * MODULE]];

export default {
  id: 'xlsx-014',
  slug: 'sheets-tabs',
  format: 'xlsx',
  title: 'Sheets and tabs',
  covers: [
    'xl/workbook.xml|x:sheets>x:sheet',
    'xl/workbook.xml|x:sheet@state=hidden',
    'xl/workbook.xml|x:workbook>x:bookViews',
    'xl/workbook.xml|x:workbookView@activeTab',
    'xl/worksheets/sheet2.xml|x:worksheet>x:sheetPr',
    'xl/worksheets/sheet2.xml|x:sheetPr>x:tabColor',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-014',
      title: 'Sheets and tabs',
      verify: 'Verify: three sheets, the second with a cyan tab, the third hidden; this sheet is active.',
      cols: COLS,
    });
    s.bookView(' activeTab="0"');
    s.set(`B${FIRST_ROW}`, 'Sheet 1 (this sheet): workbookView activeTab 0.');
    s.set(`B${FIRST_ROW + 1}`, `Sheet 2 “Tab colour”: sheetPr/tabColor ${COLOR.accent}.`);
    s.set(`B${FIRST_ROW + 2}`, 'Sheet 3 “Hidden”: sheet state hidden; it has no tab.');
    s.addSheet({ name: 'Tab colour', cols: COLS }).sheetPr(`<tabColor rgb="FF${COLOR.accent}"/>`)
      .set('B2', 'xlsx-014, sheet 2', { font: 'captionBold' })
      .set(`B${FIRST_ROW}`, `This sheet's tab is coloured ${COLOR.accent} by sheetPr/tabColor.`);
    s.addSheet({ name: 'Hidden', cols: COLS }).state('hidden')
      .set('B2', 'xlsx-014, sheet 3', { font: 'captionBold' })
      .set(`B${FIRST_ROW}`, 'This sheet is hidden: the workbook lists it with state hidden.');
    return s.build();
  },
};
