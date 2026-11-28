import { halfSheetCase } from '../lib/half-sheet.mjs';

// Two pages per sheet (w:printTwoOnOne with w:mirrorMargins) on an A4
// landscape sheet: Word lays out each page at half the sheet, the long side
// halved. Four pages, each named on its first line.
export default halfSheetCase({
  id: 'docx-061',
  slug: 'two-pages-per-sheet',
  title: 'Two pages per sheet',
  setting: '<w:printTwoOnOne/>',
  covers: ['word/settings.xml|w:settings>w:printTwoOnOne'],
});
