import { halfSheetCase } from '../lib/half-sheet.mjs';

// Book-fold printing (w:bookFoldPrinting with w:mirrorMargins) on an A4
// landscape sheet: Word lays out each page at half the sheet, the long side
// halved, with mirrored margins. Four pages, each named on its first line.
export default halfSheetCase({
  id: 'docx-060',
  slug: 'book-fold',
  title: 'Book-fold printing',
  setting: '<w:bookFoldPrinting/>',
  covers: ['word/settings.xml|w:settings>w:bookFoldPrinting'],
});
