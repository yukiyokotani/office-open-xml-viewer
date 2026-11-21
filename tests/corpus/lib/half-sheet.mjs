// Specimens laid out at half a sheet (book fold, two pages per sheet). The
// page is A4 landscape; Word halves its long side, so the usual header
// furniture (laid out on the 165 mm grid) would not fit: the id, title and
// what to verify are plain paragraphs, and each page names itself.

import { DOCX } from './grid.mjs';
import { docxPackage } from './docx.mjs';

const TEXT = [
  'The A series of paper sizes starts from a sheet of one square metre.',
  'Each size is half of the one before it, cut across the long side.',
  'The ratio of the sides stays the square root of two at every size.',
];

export function halfSheetCase({ id, slug, title, setting, covers }) {
  const page = (n) => [
    { kind: 'p', style: 'caption', bold: true, lines: [`Page ${n} of 4`], height: NaN, pageBreakBefore: n > 1 },
    ...TEXT.map((text) => ({ kind: 'p', text })),
  ];
  return {
    id,
    slug,
    format: 'docx',
    title,
    covers: [...covers, 'word/settings.xml|w:settings>w:mirrorMargins', 'word/document.xml|w:pgSz@w:orient=landscape'],
    parts: () => docxPackage([
      { kind: 'p', style: 'caption', bold: true, lines: [id] },
      { kind: 'p', style: 'caption', lines: [title] },
      { kind: 'p', style: 'caption', lines: ['Verify: four half-sheet pages, mirrored margins,', 'laid out as in the PDF.'], after: 1 },
      ...[1, 2, 3, 4].flatMap(page),
    ], {
      firstSection: { pgSz: { w: DOCX.PAGE_H, h: DOCX.PAGE_W, orient: 'landscape' } },
      settings: `<w:mirrorMargins/>${setting}`,
    }),
  };
}
