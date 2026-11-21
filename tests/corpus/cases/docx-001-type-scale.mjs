import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;

export default {
  id: 'docx-001',
  slug: 'type-scale',
  format: 'docx',
  title: 'Type scale',
  covers: [
    'word/styles.xml|w:spacing@w:line',
    'word/styles.xml|w:spacing@w:lineRule=exact',
    'word/styles.xml|w:rPr>w:sz',
    'word/styles.xml|w:rPr>w:b',
    'word/document.xml|w:pPr>w:pStyle',
    'word/document.xml|w:spacing@w:before',
    'word/document.xml|w:spacing@w:after',
    'word/document.xml|w:rPr>w:b',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-001',
      title: 'Type scale',
      verify: [
        'Verify: sizes 8/12, 9/12 and 24/24 pt in two',
        'weights; every line on the 12 pt baseline grid.',
      ],
    }),

    { kind: 'p', at: at(2), style: 'caption', bold: true, cols: [1, 2], lines: ['Title: 24 pt type on a 24 pt line, bold'] },
    { kind: 'p', style: 'title', cols: [1, 4], lines: ['Twenty-four point'] },
    { kind: 'p', before: 1, style: 'caption', bold: true, cols: [1, 2], lines: ['Text: 9 pt type on a 12 pt line, regular'] },
    {
      kind: 'p',
      cols: [1, 2],
      lines: [
        'This paragraph is set in 9 point Noto Sans on a',
        '12 point line. The live area of the page is',
        '165 by 254 mm and holds sixty such lines;',
        'Word records the leading as line 240, exact.',
      ],
    },

    { kind: 'p', at: at(3), style: 'caption', bold: true, cols: [1, 2], lines: ['Bold: 9 pt type on a 12 pt line, bold'] },
    {
      kind: 'p',
      bold: true,
      cols: [1, 2],
      lines: ['Bold is the second and last weight. It marks', 'labels; it is not used for running text.'],
    },
    { kind: 'p', before: 1, style: 'caption', bold: true, cols: [1, 2], lines: ['Caption: 8 pt type on a 12 pt line, regular'] },
    {
      kind: 'p',
      style: 'caption',
      cols: [1, 2],
      lines: [
        'Captions are 8 point and keep the 12 point line.',
        'Three sizes and two weights are the complete',
        'typographic means of every file in this corpus.',
      ],
    },

    { kind: 'p', at: at(4), style: 'caption', bold: true, cols: [1, 3], lines: ['Paragraph spacing: multiples of the 12 pt line'] },
    { kind: 'p', cols: [1, 3], lines: ['Space before 0 pt, space after 0 pt. This line starts on grid line 32.'] },
    { kind: 'p', before: 1, cols: [1, 3], lines: ['Space before 12 pt (before 240): one empty line above grid line 34.'] },
    { kind: 'p', before: 2, after: 1, cols: [1, 3], lines: ['Space before 24 pt (before 480) and after 12 pt (after 240).'] },
    { kind: 'p', cols: [1, 3], lines: ['This paragraph follows the space after and starts on grid line 39.'] },
  ]),
};
