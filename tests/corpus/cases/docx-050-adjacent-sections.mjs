import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Continuous section breaks mid-page at one exact page geometry, the one a
// common Japanese template uses: A4, margins 851 twips (15 mm) on every
// side, header 851 and footer 992 from the edge, and two columns 425 twips
// (7.5 mm) apart. The geometry is the subject, so this specimen is off the
// corpus grid on purpose: 851 twips is 42.55 pt, and a layout that adds such
// lengths in floating point can leave a section boundary a fraction of a
// point off.
//   A. One-column section, a few paragraphs.
//   B. A continuous break into two columns that share the page with A:
//      the columns start right under A, on page 1.
//   C. A continuous break back to one column, under the balanced columns.
const GEOMETRY = { top: 851, right: 851, bottom: 851, left: 851, header: 851, footer: 992 };
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const para = (text) => ({ kind: 'p', runs: [`<w:r>${t(text)}</w:r>`], height: NaN });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const TEXT = [
  'The A series of paper sizes starts from a sheet of one square metre.',
  'Each size is half of the one before it, cut across the long side.',
  'The ratio of the sides stays the square root of two at every size.',
  'A4 is the sixteenth part of A0, and A5 is half of A4.',
  'The B series sits between the A sizes, and the C series fits envelopes.',
  'Envelopes of the C series hold the matching A size flat or folded.',
];

export default {
  id: 'docx-050',
  slug: 'adjacent-sections',
  format: 'docx',
  title: 'Adjacent sections, exact geometry',
  covers: [
    'word/document.xml|w:type@w:val=continuous',
    'word/document.xml|w:sectPr>w:cols',
    'word/document.xml|w:cols@w:space',
    'word/document.xml|w:sectPr>w:pgMar',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-050',
      title: 'Adjacent sections, exact geometry',
      verify: [
        'Verify: everything is on page 1; B two columns start',
        'right under A and are balanced; C one column follows',
        'under them. Margins 15 mm, column gap 7.5 mm.',
      ],
    }),
    caption('A. One column', { at: DOCX.rowLine(2), before: 0 }),
    ...TEXT.slice(0, 3).map(para),
    { kind: 'section', type: 'continuous', pgMar: GEOMETRY, cols: { num: 2, space: 425 } },
    caption('B. Two columns after a continuous break'),
    ...TEXT.map(para),
    ...TEXT.map(para),
    { kind: 'section', type: 'continuous', pgMar: GEOMETRY },
    caption('C. One column again'),
    ...TEXT.slice(3).map(para),
  ], {
    firstSection: { pgMar: GEOMETRY },
  }),
};
