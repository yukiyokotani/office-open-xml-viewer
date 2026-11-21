import { DOCX } from '../lib/grid.mjs';
import { docxPackage, headerSection } from '../lib/docx.mjs';

// Two equal text columns of two grid columns each; the space between them is
// the grid gutter after column 2.
const COLS = { num: 2, space: DOCX.gutterAfter(2) };
const COL_W = (DOCX.LIVE_W - COLS.space) / 2;

const PROSE = [
  'A column is a measure chosen before the text is set. This section is set in two equal '
    + `columns of 80 mm, each the width of two grid columns and the gutter between them; the space `
    + `between the two text columns is one gutter of 5 mm. Word stores the arrangement in the section `
    + `properties as w:cols with w:num ${COLS.num} and w:space ${COLS.space}, and derives each column `
    + `width from the live area: (${DOCX.LIVE_W} − ${COLS.space}) ÷ 2 = ${COL_W} twentieths of a point.`,
  'The section begins with a continuous section break, so it shares its page with the single-column '
    + 'header above. A continuous break ends the previous section without starting a new page.',
  'At 9 point, a line of 80 mm holds about fifty characters, within the range usually recommended '
    + 'for continuous reading. A wider measure asks the eye to travel further from the end of one line '
    + 'to the start of the next; a narrower one breaks words and phrases more often.',
  'The lines of these paragraphs are not broken by hand. Where they end is decided by the application '
    + 'from the column width, the font and the text. A second continuous section break follows the last '
    + 'paragraph, so the two columns are balanced and end on nearly the same line.',
];

export default {
  id: 'docx-004',
  slug: 'section-columns',
  format: 'docx',
  title: 'Section columns',
  covers: [
    'word/document.xml|w:sectPr>w:cols',
    'word/document.xml|w:cols@w:num',
    'word/document.xml|w:cols@w:space',
    'word/document.xml|w:sectPr>w:type',
    'word/document.xml|w:type@w:val=continuous',
    'word/document.xml|w:pPr>w:sectPr',
  ],
  parts: () => docxPackage([
    ...headerSection({
      id: 'docx-004',
      title: 'Section columns',
      verify: [
        'Verify: one column above, two balanced columns',
        `of ${COL_W} twips with a ${COLS.space} twip space, one column below.`,
      ],
    }, { cols: COLS }),
    ...PROSE.map((text, i) => ({ kind: 'p', text, before: i ? 1 : 0 })),
    { kind: 'section', type: 'continuous' },
    {
      kind: 'p',
      before: 1,
      style: 'caption',
      text: 'This caption is in a third section of one column, after the continuous break that balances the columns above.',
    },
  ]),
};
