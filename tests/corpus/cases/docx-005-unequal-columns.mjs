import { DOCX } from '../lib/grid.mjs';
import { docxPackage, headerSection } from '../lib/docx.mjs';

// One grid column for notes, three for text: explicit widths on the grid.
const NARROW = DOCX.span(1, 1);
const WIDE = DOCX.span(2, 4);
const COLS = { widths: [NARROW, WIDE], space: DOCX.gutterAfter(1) };

export default {
  id: 'docx-005',
  slug: 'unequal-columns',
  format: 'docx',
  title: 'Unequal columns',
  covers: [
    'word/document.xml|w:cols@w:equalWidth',
    'word/document.xml|w:cols>w:col',
    'word/document.xml|w:col@w:w',
    'word/document.xml|w:col@w:space',
    'word/document.xml|w:br@w:type=column',
  ],
  parts: () => docxPackage([
    ...headerSection({
      id: 'docx-005',
      title: 'Unequal columns',
      verify: [
        `Verify: a ${NARROW} twip column beside a ${WIDE} twip column;`,
        'the text column starts at its top after the column break.',
      ],
    }, { cols: COLS }),
    { kind: 'p', style: 'caption', bold: true, lines: ['Column 1'] },
    { kind: 'p', style: 'caption', lines: [`w:col w:w ${NARROW}`, `w:space ${COLS.space}`, 'one grid column'] },
    { kind: 'p', before: 1, style: 'caption', lines: ['A column break', 'ends this note.'] },
    {
      kind: 'p',
      columnBreakBefore: true,
      text: 'The second column is three grid columns and two gutters wide. Unequal columns are written '
        + `with w:equalWidth 0 and one w:col element per column, each with its own width: here ${NARROW} `
        + `and ${WIDE} twentieths of a point, separated by a space of ${COLS.space}. Together they fill the `
        + `live area of ${DOCX.LIVE_W} exactly.`,
    },
    {
      kind: 'p',
      before: 1,
      text: 'The explicit column break (w:br with w:type column) at the start of the first paragraph of this '
        + 'column ends the note in the first column, so this text begins at the top of the second column '
        + 'however much room the first column has left. The narrow column suits short notes and labels; '
        + 'the wide one carries running text.',
    },
  ]),
};
