import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { r } from '../lib/wml.mjs';

// What a column break leaves behind. Each case is a two-column section
// (continuous) between captions:
//   A. A paragraph holding only <w:br w:type="column"/>: its mark becomes an
//      empty line at the top of column 2, above "after".
//   B. Text that ends with a column break.
//   C. As A, with w:lastRenderedPageBreak in the paragraph after the break.
//   D. Page 2 (after a next-page section break) starts with a paragraph
//      that has w:pageBreakBefore: no extra page.
const { GUTTER } = DOCX;
const COLS = { num: 2, space: GUTTER };
const BR = '<w:r><w:br w:type="column"/></w:r>';
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, height: NaN, ...extra });
const para = (runs, extra = {}) => ({ kind: 'p', runs: [runs], height: NaN, ...extra });

export default {
  id: 'docx-059',
  slug: 'column-break-mark',
  format: 'docx',
  title: 'Column break and its mark',
  covers: [
    'word/document.xml|w:br@w:type=column',
    'word/document.xml|w:r>w:lastRenderedPageBreak',
    'word/document.xml|w:pPr>w:pageBreakBefore',
    'word/document.xml|w:type@w:val=continuous',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-059',
      title: 'Column break and its mark',
      verify: [
        'Verify: in A and C column 2 starts with an empty line',
        'above "after"; in B the mark line follows the break;',
        'D adds no blank page. Compare the PDF.',
      ],
    }),
    caption('A. A paragraph holding only a column break', { at: DOCX.rowLine(2), before: 0 }),
    { kind: 'section', type: 'continuous', cols: COLS },
    para(r('A before the break, column 1.')),
    para(BR),
    para(r('A after, column 2.')),
    { kind: 'section', type: 'continuous' },
    caption('B. Text ending with a column break'),
    { kind: 'section', type: 'continuous', cols: COLS },
    para(r('B text, then the break.') + BR),
    para(r('B after, column 2.')),
    { kind: 'section', type: 'continuous' },
    caption('C. As A, with lastRenderedPageBreak after'),
    { kind: 'section', type: 'continuous', cols: COLS },
    para(r('C before the break, column 1.')),
    para(BR),
    para('<w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">C after, column 2.</w:t></w:r>'),
    { kind: 'section', type: 'continuous' },
    caption('D follows on page 2, after a next-page section break.'),
    { kind: 'section', type: 'nextPage' },
    { kind: 'raw', height: NaN, xml: `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${r('D. Page 2 starts with this pageBreakBefore paragraph.')}</w:p>` },
  ]),
};
