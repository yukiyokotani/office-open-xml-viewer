import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

// Every way Word starts a new page or column, one per page. Each page names
// itself, so a missing or extra page is visible: an even-page section after
// page 4 makes Word insert a blank page 5; the odd-page section after page 6
// starts on page 7. A w:lastRenderedPageBreak (a hint Word writes where the
// page broke when it last saved) sits in the middle of page 1 and must not
// break the page.
const COLS = { num: 2, space: DOCX.gutterAfter(2) };
const page = (n, how) => ({ kind: 'p', bold: true, lines: [`Page ${n}`] , ...(how ?? {}) });
const say = (text, extra = {}) => ({ kind: 'p', text, before: 1, ...extra });

export default {
  id: 'docx-023',
  slug: 'breaks',
  format: 'docx',
  title: 'Breaks',
  covers: [
    'word/document.xml|w:br@w:type=page',
    'word/document.xml|w:br@w:type=column',
    'word/document.xml|w:pPr>w:pageBreakBefore',
    'word/document.xml|w:r>w:lastRenderedPageBreak',
    'word/document.xml|w:type@w:val=nextPage',
    'word/document.xml|w:type@w:val=evenPage',
    'word/document.xml|w:type@w:val=oddPage',
    'word/document.xml|w:type@w:val=continuous',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-023',
      title: 'Breaks',
      verify: [
        'Verify: seven pages, each naming itself; page 5 is blank',
        '(inserted for the even-page section); page 7 ends in two',
        'columns, the right one starting at a column break.',
      ],
    }),
    page(1, { before: 1 }),
    {
      kind: 'p',
      before: 1,
      runs: [
        run('This paragraph carries a w:lastRenderedPageBreak here'),
        '<w:r><w:lastRenderedPageBreak/></w:r>',
        run(', in the middle of its text. It records where Word broke the page when the file was saved and is not a break.'),
      ],
    },
    say('The next paragraph starts with a manual page break, w:br w:type="page".'),
    page(2, { pageBreakBefore: true }),
    say('The next paragraph has the paragraph property w:pageBreakBefore.'),
    { kind: 'raw', xml: '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Page 3</w:t></w:r></w:p>' },
    say('A section break (next page) follows this paragraph.'),
    { kind: 'section', type: 'nextPage' },
    page(4),
    say('A section break (even page) follows. Page 5 is odd, so Word leaves it blank and starts the section on page 6.'),
    { kind: 'section', type: 'evenPage' },
    page(6),
    say('A section break (odd page) follows. Page 7 is odd, so no blank page is needed.'),
    { kind: 'section', type: 'oddPage' },
    page(7),
    say('A continuous section break follows; the section after it has two columns.'),
    { kind: 'section', type: 'continuous', cols: COLS },
    say('Left column. A column break, w:br w:type="column", ends this paragraph.', { before: 0 }),
    say('Right column: this paragraph starts after the column break.', { before: 0, columnBreakBefore: true }),
  ]),
};
