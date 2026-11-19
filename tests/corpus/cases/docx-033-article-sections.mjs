import { DOCX } from '../lib/grid.mjs';
import { R_NS, WML, W, docxPackage, header, run } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

// Section breaks arranged the way article templates write them, which is
// where readers that take column settings from the wrong section fail:
// - every break sits in an empty paragraph of its own;
// - the first section states one column explicitly (w:cols without w:num)
//   and has a title page head; the two-column sections inherit the heads;
// - a continuous section whose first paragraph starts page 2 brings new
//   heads, its own title page and a page number restart;
// - the document's final (body-level) w:sectPr is the two-column one, so a
//   reader that applies the body sectPr to the whole document, or each
//   sectPr to the content after it, sets page 1 in two columns.
const COLS2 = { num: 2, space: DOCX.gutterAfter(2) };
const ONE = `<w:cols w:space="${DOCX.gutterAfter(2)}"/>`;

const RIGHT_TAB = `<w:tabs><w:tab w:val="right" w:pos="${DOCX.LIVE_W}"/></w:tabs>`;
const PAGE = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
const furniture = (tag, left) => XML_DECL + `<w:${tag} ${W}><w:p><w:pPr><w:pStyle w:val="Caption"/>${RIGHT_TAB}</w:pPr>`
  + `${run(left)}${run('\tPage ')}${PAGE}</w:p></w:${tag}>`;
const PARTS = {
  'word/header1.xml': furniture('hdr', 'Head of section 1, first page'),
  'word/header2.xml': furniture('hdr', 'Head of section 1, other pages'),
  'word/footer1.xml': furniture('ftr', 'Foot of section 1'),
  'word/header3.xml': furniture('hdr', 'Head of section 3, first page'),
  'word/header4.xml': furniture('hdr', 'Head of section 3, other pages'),
  'word/footer2.xml': furniture('ftr', 'Foot of section 3'),
};
const refs = (first, other, foot) => `<w:headerReference w:type="default" r:id="${other}"/><w:footerReference w:type="default" r:id="${foot}"/>`
  + `<w:headerReference w:type="first" r:id="${first}"/><w:footerReference w:type="first" r:id="${foot}"/>`;

const text = (t, before = 0) => ({ kind: 'p', text: t, before });
const heading = (t) => ({ kind: 'p', bold: true, lines: [t], before: 1 });
// The empty paragraph that carries a section break, one line below the text.
const marker = { kind: 'p', lines: [''], before: 1 };

const BODY_1 = [
  'Section 2 starts continuously below the abstract and sets two columns of 80 millimetres with a 5 millimetre gutter. '
    + 'It declares no heads of its own, so it inherits those of section 1. Its sectPr is the second one in the document and '
    + 'sits in the empty paragraph that ends it.',
  'The ISO 216 A series starts at A0, one square metre in area with sides in the ratio of the square root of two. '
    + 'Each following size halves the previous one across its long side, so A1 is 594 by 841 millimetres, A2 420 by 594, '
    + 'A3 297 by 420 and A4 210 by 297.',
  'Section 3 begins with a continuous break, but its first paragraph has a page break before it, so the section starts '
    + 'at the top of page 2. These columns are balanced above that break and end on the same line.',
];
const BODY_2 = [
  'Section 4 is the last section, so its properties are written in the body-level sectPr at the end of the document. '
    + 'It sets two columns again. Reading this sectPr as the layout of the whole document would set the abstract on page 1 '
    + 'in two columns as well.',
  'The B series sizes are the geometric means of adjacent A sizes: B0 measures 1000 by 1414 millimetres and B1 707 by 1000. '
    + 'The C series, used for envelopes, lies between the A and B sizes: C4 measures 229 by 324 millimetres and holds an '
    + 'unfolded A4 sheet.',
  'A column break starts this paragraph in the second column. The heads and feet on this page come from section 3: the first-page head, because section 3 sets a title page and '
    + 'begins here, and the page number 2 from its restart.',
];

export default {
  id: 'docx-033',
  slug: 'article-sections',
  format: 'docx',
  title: 'Article sections',
  covers: [
    'word/document.xml|w:pPr>w:sectPr',
    'word/document.xml|w:body>w:sectPr',
    'word/document.xml|w:type@w:val=continuous',
    'word/document.xml|w:sectPr>w:titlePg',
    'word/document.xml|w:sectPr>w:pgNumType',
    'word/document.xml|w:pgNumType@w:start',
    'word/document.xml|w:headerReference@w:type=first',
    'word/document.xml|w:cols@w:num',
    'word/document.xml|w:pPr>w:pageBreakBefore',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-033',
      title: 'Article sections',
      verify: [
        'Verify: page 1 sets the abstract in one column above',
        'two balanced columns; page 2 shows the first-page head',
        'of section 3, page number 2, then two columns again.',
      ],
    }),
    heading('Abstract'),
    text('This specimen arranges four sections the way article templates do. Section 1, this abstract, is one column wide '
      + 'and states so with a w:cols element that has no w:num. It sets a title page, so page 1 shows the first-page head, '
      + 'and starts page numbering at 1. Every section break is written in an empty paragraph of its own.'),
    marker,
    { kind: 'section', type: 'continuous', cols: COLS2 },
    heading('1 Two columns after the abstract'),
    ...BODY_1.map((t, i) => text(t, i ? 1 : 0)),
    marker,
    {
      kind: 'section', type: 'continuous', cols: { raw: ONE }, titlePg: true,
      refs: refs('rId5', 'rId6', 'rId8'), pgNumType: '<w:pgNumType w:start="2"/>',
      headerDist: DOCX.HEADER_DIST, footerDist: DOCX.FOOTER_DIST,
    },
    { kind: 'p', style: 'title', lines: ['Part two'], pPrMid: '<w:pageBreakBefore/>' },
    text('Section 3 is one column again, holds this title and ends with the empty paragraph below.'),
    marker,
    { kind: 'section', type: 'continuous', cols: COLS2 },
    heading('2 Two columns to the end'),
    ...BODY_2.map((t, i) => (i === 2 ? { ...text(t), columnBreakBefore: true } : text(t, i ? 1 : 0))),
  ], {
    ns: [R_NS],
    firstSection: {
      refs: refs('rId3', 'rId4', 'rId7'), titlePg: true, cols: { raw: ONE },
      pgNumType: '<w:pgNumType w:start="1"/>',
      headerDist: DOCX.HEADER_DIST, footerDist: DOCX.FOOTER_DIST,
    },
    parts: PARTS,
    rels: [
      ['rId3', 'header', 'header1.xml'], ['rId4', 'header', 'header2.xml'],
      ['rId5', 'header', 'header3.xml'], ['rId6', 'header', 'header4.xml'],
      ['rId7', 'footer', 'footer1.xml'], ['rId8', 'footer', 'footer2.xml'],
    ],
    overrides: Object.fromEntries(Object.keys(PARTS).map((p) => [p, `${WML}.${p.includes('header') ? 'header' : 'footer'}+xml`])),
  }),
};
