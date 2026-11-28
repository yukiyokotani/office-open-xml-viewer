import { DOCX } from '../lib/grid.mjs';
import { WML, docxPackage, header, run } from '../lib/docx.mjs';
import { A_SERIES, B_SERIES } from '../lib/iso216.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';
import { W } from '../lib/docx.mjs';

// The furniture of a paper. A reference list numbered [1] to [12] by
// w:numbering with a tab after the number and a hanging indent of one
// gutter: one-digit numbers fit inside the hang, two-digit numbers do not, so
// their text starts at the next default tab stop instead. A code block keeps
// its spaces. At the foot of page 1 a heading with w:keepNext stands before a
// paragraph with w:keepLines that does not fit there, so both start page 2.
const { LINE, GUTTER } = DOCX;
const TAB = 2 * GUTTER;
const NUMBERING = XML_DECL + `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>`
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="[%1]"/><w:lvlJc w:val="left"/>'
  + `<w:pPr><w:ind w:left="${GUTTER}" w:hanging="${GUTTER}"/></w:pPr></w:lvl></w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const REF = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
const REFERENCES = [
  'ISO 216, Writing paper and certain classes of printed matter: trimmed sizes, A and B series.',
  'ISO 269, Correspondence envelopes: designation and sizes.',
  'ECMA-376, Office Open XML File Formats, Part 1: Fundamentals and Markup Language Reference.',
  'ECMA-376, Part 2: Open Packaging Conventions.',
  'Unicode Standard Annex #9, Unicode Bidirectional Algorithm.',
  'Unicode Standard Annex #14, Unicode Line Breaking Algorithm.',
  'JIS X 4051, Formatting rules for Japanese documents.',
  'ISO 8601, Date and time: representations for information interchange.',
  'ISO 80000-2, Quantities and units: mathematics.',
  'ISO/IEC 14496-22, Open Font Format.',
  'ISO 32000-2, Document management: Portable Document Format.',
  'IEC 60027-2, Letter symbols to be used in electrical technology: telecommunications and electronics.',
];
const CODE = [
  'function halve([w, h]) {',
  '    // A sheet folded across its long side',
  '    return [Math.floor(h / 2), w];',
  '}',
  '',
  'halve([210, 297]);   // → [148, 210]',
];
const CODE_BOX = `<w:pBdr><w:left w:val="single" w:sz="12" w:space="${GUTTER / 20}" w:color="000000"/></w:pBdr>`;
const SIZES = [...A_SERIES, ...B_SERIES];
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const KEEP_PARAGRAPH = 'This paragraph carries w:keepLines and is several lines long, so it cannot be split between the foot of '
  + 'page 1 and the head of page 2. No line is left on page 1 below the heading, so the paragraph starts page 2; the '
  + 'heading above it carries w:keepNext, so it moves to page 2 as well and stands above the paragraph it introduces. '
  + 'Page 1 therefore ends with the list of sizes.';

export default {
  id: 'docx-028',
  slug: 'paper-furniture',
  format: 'docx',
  title: 'Paper furniture',
  covers: [
    'word/numbering.xml|w:lvl>w:lvlText',
    'word/numbering.xml|w:lvlJc@w:val=left',
    'word/document.xml|w:pPr>w:numPr',
    'word/document.xml|w:pPr>w:keepNext',
    'word/document.xml|w:pPr>w:keepLines',
    'word/settings.xml|w:settings>w:defaultTabStop',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-028',
      title: 'Paper furniture',
      verify: [
        'Verify: [1]-[9] start their text one gutter in, [10]-[12] at',
        'the next default tab stop; the code keeps its indents; page 1',
        'ends after B10 and page 2 starts with heading D.',
      ],
    }),
    caption('A. Code block: spaces are kept, a left rule marks the block'),
    { kind: 'p', pBdr: CODE_BOX, lines: CODE.map((l) => l || ' ') },
    caption('B. References: [%1] with a hanging indent of one gutter'),
    ...REFERENCES.map((text) => ({ kind: 'p', numPr: REF, lines: [text] })),
    caption('C. Sizes; heading D is written on the last line of page 1'),
    ...SIZES.map(([n, w, h]) => ({ kind: 'p', tabs: [['left', DOCX.colStart(2)]], lines: [`${n}\t${w} × ${h} mm`] })),
    {
      kind: 'raw',
      at: DOCX.LINES, // the last line of page 1
      height: 1,
      xml: '<w:p><w:pPr><w:keepNext/><w:spacing w:before="0"/></w:pPr>'
        + `${run('D. A heading kept with the next paragraph', '<w:b/>')}</w:p>`,
    },
    { kind: 'raw', xml: `<w:p><w:pPr><w:keepLines/></w:pPr>${run(KEEP_PARAGRAPH)}</w:p>` },
  ], {
    settings: `<w:defaultTabStop w:val="${TAB}"/>`,
    parts: { 'word/numbering.xml': NUMBERING },
    rels: [['rId3', 'numbering', 'numbering.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml` },
  }),
};
