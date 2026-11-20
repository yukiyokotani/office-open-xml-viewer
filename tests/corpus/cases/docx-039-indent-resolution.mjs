import { DOCX, FONT } from '../lib/grid.mjs';
import { WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';
import { W } from '../lib/docx.mjs';

// How Word resolves a paragraph's indents when several sources set them.
//   A. Numbering whose hanging indent equals its left indent, so the number
//      starts at the margin, with the three number suffixes: nothing, a
//      space and a tab. With nothing and a space the body follows the number
//      and continuation lines take the left indent; the tab case is the
//      control. Each item has a second line.
//   B. Character-unit indents with Word's twip sentinel: w:left and
//      w:firstLine of 31680 twips (22 inches) next to w:leftChars and
//      w:hangingChars. Word ignores the sentinel: the first line starts at
//      leftChars, the following lines at leftChars + hangingChars.
//   C. The indent cascade in a table cell: a style with leftChars 400, a
//      bulleted list level, and a paragraph w:leftChars="0" that cancels
//      the style's character indent. Four rows remove one layer at a time.
const { colStart, span } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const STEP = colStart(2);
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const rl = (lines, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${lines.map((x, i) => `${i ? '<w:br/>' : ''}${t(x)}`).join('')}</w:r>`;
const p = (runs, pPr = '', height = 1) => ({ kind: 'raw', height, xml: `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>` });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const num = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;

const lvl = (fmt, text, suff, rPr = '') => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>`
  + (suff ? `<w:suff w:val="${suff}"/>` : '') + `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>`
  + `<w:pPr><w:ind w:left="${STEP}" w:hanging="${STEP}"/></w:pPr>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:lvl>`;
const NUMBERING = XML_DECL + `<w:numbering ${W}>`
  + `<w:abstractNum w:abstractNumId="1">${lvl('decimal', '%1.', 'nothing')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="2">${lvl('decimal', '%1.', 'space')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="3">${lvl('decimal', '%1.', 'tab')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="□"/><w:lvlJc w:val="left"/>`
  + `<w:pPr><w:ind w:left="${DOCX.GUTTER}" w:hanging="${DOCX.GUTTER}"/></w:pPr><w:rPr>${JA}</w:rPr></w:lvl></w:abstractNum>`
  + [1, 2, 3, 4].map((n) => `<w:num w:numId="${n}"><w:abstractNumId w:val="${n}"/></w:num>`).join('')
  + '</w:numbering>';
const LIST_STYLE = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
  + `<w:pPr><w:ind w:leftChars="400" w:left="${STEP}"/></w:pPr></w:style>`;

const W1 = colStart(2) - colStart(1);
const W2 = span(2, 3);
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const ITEM = '学会・研究会等での公表（予定時期　年　月; 予定時期　年　月）';
const ROWS = [
  ['style + list + leftChars 0', '<w:pStyle w:val="ListParagraph"/>' + num(4) + '<w:ind w:leftChars="0"/>'],
  ['style + list', '<w:pStyle w:val="ListParagraph"/>' + num(4)],
  ['list only', num(4)],
  ['style only', '<w:pStyle w:val="ListParagraph"/>'],
];
const CASCADE = `<w:tbl><w:tblPr><w:tblW w:w="${W1 + W2}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W2}"/></w:tblGrid>`
  + ROWS.map(([label, pPr]) => '<w:tr>'
    + `<w:tc><w:tcPr><w:tcW w:w="${W1}" w:type="dxa"/></w:tcPr><w:p>${r(label)}</w:p></w:tc>`
    + `<w:tc><w:tcPr><w:tcW w:w="${W2}" w:type="dxa"/></w:tcPr><w:p><w:pPr>${pPr}</w:pPr>${r(ITEM, JA)}</w:p></w:tc></w:tr>`).join('')
  + '</w:tbl>';

export default {
  id: 'docx-039',
  slug: 'indent-resolution',
  format: 'docx',
  title: 'Indent resolution',
  covers: [
    'word/numbering.xml|w:lvl>w:suff',
    'word/numbering.xml|w:suff@w:val=nothing',
    'word/numbering.xml|w:suff@w:val=space',
    'word/numbering.xml|w:suff@w:val=tab',
    'word/document.xml|w:ind@w:leftChars',
    'word/document.xml|w:ind@w:hangingChars',
    'word/styles.xml|w:ind@w:leftChars',
    'word/document.xml|w:pPr>w:numPr',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-039',
      title: 'Indent resolution',
      verify: [
        'Verify: A numbers start at the margin, the body follows them',
        'and second lines indent; B ignores the 22-inch sentinel;',
        'C rows show which layer sets the indent.',
      ],
    }),
    caption('A. Hanging = left; suffix nothing, space, tab', { at: DOCX.rowLine(2), before: 0 }),
    p(rl(['Suffix nothing: the body follows the number,', 'and this second line starts at the left indent.']), num(1), 2),
    p(rl(['Suffix space: one space follows the number,', 'and this second line starts at the left indent.']), num(2), 2),
    p(rl(['Suffix tab: the tab reaches the left indent,', 'and this second line starts there too.']), num(3), 2),
    caption('B. 31680-twip sentinel with character indents'),
    p(rl(['hangingChars 400: this first line starts at the margin,', 'the following lines four characters in.'], JA),
      '<w:ind w:left="31680" w:hangingChars="400" w:firstLine="31680"/>', 2),
    p(rl(['leftChars 100 and hangingChars 1100: first line one character in,', 'the following lines twelve characters in.'], JA),
      '<w:ind w:leftChars="100" w:left="31680" w:hangingChars="1100" w:firstLine="31680"/>', 2),
    caption('C. Indent cascade in a cell: style, list level, paragraph leftChars 0'),
    { kind: 'raw', xml: CASCADE, height: NaN },
  ], {
    styles: LIST_STYLE,
    parts: { 'word/numbering.xml': NUMBERING },
    rels: [['rId3', 'numbering', 'numbering.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml` },
  }),
};
