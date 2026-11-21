import { COLOR, DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS, wpsShape } from '../lib/drawing.mjs';
import { WML, W, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';

// Lists, a contents block, a drop cap and two-sided wrapping:
//   A. Contents lines: a right tab with a dot leader to italic page numbers.
//   B. A lowercase-roman list whose numbers are right-aligned (w:lvlJc
//      right), so i to viii line up on their right edge; it is interrupted
//      by a plain paragraph and continues at v.
//   C. A drop cap three lines high (w:framePr w:dropCap="drop"), the text
//      wrapping beside it.
//   D. One paragraph between a square on the left and a square on the
//      right, both with square wrapping at the same height: every line
//      runs between them.
const { colStart, span, emu, LINE } = DOCX;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const TAB = '<w:r><w:tab/></w:r>';
const p = (runs, pPr = '', height = 1) => ({ kind: 'raw', height, xml: `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>` });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

const STEP = colStart(2) - colStart(1);
const NUMBERING = XML_DECL + `<w:numbering ${W}>`
  + '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/>'
  + `<w:lvlJc w:val="right"/><w:pPr><w:ind w:left="${STEP}" w:hanging="${DOCX.GUTTER}"/></w:pPr></w:lvl></w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
const num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
const TOC_TABS = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${span(1, 3)}"/></w:tabs>`;
const toc = (title, page) => p(r(title) + TAB + r(String(page), '<w:i/>'), TOC_TABS);

const ROMAN = ['First item', 'Second item', 'Third item', 'Fourth item', 'Fifth item, after the break', 'Sixth item', 'Seventh item', 'Eighth item, viii'];
const DROP = `<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>`;
const DROP_TEXT = ['he A series begins with A0, a sheet of one square metre whose sides are in the', 'ratio of one to the square root of two. Folding it in half across the long side gives A1,', 'and every further fold gives the next size with the same ratio. The drop cap', 'occupies the first three lines; this fourth line returns to the margin.'];

const parts = () => {
  let id = 1;
  const sq = (x) => wpsShape({ id: id++, x: emu(x), y: 0, cx: emu(span(1, 1)), cy: emu(4 * LINE), fill: COLOR.accent, wrap: '<wp:wrapSquare wrapText="bothSides"/>' });
  const WRAPPED = 'This paragraph is set between two squares wrapped on both sides. Every line starts after the left square and ends before the right one, so the text runs in the two middle columns until the squares end, and then returns to the full width of the live area below them.';
  return docxPackage([
    ...header({
      id: 'docx-046',
      title: 'Lists, contents and drop cap',
      verify: [
        'Verify: A leaders reach italic page numbers; B numerals',
        'align right and continue at v; C drop cap spans three',
        'lines; D text runs between the two squares.',
      ],
    }),
    caption('A. Contents: dot leaders, italic page numbers', { at: DOCX.rowLine(2), before: 0 }),
    toc('1  Paper sizes', 1),
    toc('2  The A series', 2),
    toc('3  Envelopes of the C series', 14),
    caption('B. Lower roman, right-aligned numbers, continued after a break'),
    ...ROMAN.slice(0, 4).map((x) => p(r(x), num)),
    p(r('An unnumbered paragraph interrupts the list.')),
    ...ROMAN.slice(4).map((x) => p(r(x), num)),
    caption('C. Drop cap, three lines'),
    { kind: 'raw', height: NaN, xml: `<w:p><w:pPr>${DROP}<w:spacing w:line="${3 * LINE}" w:lineRule="exact"/></w:pPr>${r('T', '<w:position w:val="-14"/><w:sz w:val="72"/>')}</w:p>` },
    { kind: 'raw', height: NaN, xml: `<w:p><w:r>${DROP_TEXT.map((x, i) => `${i ? '<w:br/>' : ''}${t(x)}`).join('')}</w:r></w:p>` },
    caption('D. Text between two squares wrapped on both sides'),
    { kind: 'raw', height: NaN, xml: `<w:p>${sq(0)}${sq(colStart(4))}${r(WRAPPED)}</w:p>` },
  ], {
    parts: { 'word/numbering.xml': NUMBERING },
    rels: [['rId3', 'numbering', 'numbering.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml` },
    ns: [A_NS, WP_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-046',
  slug: 'lists-contents-dropcap',
  format: 'docx',
  title: 'Lists, contents and drop cap',
  covers: [
    'word/document.xml|w:tab@w:leader=dot',
    'word/document.xml|w:rPr>w:i',
    'word/numbering.xml|w:numFmt@w:val=lowerRoman',
    'word/numbering.xml|w:lvlJc@w:val=right',
    'word/document.xml|w:pPr>w:framePr',
    'word/document.xml|w:framePr@w:dropCap=drop',
    'word/document.xml|wp:anchor>wp:wrapSquare',
  ],
  parts,
};
