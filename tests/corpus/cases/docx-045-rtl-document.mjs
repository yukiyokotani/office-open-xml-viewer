import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { WML, W, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';

// Right-to-left text in Word (w:bidi paragraphs, w:rtl runs), in Arabic.
//   A. A paragraph with an English word inside the sentence, Western and
//      Arabic-Indic digits, and an ISO date, which Word keeps in its LTR
//      order inside the RTL line.
//   B. Bold and size on the complex-script axis: w:b alone (bold), w:b with
//      w:bCs="0" (not bold), w:sz alone (the size applies), and w:szCs.
//   C. A bulleted and a numbered list in RTL paragraphs: markers on the
//      right, hanging indents mirrored.
//   D. Contents lines: a right-aligned tab with a dot leader to the page
//      number, in RTL paragraphs that also say w:jc="left".
//   E. A table with w:bidiVisual (the first column on the right), a cell
//      spanning two columns, and a date column.
const { LINE, colStart, span } = DOCX;
const AR = `<w:rFonts w:cs="${FONT.arabic}"/>`;
const RTL = `${AR}<w:rtl/>`;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const ar = (text, extra = '') => `<w:r><w:rPr>${RTL}${extra}</w:rPr>${t(text)}</w:r>`;
const en = (text) => `<w:r>${t(text)}</w:r>`;
const BIDI = '<w:bidi/>';
const p = (runs, pPr = '', height = 1) => ({ kind: 'raw', height, xml: `<w:p><w:pPr>${BIDI}${pPr}</w:pPr>${runs}</w:p>` });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const num = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;

const STEP = colStart(2) - colStart(1);
const NUMBERING = XML_DECL + `<w:numbering ${W}>`
  + `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${STEP}" w:hanging="${DOCX.GUTTER}"/></w:pPr></w:lvl></w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${STEP}" w:hanging="${DOCX.GUTTER}"/></w:pPr></w:lvl></w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>';

const TOC_TAB = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${span(1, 2)}"/></w:tabs>`;
const toc = (title, page) => p(ar(title) + `<w:r><w:rPr>${RTL}</w:rPr><w:tab/></w:r>` + ar(String(page)),
  `${TOC_TAB}<w:ind w:right="0" w:left="${DOCX.LIVE_W - span(1, 2)}"/><w:jc w:val="left"/>`);

const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('') + '</w:tblBorders>';
const GRID = [colStart(2) - colStart(1), colStart(3) - colStart(2), span(3, 4)];
const tc = (w, runs, extra = '') => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:pPr>${BIDI}</w:pPr>${runs}</w:p></w:tc>`;
const tr = (cells) => `<w:tr><w:trPr><w:trHeight w:val="${LINE}" w:hRule="exact"/></w:trPr>${cells}</w:tr>`;
const TABLE = `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="${GRID.reduce((a, b) => a + b, 0)}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid>${GRID.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  + tr(tc(GRID[0] + GRID[1], ar('المقاس والأبعاد (خلية تمتد على عمودين)', '<w:b/>'), '<w:gridSpan w:val="2"/>') + tc(GRID[2], ar('تاريخ الإصدار', '<w:b/>')))
  + tr(tc(GRID[0], ar('A4')) + tc(GRID[1], ar('210 × 297')) + tc(GRID[2], ar('2026-02-28')))
  + tr(tc(GRID[0], ar('B5')) + tc(GRID[1], ar('176 × 250')) + tc(GRID[2], ar('2026-11-19')))
  + '</w:tbl>';

export default {
  id: 'docx-045',
  slug: 'rtl-document',
  format: 'docx',
  title: 'Right-to-left document',
  covers: [
    'word/document.xml|w:pPr>w:bidi',
    'word/document.xml|w:rPr>w:rtl',
    'word/document.xml|w:rPr>w:bCs',
    'word/document.xml|w:rPr>w:szCs',
    'word/document.xml|w:rFonts@w:cs',
    'word/document.xml|w:tblPr>w:bidiVisual',
    'word/document.xml|w:tab@w:leader=dot',
    'word/numbering.xml|w:numFmt@w:val=bullet',
    'word/numbering.xml|w:numFmt@w:val=decimal',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-045',
      title: 'Right-to-left document',
      verify: [
        'Verify: RTL lines start at the right; the English word',
        'and numbers keep their order; list markers and the',
        'first column are on the right; leaders reach numbers.',
      ],
    }),
    caption('A. Mixed text, digits and a date in an RTL paragraph', { at: DOCX.rowLine(2), before: 0 }),
    p(ar('هذه الفقرة مكتوبة من اليمين إلى اليسار، وفيها كلمة ') + en('English') + ar(' وسط الجملة.')),
    p(ar('مقاس A4 هو 210 × 297 مم، وبالأرقام الهندية ٢١٠ × ٢٩٧ مم.')),
    p(ar('تاريخ الإصدار 2026-02-28 يبقى بترتيبه.')),
    caption('B. b, b with bCs 0, sz alone, szCs'),
    p(ar('عريض بالخاصية b وحدها', '<w:b/>') + ar('، ') + ar('b مع bCs=0', '<w:b/><w:bCs w:val="0"/>')),
    p(ar('حجم sz وحده 12 نقطة', '<w:sz w:val="24"/>') + ar('، ') + ar('حجم szCs 12 نقطة', '<w:szCs w:val="24"/>'), '', 1),
    caption('C. Bulleted and numbered RTL lists'),
    p(ar('المقاس A3'), num(1)),
    p(ar('المقاس A4'), num(1)),
    p(ar('الخطوة الأولى'), num(2)),
    p(ar('الخطوة الثانية'), num(2)),
    caption('D. Contents with dot leaders, RTL paragraphs with jc left'),
    toc('المقدمة', 1),
    toc('مقاسات السلسلة A', 2),
    caption('E. bidiVisual table with a spanning cell'),
    { kind: 'raw', xml: TABLE, height: 3 },
  ], {
    parts: { 'word/numbering.xml': NUMBERING },
    rels: [['rId3', 'numbering', 'numbering.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml` },
  }),
};
