import { DOCX } from '../lib/grid.mjs';
import { WML, W, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';

// Office font regression set. The text names fonts that ship with Microsoft
// Office or Windows and are not embedded, so outside Office they are drawn
// with substitutes whose widths and line heights differ: line breaks and the
// page count will not match Office. What must hold is the layout: no text
// outside its cell or column, no overlapping lines, nothing cut at the page
// foot. Page 1 sets one sample per font in a table whose rows grow with the
// text; page 2 is a packed Japanese paper page on a tight document grid
// (linesAndChars, 14.3 pt pitch, narrowed character pitch) with a one-column
// title block, an English subtitle that nearly fills the line, and a
// two-column body that runs to the foot of the page.
const LATIN = [
  ['Calibri', 'swiss'], ['Cambria', 'roman'], ['Aptos', 'swiss'], ['Times New Roman', 'roman'], ['Arial', 'swiss'], ['Century', 'roman'],
];
const JAPANESE = [
  ['游明朝', 'Yu Mincho', 'roman'], ['游ゴシック', 'Yu Gothic', 'swiss'], ['ＭＳ 明朝', 'MS Mincho', 'roman'],
  ['ＭＳ ゴシック', 'MS Gothic', 'modern'], ['ＭＳ Ｐゴシック', 'MS PGothic', 'swiss'], ['メイリオ', 'Meiryo', 'swiss'],
];
const FONT_TABLE = XML_DECL + `<w:fonts ${W}>`
  + LATIN.map(([name, family]) => `<w:font w:name="${name}"><w:charset w:val="00"/><w:family w:val="${family}"/><w:pitch w:val="variable"/></w:font>`).join('')
  + JAPANESE.map(([name, alt, family]) => `<w:font w:name="${name}"><w:altName w:val="${alt}"/><w:charset w:val="80"/><w:family w:val="${family}"/>`
    + `<w:pitch w:val="${family === 'modern' ? 'fixed' : 'variable'}"/></w:font>`).join('')
  + '</w:fonts>';

const fonts = (latin, ea = latin) => `<w:rFonts w:ascii="${latin}" w:hAnsi="${latin}" w:eastAsia="${ea}" w:cs="${latin}"/>`;
const r = (text, rPr) => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;

const EN = 'The A series halves from A0, one square metre, down to A10; every size keeps the ratio 1 : 1.414, '
  + 'so a sheet folded across its long side keeps its shape.';
const JA = 'Ａ列はＡ０（面積１平方メートル）から始まり、長辺を半分に折るたびに次の判になる。'
  + '縦横の比は１：１．４１４で、折っても形は変わらない。';

// Page 1: font name in column 1, sample in columns 2-4, 9 pt on single
// (auto) line spacing so each row is as tall as the substitute makes it.
const NAME_W = DOCX.span(1, 1) + DOCX.gutterAfter(1);
const SAMPLE_W = DOCX.span(2, 4);
const CELL_P = '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>';
const cell = (w, content) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p>${CELL_P}${content}</w:p></w:tc>`;
const SZ = '<w:sz w:val="18"/>';
const row = (name, text, font, ea) => `<w:tr>${cell(NAME_W, r(name, `${fonts('Noto Sans', 'Noto Sans JP')}${SZ}`))}`
  + `${cell(SAMPLE_W, r(text, `${fonts(font, ea)}${SZ}`))}</w:tr>`;
const RULES = '<w:tblBorders>' + ['top', 'bottom', 'insideH'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const TABLE = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${DOCX.LIVE_W}" w:type="dxa"/>${RULES}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${NAME_W}"/><w:gridCol w:w="${SAMPLE_W}"/></w:tblGrid>`
  + LATIN.map(([name]) => row(name, EN, name)).join('')
  + JAPANESE.map(([name]) => row(name, JA, 'Times New Roman', name)).join('')
  + '</w:tbl>';

// Page 2: a packed paper page.
const MINCHO = fonts('Century', 'ＭＳ 明朝');
const GOTHIC = fonts('Arial', 'ＭＳ ゴシック');
const BODY_JA = [
  '紙の寸法の国際規格ISO 216は、Ａ列とＢ列の二つの系列を定める。Ａ列の基準となるＡ０は面積が１平方メートルで、縦横の比が１：√２の長方形である。',
  'この比を選んだ理由は、長辺の中点で二つに折ったとき、半分の紙が元と同じ形になることにある。Ａ０を半分にするとＡ１、さらに半分にするとＡ２となり、Ａ４は210×297ミリメートルである。',
  'Ｂ列は隣り合うＡ列の寸法の幾何平均で、Ｂ０は1000×1414ミリメートルである。封筒に用いるＣ列はＡ列とＢ列の間にあり、Ｃ４の封筒はＡ４の紙を折らずに収める。',
  '日本の工業規格JIS P 0138は、Ａ列を国際規格と同じく定め、Ｂ列は独自の寸法とした。JISのＢ列はＡ列と面積の比が１．５倍になるように決められ、Ｂ５は182×257ミリメートルである。',
  '寸法はミリメートル単位で切り捨てて表す。そのため、比を厳密に計算すると末尾がわずかに合わないことがあるが、実用上は差し支えない。',
];
// Single (auto) spacing so lines snap to the document grid, not the 12 pt exact default.
const para = (text, rPr, extra = '') => `<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>${extra}</w:pPr>${r(text, rPr)}</w:p>`;
const PAGE2_GRID = '<w:docGrid w:type="linesAndChars" w:linePitch="286" w:charSpace="-3426"/>';
const TITLE_BLOCK = [
  para('紙の寸法の規格と、その比が保たれる理由', `${GOTHIC}<w:b/><w:sz w:val="28"/>`, '<w:jc w:val="center"/>'),
  para('Paper sizes in the ISO 216 and JIS P 0138 series and why folding a sheet keeps its proportions', `${fonts('Times New Roman')}<w:sz w:val="24"/>`, '<w:jc w:val="center"/>'),
  para('概要：Ａ列・Ｂ列・Ｃ列の寸法と縦横比を整理し、折っても形が変わらない理由を説明する。', `${MINCHO}${SZ}`, '<w:ind w:leftChars="200" w:rightChars="200"/>'),
].join('');
const BODY = [
  para('１　はじめに', `${GOTHIC}${SZ}<w:b/>`),
  ...BODY_JA.map((t) => para(t, `${MINCHO}${SZ}`, '<w:ind w:firstLineChars="100"/><w:jc w:val="both"/>')),
  para('２　寸法の一覧', `${GOTHIC}${SZ}<w:b/>`),
  ...[...BODY_JA, ...BODY_JA, ...BODY_JA, ...BODY_JA].map((t) => para(t, `${MINCHO}${SZ}`, '<w:ind w:firstLineChars="100"/><w:jc w:val="both"/>')),
  para('３　まとめ', `${GOTHIC}${SZ}<w:b/>`),
  ...BODY_JA.slice(0, 3).map((t) => para(t, `${MINCHO}${SZ}`, '<w:ind w:firstLineChars="100"/><w:jc w:val="both"/>')),
].join('');

export default {
  id: 'docx-034',
  slug: 'office-fonts',
  format: 'docx',
  title: 'Office fonts',
  covers: [
    'word/fontTable.xml|w:fonts>w:font',
    'word/fontTable.xml|w:font>w:altName',
    'word/fontTable.xml|w:font>w:family',
    'word/fontTable.xml|w:family@w:val=roman',
    'word/fontTable.xml|w:family@w:val=swiss',
    'word/fontTable.xml|w:family@w:val=modern',
    'word/document.xml|w:rPr>w:rFonts',
    'word/document.xml|w:ind@w:firstLineChars',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-034',
      title: 'Office fonts',
      verify: [
        'Verify: Office fonts not embedded; substitutes may move',
        'line breaks and pages, but no text leaves its cell or',
        'column, no lines overlap and nothing is cut at a foot.',
      ],
    }),
    { kind: 'raw', at: DOCX.rowLine(2), xml: TABLE },
    { kind: 'p', lines: [''] },
    // Page 2: the title block in one column on the tight grid, then the body
    // in two columns on the same grid (the body-level sectPr).
    { kind: 'section', type: 'nextPage', cols: { raw: `<w:cols w:space="${DOCX.gutterAfter(2)}"/>` }, docGrid: PAGE2_GRID },
    { kind: 'raw', xml: TITLE_BLOCK },
    { kind: 'p', lines: [''] },
    { kind: 'section', type: 'continuous', cols: { num: 2, space: DOCX.gutterAfter(2) }, docGrid: PAGE2_GRID },
    { kind: 'raw', xml: BODY },
  ], {
    settings: '<w:characterSpacingControl w:val="compressPunctuation"/>',
    parts: { 'word/fontTable.xml': FONT_TABLE },
    rels: [['rId3', 'fontTable', 'fontTable.xml']],
    overrides: { 'word/fontTable.xml': `${WML}.fontTable+xml` },
  }),
};
