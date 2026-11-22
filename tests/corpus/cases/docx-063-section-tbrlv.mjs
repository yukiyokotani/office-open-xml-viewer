import { DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, W, WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';
import { ORIENT_PNG, inlinePic, p, r, spacing } from '../lib/wml.mjs';

// A section whose w:textDirection is tbRlV, next to the same content in a
// tbLrV section. Pages 2-3 are tbRlV and pages 4-5 tbLrV; each page holds
// the same paragraphs: underlined and emphasis-marked Japanese and Latin
// runs, a 2 × 2 table and an upright 48 × 24 pt inline picture, in two
// columns (a column break in the middle). Both sections have a 10 mm gutter,
// a header and a footer; the document mirrors its margins, so pages 2 and 4
// match, and pages 3 and 5. Page 1 is the usual horizontal header.
const { emu, LINE, MARGIN, span } = DOCX;
const PT = 20;
const BOOK = { top: MARGIN.top, bottom: MARGIN.bottom, left: MARGIN.left, right: MARGIN.right, header: 567, footer: 567, gutter: 567 };
const COLS = { num: 2, space: 567 };
const JA = `<w:rFonts w:ascii="${FONT.latin}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.latin}"/>`;
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const line = (content, pPr = '') => raw(p(`${pPr}${spacing(2 * LINE, 'exact')}`, content));
const REFS = '<w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>';
const furniture = (tag, text) => XML_DECL + `<w:${tag} ${W}>${p('<w:pStyle w:val="Caption"/>', r(text))}</w:${tag}>`;

const TBL_W = 2 * span(1, 1);
const cell = (text) => `<w:tc><w:tcPr><w:tcW w:w="${TBL_W / 2}" w:type="dxa"/></w:tcPr>${p('', r(text))}</w:tc>`;
const table = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${TBL_W}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')}</w:tblBorders>`
  + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${TBL_W / 2}"/><w:gridCol w:w="${TBL_W / 2}"/></w:tblGrid>`
  + `<w:tr>${cell('表 1')}${cell('表 2')}</w:tr><w:tr>${cell('表 3')}${cell('表 4')}</w:tr></w:tbl>`;

let nextId = 1;
const page = (name, n, pageBreak) => [
  line(`${pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''}${r(`${name}、${n} ページ目の一行目。`)}`),
  line(`${r('下線', `${JA}<w:u w:val="single"/>`)}${r('と')}${r('傍点', `${JA}<w:em w:val="dot"/>`)}${r('と Latin text 2026。')}`),
  raw(table),
  line(`${r('画像 ')}${inlinePic({ id: nextId++, rId: 'rId5', cx: emu(48 * PT), cy: emu(24 * PT) })}${r(' は立ったまま。')}`),
  line(`<w:r><w:br w:type="column"/></w:r>${r('二段目の一行目。')}`),
  line(r('二段目の二行目、とじしろは綴じる側。')),
];

const parts = () => {
  nextId = 1;
  const section = (textDirection) => ({ kind: 'section', type: 'nextPage', textDirection, pgMar: BOOK, cols: COLS, refs: REFS });
  return docxPackage([
    ...header({
      id: 'docx-063',
      title: 'Section text tbRlV',
      verify: [
        'Verify: pages 2-3 (tbRlV) match pages 4-5 (tbLrV):',
        'lines run left to right from the left margin, glyphs',
        'as in tbRl. Compare the PDF.',
      ],
    }),
    { kind: 'p', style: 'caption', lines: ['Pages 2-3 tbRlV, pages 4-5 tbLrV.'], at: DOCX.rowLine(2) },
    section('tbRlV'),
    ...page('tbRlV', 2, false),
    ...page('tbRlV', 3, true),
    { kind: 'p', lines: ['tbRlV の終わり。'] },
    section('tbLrV'),
    ...page('tbLrV', 4, false),
    ...page('tbLrV', 5, true),
  ], {
    settings: '<w:mirrorMargins/>',
    parts: {
      'word/media/image1.png': ORIENT_PNG,
      'word/header1.xml': furniture('hdr', 'Header of the vertical sections'),
      'word/footer1.xml': furniture('ftr', 'Footer of the vertical sections'),
    },
    rels: [['rId3', 'header', 'header1.xml'], ['rId4', 'footer', 'footer1.xml'], ['rId5', 'image', 'media/image1.png']],
    overrides: { 'word/header1.xml': `${WML}.header+xml`, 'word/footer1.xml': `${WML}.footer+xml` },
    defaults: { png: 'image/png' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
  });
};

export default {
  id: 'docx-063',
  slug: 'section-tbrlv',
  format: 'docx',
  title: 'Section text tbRlV',
  offGrid: true,
  covers: [
    'word/document.xml|w:textDirection@w:val=tbRlV',
    'word/document.xml|w:textDirection@w:val=tbLrV',
    'word/document.xml|w:pgMar@w:gutter',
    'word/settings.xml|w:settings>w:mirrorMargins',
    'word/document.xml|w:rPr>w:em',
    'word/document.xml|w:sectPr>w:cols',
  ],
  parts,
};
