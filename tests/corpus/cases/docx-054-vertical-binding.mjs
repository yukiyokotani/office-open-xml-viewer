import { DOCX, FONT, LANG } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';
import { ORIENT_PNG, inlinePic, p, r, spacing } from '../lib/wml.mjs';

// A tbLrV section bound like a book: mirrored margins (w:mirrorMargins) and
// a 10 mm gutter, over two facing pages (2 and 3). Each page mixes Japanese
// and Latin text, ruby, an underline and an upright 60 × 40 pt inline
// picture. Inside each line the text runs as in tbRl; only the order of the
// lines is reversed, and the gutter is on the binding side tbRl uses.
// Page 1 is the usual horizontal header.
const { emu, LINE, MARGIN } = DOCX;
const PT = 20;
const BOOK = { top: MARGIN.top, bottom: MARGIN.bottom, left: MARGIN.left, right: MARGIN.right, header: 0, footer: 0, gutter: 567 };
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const ruby = (base, text) => '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="9"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="18"/>'
  + `<w:lid w:val="${LANG.eastAsia}"/></w:rubyPr><w:rt><w:r><w:rPr>${JA}<w:sz w:val="9"/></w:rPr><w:t>${esc(text)}</w:t></w:r></w:rt>`
  + `<w:rubyBase><w:r><w:rPr>${JA}</w:rPr><w:t>${esc(base)}</w:t></w:r></w:rubyBase></w:ruby></w:r>`;
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const line = (content) => raw(p(spacing(2 * LINE, 'exact'), content));

let nextId = 1;
const page = (n, pageBreak) => [
  raw(p(spacing(2 * LINE, 'exact'), `${pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''}${r(`${n} ページ目、行は左から右へ並ぶ。`)}`)),
  line(`${r('縦書きに Latin text と数字 2026 を混ぜる。')}`),
  line(`${ruby('紫陽花', 'あじさい')}${r('と')}${ruby('五月雨', 'さみだれ')}${r('にルビを付ける。')}`),
  line(`${r('下線を引いた語', '<w:u w:val="single"/>')}${r('と、続く本文。')}`),
  line(`${r('画像 ')}${inlinePic({ id: nextId++, rId: 'rId5', cx: emu(60 * PT), cy: emu(40 * PT) })}${r(' は立ったまま 60 × 40 pt。')}`),
  line(r('とじしろ 10 mm は綴じる側にある。')),
];

const parts = () => {
  nextId = 1;
  return docxPackage([
    ...header({
      id: 'docx-054',
      title: 'Vertical book binding',
      verify: [
        'Verify: p.2-3 tbLrV, lines left to right, glyphs as in',
        'tbRl; the 10 mm gutter is on the binding side; the',
        'picture is upright, 60 × 40 pt. Compare the PDF.',
      ],
    }),
    { kind: 'p', style: 'caption', lines: ['Pages 2 and 3 face each other (mirrored margins).'], at: DOCX.rowLine(2) },
    { kind: 'section', type: 'nextPage', textDirection: 'tbLrV', pgMar: BOOK },
    ...page(2, false),
    ...page(3, true),
  ], {
    settings: '<w:mirrorMargins/>',
    parts: { 'word/media/image1.png': ORIENT_PNG },
    rels: [['rId5', 'image', 'media/image1.png']],
    defaults: { png: 'image/png' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
  });
};

export default {
  id: 'docx-054',
  slug: 'vertical-binding',
  format: 'docx',
  title: 'Vertical book binding',
  offGrid: true,
  covers: [
    'word/document.xml|w:textDirection@w:val=tbLrV',
    'word/document.xml|w:pgMar@w:gutter',
    'word/settings.xml|w:settings>w:mirrorMargins',
    'word/document.xml|w:r>w:ruby',
    'word/document.xml|wp:inline>wp:extent',
  ],
  parts,
};
