import { DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, MC_NS, PIC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header } from '../lib/docx.mjs';
import { ORIENT_PNG, anchor, inlinePic, p, r, spacing } from '../lib/wml.mjs';

// Vertical sections: what sits on the physical page.
//   Page 2, tbRl: underlines (single, double, thick, wavy) on Japanese and
//     Latin runs, a raised underlined run, an underlined tab and a struck
//     run; a picture raised 5 pt on an exact 24 pt line; two rectangles
//     anchored to the page and to the margin at fixed offsets.
//   Page 3, tbLrV, margins 50 mm left and 10 mm right: the same lines and
//     rectangles. The text starts at the physical left margin.
// Page 1 is the usual horizontal header.
const { emu, LINE, MARGIN } = DOCX;
const PT = 20;
const JA = `<w:rFonts w:ascii="${FONT.latin}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.latin}"/>`;
const ASYM = { top: MARGIN.top, bottom: MARGIN.bottom, left: 2835, right: 567, header: 0, footer: 0 };
const BOX = emu(2 * LINE);
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const line = (content) => raw(p(spacing(LINE, 'auto'), content));

let nextId = 1;
const body = (name) => [
  line(`${r(`${name}: `)}${r('一重下線', `${JA}<w:u w:val="single"/>`)}${r(' ')}${r('Latin single', '<w:u w:val="single"/>')}`),
  line(`${r('二重下線', `${JA}<w:u w:val="double"/>`)}${r(' ')}${r('太い下線', `${JA}<w:u w:val="thick"/>`)}${r(' ')}${r('波線', `${JA}<w:u w:val="wave"/>`)}`),
  line(`${r('基準')}${r('上付き下線', `${JA}<w:position w:val="6"/><w:u w:val="single"/>`)}${r('基準')}`),
  line(`${r('タブ')}<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:tab/></w:r>${r('の下線と')}${r('取り消し線', `${JA}<w:strike/>`)}`),
  raw(p(spacing(24 * PT, 'exact'), `${r('行 ')}${inlinePic({ id: nextId++, rId: 'rId5', cx: emu(24 * PT), cy: emu(24 * PT), position: 10 })}${r(' 5 pt 上げた画像')}`)),
  line(`${anchor({ id: nextId++, h: { from: 'page', offset: emu(1134) }, v: { from: 'page', offset: emu(1134) }, cx: BOX, cy: BOX })}`
    + `${anchor({ id: nextId++, h: { from: 'margin', offset: 0 }, v: { from: 'margin', offset: 0 }, cx: BOX, cy: BOX, graphic: undefined })}`
    + r('四角はページ (20 mm, 20 mm) と余白の原点')),
];

const parts = () => {
  nextId = 1;
  return docxPackage([
    ...header({
      id: 'docx-053',
      title: 'Vertical sections on the page',
      verify: [
        'Verify: p.2 tbRl, p.3 tbLrV; underlines on the right of',
        'each column; squares at page 20 mm and margin origin;',
        'p.3 text starts at the 50 mm left margin.',
      ],
    }),
    { kind: 'p', style: 'caption', lines: ['Pages 2 and 3 are vertical sections.'], at: DOCX.rowLine(2) },
    { kind: 'section', type: 'nextPage', textDirection: 'tbRl' },
    ...body('tbRl'),
    { kind: 'p', lines: ['tbRl ページの終わり。'] },
    { kind: 'section', type: 'nextPage', textDirection: 'tbLrV', pgMar: ASYM },
    ...body('tbLrV'),
  ], {
    parts: { 'word/media/image1.png': ORIENT_PNG },
    rels: [['rId5', 'image', 'media/image1.png']],
    defaults: { png: 'image/png' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-053',
  slug: 'vertical-sections',
  format: 'docx',
  title: 'Vertical sections on the page',
  offGrid: true,
  covers: [
    'word/document.xml|w:textDirection@w:val=tbRl',
    'word/document.xml|w:textDirection@w:val=tbLrV',
    'word/document.xml|w:u@w:val=double',
    'word/document.xml|w:u@w:val=thick',
    'word/document.xml|w:u@w:val=wave',
    'word/document.xml|w:rPr>w:strike',
    'word/document.xml|wp:positionH@relativeFrom=page',
  ],
  parts,
};
