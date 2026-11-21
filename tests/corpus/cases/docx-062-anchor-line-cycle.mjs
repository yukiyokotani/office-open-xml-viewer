import { DOCX } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage } from '../lib/docx.mjs';
import { ORIENT_PNG, anchor, p, picGraphic, r, spacing } from '../lib/wml.mjs';

// Two pictures wrapped top and bottom whose anchor paragraphs push each
// other: whether the second picture's paragraph fits on page 1 decides
// whether the first one's does. The geometry is the subject: US Letter,
// 1 in margins, 30 one-line paragraphs of exactly 24 pt (27 to a page).
//   Part 1 (from page 1): paragraph 003 anchors a 200 × 300 pt picture,
//     paragraph 010 a 200 × 600 pt one, both at the margin origin.
//   Part 2 (a new page): the same, plus paragraph 011 anchoring a 100 pt
//     tall picture.
// Paragraph 000 names the case; Word's own layout is the reference.
const PT = 12700;
const LETTER = { w: 12240, h: 15840 };
const MARGINS = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 };
const LINE24 = spacing(480, 'exact');

let nextId = 1;
const pic = (hPt) => anchor({
  id: nextId, h: { from: 'margin', offset: 0 }, v: { from: 'margin', offset: 0 }, cx: 200 * PT, cy: hPt * PT,
  wrap: '<wp:wrapTopAndBottom/>', graphic: picGraphic({ id: nextId++, rId: 'rId5', cx: 200 * PT, cy: hPt * PT }),
});
const part = (prefix, pictures, first) => Array.from({ length: 30 }, (_, i) => {
  const name = `${prefix}${String(i).padStart(3, '0')}`;
  const text = i === 0 ? first : `${name}${pictures[i] ? ` anchors a ${pictures[i]} pt picture` : ''}`;
  return { kind: 'raw', height: NaN, xml: p(LINE24, `${pictures[i] ? pic(pictures[i]) : ''}${r(text)}`) };
});

const parts = () => {
  nextId = 1;
  const one = part('P', { 3: 300, 10: 600 }, 'P000 docx-062 anchor-line cycle. Verify: as in the PDF.');
  const two = part('Q', { 3: 300, 10: 600, 11: 100 }, 'Q000 part 2: as part 1, plus a 100 pt picture in Q011.');
  return docxPackage([
    ...one.slice(0, -1),
    { kind: 'p', runs: [r('P029')], height: NaN, pPrMid: LINE24 },
    { kind: 'section', type: 'nextPage', pgSz: LETTER, pgMar: MARGINS },
    ...two,
  ], {
    firstSection: { pgSz: LETTER, pgMar: MARGINS },
    parts: { 'word/media/image1.png': ORIENT_PNG },
    rels: [['rId5', 'image', 'media/image1.png']],
    defaults: { png: 'image/png' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
  });
};

export default {
  id: 'docx-062',
  slug: 'anchor-line-cycle',
  format: 'docx',
  title: 'Anchors that push each other',
  offGrid: true,
  covers: [
    'word/document.xml|wp:anchor>wp:wrapTopAndBottom',
    'word/document.xml|wp:positionV@relativeFrom=margin',
    'word/document.xml|w:spacing@w:lineRule=exact',
  ],
  parts,
};
