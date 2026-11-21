import { DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header } from '../lib/docx.mjs';
import { WP14_NS, anchor, p, r, spacing } from '../lib/wml.mjs';

// Aligned and percentage-positioned anchors whose wp:effectExtent is not
// zero. Every box is an accent rectangle one column wide and 24 pt tall.
//   H (page 1), one box under each label: aligned left, right and centre
//     to the margin, inside and outside to the page, with effect extents of
//     +10 pt, −5 pt and +20 pt on the left only; and a 25 % horizontal
//     percentage offset with +10 pt.
//   V (page 2), one box per column: vertical inside and outside to the
//     page (the header and footer distances differ), top and bottom to the
//     margin with +10 pt, a 50 % vertical percentage offset with +10 pt, and
//     inside and outside to the paragraph.
const { emu, LINE, colStart, span } = DOCX;
const W = emu(span(1, 1));
const H = emu(2 * LINE);
const PT = 12700; // EMU per point
const all = (pt) => [pt * PT, pt * PT, pt * PT, pt * PT];
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const label = (text) => raw(p(spacing(LINE, 'exact'), r(text)));
const slot = (content) => raw(p(spacing(3 * LINE, 'exact'), content));
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

let nextId = 1;
const box = (h, v, effect = [0, 0, 0, 0]) => anchor({ id: nextId++, h, v, cx: W, cy: H, effect });
const H_CASES = [
  ['H1 margin left, no effect', { from: 'margin', align: 'left' }, [0, 0, 0, 0]],
  ['H2 margin left, effect +10 pt', { from: 'margin', align: 'left' }, all(10)],
  ['H3 margin right, effect +10 pt', { from: 'margin', align: 'right' }, all(10)],
  ['H4 margin centre, effect +20 pt on the left', { from: 'margin', align: 'center' }, [20 * PT, 0, 0, 0]],
  ['H5 page inside, effect +10 pt', { from: 'page', align: 'inside' }, all(10)],
  ['H6 page outside, effect +10 pt', { from: 'page', align: 'outside' }, all(10)],
  ['H7 margin left, effect −5 pt', { from: 'margin', align: 'left' }, all(-5)],
  ['H8 margin 25 %, effect +10 pt', { from: 'margin', pct: 25000, fallback: emu(colStart(2)) }, all(10)],
];
const col = (c) => ({ from: 'margin', offset: emu(colStart(c)) });

const parts = () => {
  nextId = 1;
  return docxPackage([
    ...header({
      id: 'docx-055',
      title: 'Anchors with effect extents',
      verify: [
        'Verify: each box sits where Word puts it: the box plus',
        'its effect extent is aligned; a percentage offset adds',
        'the left or top effect. Compare the PDF.',
      ],
    }),
    caption('H. Horizontal alignment (box under each label)', { at: DOCX.rowLine(2), before: 0 }),
    ...H_CASES.flatMap(([name, h, effect]) => [label(name), slot(box(h, { from: 'paragraph', offset: 0 }, effect))]),
    { kind: 'p', style: 'caption', bold: true, lines: ['V. Vertical alignment, one box per column'], height: NaN, pageBreakBefore: true },
    label('Column 1: V1 page inside and V2 page outside (no effect).'),
    label('Column 2: V3 margin top and V4 margin bottom, effect +10 pt.'),
    label('Column 3: V5 margin 50 %, effect +10 pt.'),
    label('Column 4: V6 inside and V7 outside the paragraph below.'),
    raw(p(spacing(LINE, 'exact'), box(col(1), { from: 'page', align: 'inside' })
      + box(col(1), { from: 'page', align: 'outside' })
      + box(col(2), { from: 'margin', align: 'top' }, all(10))
      + box(col(2), { from: 'margin', align: 'bottom' }, all(10))
      + box(col(3), { from: 'margin', pct: 50000, fallback: emu(30 * LINE) }, all(10))
      + r('This paragraph holds V1 to V5.'))),
    raw(p(spacing(8 * LINE, 'exact'), box(col(4), { from: 'paragraph', align: 'inside' })
      + box(col(4), { from: 'paragraph', align: 'outside' })
      + r('V6 and V7 are anchored in this 96 pt paragraph.'))),
  ], {
    firstSection: { headerDist: DOCX.HEADER_DIST, footerDist: DOCX.FOOTER_DIST },
    ns: [R_NS, WP_NS, A_NS, WPS_NS, MC_NS, WP14_NS],
  });
};

export default {
  id: 'docx-055',
  slug: 'anchor-alignment-effects',
  format: 'docx',
  title: 'Anchors with effect extents',
  covers: [
    'word/document.xml|wp:anchor>wp:effectExtent',
    'word/document.xml|wp:positionH>wp:align',
    'word/document.xml|wp:positionV>wp:align',
    'word/document.xml|wp:positionH>wp14:pctPosHOffset',
    'word/document.xml|wp:positionV>wp14:pctPosVOffset',
  ],
  parts,
};
