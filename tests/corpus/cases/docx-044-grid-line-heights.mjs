import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS, wpsShape } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// How many line-grid pitches a single-spaced line takes. In section A
// (w:docGrid type="lines", linePitch 18 pt) each paragraph holds a 9 pt
// label, but its paragraph mark is set larger: 9, 10.5, 12, 14 and 18 pt
// in Noto Sans JP, then 14 pt in Noto Sans. A line whose font height
// exceeds the pitch takes two pitches. Rules anchored to the margin every
// 18 pt in the right column show the pitch, so a paragraph that drifts off
// it is visible. Section B repeats the paragraphs with no grid: each line
// is as tall as its font.
const { LINE, colStart, span, emu } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const LA = `<w:rFonts w:ascii="${FONT.latin}" w:eastAsia="${FONT.latin}" w:hAnsi="${FONT.latin}"/>`;
const PITCH = 360;
const OUTLINE = `<a:ln w="3175"><a:solidFill><a:srgbClr val="${COLOR.accent}"/></a:solidFill></a:ln>`;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const ROWS = [
  ['9 pt', 18, JA], ['10.5 pt', 21, JA], ['12 pt', 24, JA], ['14 pt', 28, JA], ['18 pt', 36, JA], ['14 pt, Noto Sans', 28, LA],
];
const para = ([label, sz, font], runs = '') => ({
  kind: 'p',
  lineAuto: true,
  markRPr: `${font}<w:sz w:val="${sz}"/>`,
  runs: [`${runs}<w:r><w:rPr>${font}</w:rPr>${t(`段落記号 ${label}：この行の高さ`)}</w:r>`],
  height: NaN,
});
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

const parts = () => {
  let id = 1;
  // Rules from 108 pt to 504 pt below the top margin, every 18 pt.
  const rules = Array.from({ length: 23 }, (_, k) => wpsShape({
    id: id++, h: { from: 'margin', offset: emu(colStart(4)) }, v: { from: 'margin', offset: emu((6 + k) * PITCH) },
    cx: emu(span(4, 4)), cy: 0, prst: 'line', outline: OUTLINE,
  })).join('');
  return docxPackage([
    ...header({
      id: 'docx-044',
      title: 'Line heights on a document grid',
      verify: [
        'Verify: in A each line starts on an 18 pt rule; marks',
        'taller than 18 pt take two pitches. B, off the grid,',
        'sets each line at its font height.',
      ],
    }),
    { kind: 'section', type: 'continuous', docGrid: `<w:docGrid w:type="lines" w:linePitch="${PITCH}"/>` },
    caption('A. docGrid lines, linePitch 18 pt; rules every 18 pt at right', { at: DOCX.rowLine(2), before: 0 }),
    para(ROWS[0], rules),
    ...ROWS.slice(1).map((row) => para(row)),
    { kind: 'section', type: 'continuous' },
    caption('B. No grid: the same paragraphs'),
    ...ROWS.map((row) => para(row)),
  ], {
    ns: [A_NS, WP_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-044',
  slug: 'grid-line-heights',
  format: 'docx',
  title: 'Line heights on a document grid',
  covers: [
    'word/document.xml|w:docGrid@w:type=lines',
    'word/document.xml|w:docGrid@w:linePitch',
    'word/document.xml|w:pPr>w:rPr',
    'word/document.xml|w:spacing@w:lineRule=auto',
    'word/document.xml|wp:positionV@relativeFrom=margin',
  ],
  parts,
};
