import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS, wpsShape } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// How tall an empty paragraph is, on a line grid and off it. Each case is a
// label, two empty paragraphs whose mark has the case's font and size, and a
// marker line; the marker's position shows the height the two empty
// paragraphs took. One thing changes per case:
//   A. docGrid lines, pitch 18 pt: mark size 10, 11, 12 and 18 pt (Noto
//      Sans JP); then 12 pt with a floating shape anchored in the first
//      empty paragraph.
//   B. Pitch 20 pt: 11, 12 and 13 pt.
//   C. No grid: 11, 12 and 13 pt.
//   D. Pitch 18 pt, the mark in Noto Sans (no East Asian font) at 12 and
//      18 pt.
//   E. Pitch 18 pt, a mark with Latin and East Asian fonts and
//      w:hint="eastAsia", 12 pt.
const { emu, colStart, span, LINE } = DOCX;
const fonts = (ascii, ea, hint = '') => `<w:rFonts w:ascii="${ascii}" w:eastAsia="${ea}" w:hAnsi="${ascii}"${hint}/>`;
const JA = fonts(FONT.eastAsia, FONT.eastAsia);
const LA = fonts(FONT.latin, FONT.latin);
const HINT = fonts(FONT.latin, FONT.eastAsia, ' w:hint="eastAsia"');
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const label = (text) => ({ kind: 'p', style: 'caption', lines: [text] });
const grid = (pitch) => `<w:docGrid w:type="lines" w:linePitch="${pitch}"/>`;

let nextId = 1;
const empty = (font, size, runs = '') => ({ kind: 'p', lineAuto: true, markRPr: `${font}<w:sz w:val="${2 * size}"/>`, runs: [runs], height: NaN });
const marker = (name) => ({ kind: 'p', lineAuto: true, runs: [`<w:r>${t(`▲ ${name}: after two empty paragraphs`)}</w:r>`], height: NaN });
const kase = (name, font, size, anchor = false) => {
  const shape = anchor ? wpsShape({ id: nextId++, x: emu(colStart(4)), y: 0, cx: emu(span(4, 4)), cy: emu(2 * LINE), fill: COLOR.accent }) : '';
  return [label(`${name}: ${size} pt${anchor ? ', shape anchored in the first' : ''}`), empty(font, size, shape), empty(font, size), marker(name)];
};

const parts = () => {
  nextId = 1;
  return docxPackage([
    ...header({
      id: 'docx-051',
      title: 'Empty paragraphs on a line grid',
      verify: [
        'Verify: each marker sits where Word puts it: on the grid',
        'an empty mark taller than the pitch takes two pitches;',
        'off the grid it is as tall as its font. Compare the PDF.',
      ],
    }),
    { kind: 'section', type: 'continuous', docGrid: grid(360) },
    caption('A. Pitch 18 pt, Noto Sans JP marks', { at: DOCX.rowLine(2), before: 0 }),
    ...kase('A1', JA, 10), ...kase('A2', JA, 11), ...kase('A3', JA, 12), ...kase('A4', JA, 18), ...kase('A5', JA, 12, true),
    { kind: 'section', type: 'continuous', docGrid: grid(400) },
    caption('B. Pitch 20 pt'),
    ...kase('B1', JA, 11), ...kase('B2', JA, 12), ...kase('B3', JA, 13),
    { kind: 'section', type: 'continuous' },
    caption('C. No grid'),
    ...kase('C1', JA, 11), ...kase('C2', JA, 12), ...kase('C3', JA, 13),
    { kind: 'section', type: 'continuous', docGrid: grid(360) },
    caption('D. Pitch 18 pt, Noto Sans marks'),
    ...kase('D1', LA, 12), ...kase('D2', LA, 18),
    caption('E. Pitch 18 pt, hint eastAsia'),
    ...kase('E1', HINT, 12),
  ], {
    ns: [A_NS, WP_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-051',
  slug: 'empty-paragraph-grid',
  format: 'docx',
  title: 'Empty paragraphs on a line grid',
  covers: [
    'word/document.xml|w:docGrid@w:type=lines',
    'word/document.xml|w:docGrid@w:linePitch',
    'word/document.xml|w:pPr>w:rPr',
    'word/document.xml|w:rFonts@w:hint',
    'word/document.xml|w:spacing@w:lineRule=auto',
  ],
  parts,
};
