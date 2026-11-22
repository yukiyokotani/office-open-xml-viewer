import { DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { p, r, spacing } from '../lib/wml.mjs';

// An empty paragraph mark (a w:p with no runs) that is the last paragraph of
// a full page, without a document grid. One case per page: an exact first
// paragraph and 58 exact 12 pt lines leave a known room above the bottom
// margin; then the empty mark, then a marker paragraph. The marker shows
// where the mark went: on the first line of the next page when the mark
// stayed, on the second when the mark moved.
//   The mark is 9 pt Noto Sans (a single line of about 12.3 pt), or 9 pt
//   Noto Sans JP in case G (Word gives it a line of about 17 pt).
//   A1/A2: single (auto 240), room 13 pt / 11 pt.
//   B, C, D: auto 276, 360 and 480, room 13 pt: only the first line has to
//     fit; the extra leading may hang past the margin.
//   E1/E2: exact 12 pt, room 13 pt / 11 pt.
//   F1/F2: at least 12 pt, room 13 pt / 11 pt.
//   G1/G2: single, Noto Sans JP mark, room 18 pt / 16 pt.
const { LINE, LINES } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const FILL = 58;
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });

const CASES = [
  ['A1', spacing(240, 'auto'), '', 260],
  ['A2', spacing(240, 'auto'), '', 220],
  ['B', spacing(276, 'auto'), '', 260],
  ['C', spacing(360, 'auto'), '', 260],
  ['D', spacing(480, 'auto'), '', 260],
  ['E1', spacing(240, 'exact'), '', 260],
  ['E2', spacing(240, 'exact'), '', 220],
  ['F1', spacing(240, 'atLeast'), '', 260],
  ['F2', spacing(240, 'atLeast'), '', 220],
  ['G1', spacing(240, 'auto'), `<w:rPr>${JA}</w:rPr>`, 360],
  ['G2', spacing(240, 'auto'), `<w:rPr>${JA}</w:rPr>`, 320],
];
const label = (spc) => spc.match(/w:line="(\d+)" w:lineRule="(\w+)"/).slice(1).join(' ');

const page = ([name, spc, markRPr, room]) => [
  // Exact first paragraph: the page holds it, FILL lines and `room` twips.
  raw(p(spacing(LINES * LINE - FILL * LINE - room, 'exact'), r(`${name}: mark ${label(spc)}, room ${room / 20} pt`))),
  ...Array.from({ length: FILL }, (_, i) => raw(p('', r(`${name} line ${String(i + 1).padStart(2, '0')}`)))),
  raw(`<w:p><w:pPr>${spc}${markRPr}</w:pPr></w:p>`),
  { kind: 'p', lines: [`▲ ${name}: after the empty mark`], height: NaN },
];

export default {
  id: 'docx-064',
  slug: 'empty-mark-page-end',
  format: 'docx',
  title: 'Empty mark at the page end',
  covers: [
    'word/document.xml|w:body>w:p',
    'word/document.xml|w:spacing@w:lineRule=auto',
    'word/document.xml|w:spacing@w:lineRule=exact',
    'word/document.xml|w:spacing@w:lineRule=atLeast',
    'word/document.xml|w:pPr>w:rPr',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-064',
      title: 'Empty mark at the page end',
      verify: [
        'Verify: a marker on line 1 of the next page means the',
        'empty mark stayed (its first line fitted the room);',
        'on line 2, it moved. Compare the PDF.',
      ],
    }),
    { kind: 'p', style: 'caption', lines: ['One case per page, no document grid.'], at: DOCX.rowLine(2) },
    ...CASES.flatMap((c) => [{ kind: 'section', type: 'nextPage' }, ...page(c)]),
  ]),
};
