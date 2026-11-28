import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { A_SERIES, B_SERIES } from '../lib/iso216.mjs';
import { esc } from '../lib/xml.mjs';

// Tables whose geometry Word computes: percentage widths, columns sized by
// their content (autofit never narrower than the widest unbreakable token),
// a table nested in a cell, centred cells, and a long table whose header row
// repeats on the next page while no row is split.
const { colStart, span, LINE } = DOCX;
const MODULES = [colStart(2) - colStart(1), colStart(3) - colStart(2), colStart(4) - colStart(3), span(4, 4)];
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const p = (text, { bold = false, jc } = {}) => `<w:p>${jc ? `<w:pPr><w:jc w:val="${jc}"/></w:pPr>` : ''}`
  + `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
const tc = (w, body, extra = '') => `<w:tc><w:tcPr>${w}${extra}</w:tcPr>${body}</w:tc>`;
const pct = (fiftieths) => `<w:tcW w:w="${fiftieths}" w:type="pct"/>`;
const dxa = (twips) => `<w:tcW w:w="${twips}" w:type="dxa"/>`;
const AUTO = '<w:tcW w:w="0" w:type="auto"/>';
const tr = (cells, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
const tbl = (tblPr, grid, rows) => `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
const caption = (text) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1 });

// A: 100 % wide; columns 20 / 30 / 50 %; a centred header row of two baselines.
const CENTRE = '<w:vAlign w:val="center"/>';
const TABLE_A = tbl(`<w:tblW w:w="5000" w:type="pct"/>${BORDERS}`, [MODULES[0], MODULES[1], MODULES[2] + MODULES[3]], [
  tr(tc(pct(1000), p('20 %', { bold: true, jc: 'center' }), CENTRE) + tc(pct(1500), p('30 %', { bold: true, jc: 'center' }), CENTRE)
    + tc(pct(2500), p('50 %, centred both ways in a row of two baselines', { bold: true, jc: 'center' }), CENTRE), `<w:trHeight w:val="${2 * LINE}" w:hRule="atLeast"/>`),
  tr(tc(pct(1000), p('A4')) + tc(pct(1500), p('210 × 297 mm')) + tc(pct(2500), p('Letters and printed pages'))),
].join(''));

// B: autofit to contents. Preferred widths are small; the dates and the
// reference may not break, so their columns grow to the widest token.
const TABLE_B = tbl(`<w:tblW w:w="0" w:type="auto"/>${BORDERS}`, [MODULES[0], MODULES[1], MODULES[2]], [
  tr(tc(AUTO, p('Size, dates and reference (gridSpan 3)', { bold: true }), '<w:gridSpan w:val="3"/>')),
  tr(tc(dxa(LINE * 2), p('A4')) + tc(dxa(LINE * 2), p('2026-02-28')) + tc(dxa(LINE * 2), p('ISO-216:2007/A4'))),
  tr(tc(dxa(LINE * 2), p('B5')) + tc(dxa(LINE * 2), p('2026-11-19')) + tc(dxa(LINE * 2), p('ISO-216:2007/B5'))),
].join(''));

// C: a table nested in the right cell of an outer two-column table.
const NESTED = tbl(`<w:tblW w:w="0" w:type="auto"/>${BORDERS}`, [MODULES[2], MODULES[3]], [
  tr(tc(dxa(MODULES[2]), p('Inner 1')) + tc(dxa(MODULES[3]), p('Inner 2'))),
  tr(tc(dxa(MODULES[2]), p('Inner 3')) + tc(dxa(MODULES[3]), p('Inner 4'))),
].join(''));
const TABLE_C = tbl(`<w:tblW w:w="${DOCX.LIVE_W}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`, [span(1, 2) + DOCX.gutterAfter(2), span(3, 4)], [
  tr(tc(dxa(span(1, 2) + DOCX.gutterAfter(2)), p('The right cell holds a table of two by two cells. A cell must end with a paragraph, so an empty one follows it.'))
    + tc(dxa(span(3, 4)), NESTED + '<w:p/>')),
].join(''));

// D: long table. The first row repeats (w:tblHeader); no row splits (w:cantSplit).
const SERIES = [...A_SERIES, ...B_SERIES];
const ROWS_D = [
  tr(tc(dxa(MODULES[0]), p('Size', { bold: true })) + tc(dxa(MODULES[1]), p('Millimetres', { bold: true })) + tc(dxa(MODULES[2] + MODULES[3]), p('Note', { bold: true })), '<w:cantSplit/><w:tblHeader/>'),
  ...SERIES.map(([n, w, h], i) => tr(tc(dxa(MODULES[0]), p(n)) + tc(dxa(MODULES[1]), p(`${w} × ${h}`))
    + tc(dxa(MODULES[2] + MODULES[3]), i % 4 === 3
      ? p(`${n} is half of the size before it. This note runs to several lines in its cell, and the row is marked w:cantSplit, so when it does not fit at the foot of a page the whole row moves to the next page, below the repeated header row.`)
      : p(`${n}: ratio ${(h / w).toFixed(3)}`)), '<w:cantSplit/>')),
];
const TABLE_D = tbl(`<w:tblW w:w="${DOCX.LIVE_W}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`, [MODULES[0], MODULES[1], MODULES[2] + MODULES[3]], ROWS_D.join(''));

export default {
  id: 'docx-021',
  slug: 'table-layout',
  format: 'docx',
  title: 'Table layout',
  covers: [
    'word/document.xml|w:tblW@w:type=pct',
    'word/document.xml|w:tcW@w:type=pct',
    'word/document.xml|w:tcW@w:type=auto',
    'word/document.xml|w:tblW@w:type=auto',
    'word/document.xml|w:tcPr>w:vAlign',
    'word/document.xml|w:trHeight@w:hRule=atLeast',
    'word/document.xml|w:tc>w:tbl',
    'word/document.xml|w:trPr>w:tblHeader',
    'word/document.xml|w:trPr>w:cantSplit',
    'word/document.xml|w:tcPr>w:gridSpan',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-021',
      title: 'Table layout',
      verify: [
        'Verify: A spans the measure at 20/30/50 %; B grows its',
        'columns so no date or reference breaks; C nests a table;',
        'D repeats its header row on page 2 and splits no row.',
      ],
    }),
    caption('A. Percentage widths, autofit'),
    { kind: 'raw', xml: TABLE_A },
    caption('B. Widths from content'),
    { kind: 'raw', xml: TABLE_B },
    caption('C. Nested table'),
    { kind: 'raw', xml: TABLE_C },
    caption('D. Long table with a repeated header row'),
    { kind: 'raw', xml: TABLE_D },
    { kind: 'p', style: 'caption', text: 'End of table D.', before: 1 },
  ]),
};
