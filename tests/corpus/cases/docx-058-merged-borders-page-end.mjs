import { DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, W, WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';
import { anchor, p, r } from '../lib/wml.mjs';

// Merged-cell borders, empty paragraphs with split fonts, and the page end.
//   A. A cell merged down three rows (w:vMerge) whose cells all have a
//      bottom border of 0.5, 1.5 or 3 pt (one table per width).
//   B. Two empty paragraphs whose mark has a Latin font different from the
//      East Asian one, at 12 and 18 pt, on a 20 pt line grid (page 1) and
//      without a grid; the marker after them shows their height.
//   R1-R3 (pages 2-7), 20 pt line grid: a 63.5 pt exact paragraph and 32
//      one-line paragraphs leave 16.5 pt, then an empty paragraph and a
//      marker. R2 adds a square-wrapped box near the top of the page; R3 a
//      box that overlaps the empty paragraph's line.
//   S (last page): a smoke test, a table of 58 exact 12 pt rows with a
//      footnote reference in the last row; it must lay out without error.
const { emu, LINE, colStart, span } = DOCX;
const GRID = '<w:docGrid w:type="lines" w:linePitch="400"/>';
const SPLIT = `<w:rFonts w:ascii="${FONT.latin}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.latin}"/>`;
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, height: NaN, ...extra });
const line = (text, extra = {}) => ({ kind: 'p', lineAuto: true, runs: [r(text)], height: NaN, ...extra });
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });

const BORDER = (sz) => `<w:bottom w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/>`;
const OUTER = '<w:tblBorders>' + ['top', 'left', 'right', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const mergedTable = (sz) => {
  const W1 = span(1, 1);
  const W2 = span(2, 2);
  const row = (i) => '<w:tr>'
    + `<w:tc><w:tcPr><w:tcW w:w="${W1}" w:type="dxa"/><w:vMerge${i === 1 ? ' w:val="restart"' : ''}/><w:tcBorders>${BORDER(sz)}</w:tcBorders></w:tcPr>${p('', i === 1 ? r(`merged, bottom ${sz / 8} pt`) : '')}</w:tc>`
    + `<w:tc><w:tcPr><w:tcW w:w="${W2}" w:type="dxa"/><w:tcBorders>${BORDER(sz)}</w:tcBorders></w:tcPr>${p('', r(`row ${i}`))}</w:tc></w:tr>`;
  return '<w:tbl><w:tblPr>'
    + `<w:tblW w:w="${W1 + W2}" w:type="dxa"/>${OUTER}<w:tblLayout w:type="fixed"/>`
    + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    + `<w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W2}"/></w:tblGrid>${row(1)}${row(2)}${row(3)}</w:tbl>`;
};

const empties = (name, size) => [
  { kind: 'p', style: 'caption', lines: [`${name}: two empty paragraphs, mark ${size} pt, Latin ≠ East Asian`] },
  { kind: 'p', lineAuto: true, markRPr: `${SPLIT}<w:sz w:val="${2 * size}"/>`, runs: [''], height: NaN },
  { kind: 'p', lineAuto: true, markRPr: `${SPLIT}<w:sz w:val="${2 * size}"/>`, runs: [''], height: NaN },
  line(`▲ ${name}: after the two empty paragraphs`),
];

let nextId = 1;
const box = (v, c) => anchor({
  id: nextId++, h: { from: 'margin', offset: emu(colStart(c)) }, v, cx: emu(span(c, c)), cy: emu(LINE), wrap: '<wp:wrapSquare wrapText="bothSides"/>',
});
const pageEnd = (name, topBox, endBox) => [
  { kind: 'p', runs: [r(`${name}: a 63.5 pt exact paragraph, 32 lines, an empty paragraph, a marker.`)], pPrMid: '', height: NaN, rPr: '' },
  ...Array.from({ length: 32 }, (_, i) => line(`${name} line ${String(i + 1).padStart(2, '0')}`, i === 0 && topBox ? { runs: [box({ from: 'paragraph', offset: 0 }, 4) + r(`${name} line 01, box at the right`)] }
    : i === 31 && endBox ? { runs: [box({ from: 'paragraph', offset: emu(400) }, 1) + r(`${name} line 32, box on the next line`)] } : {})),
  { kind: 'p', lineAuto: true, runs: [''], height: NaN },
  line(`▲ ${name}: after the empty paragraph`),
];
// The first paragraph of each page-end case is 63.5 pt exact (1270 twips).
const exactFirst = (blocks) => { blocks[0] = { kind: 'raw', height: NaN, xml: p('<w:spacing w:line="1270" w:lineRule="exact"/>', blocks[0].runs.join('')) }; return blocks; };

const NOTE_REF = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>';
const FOOTNOTES = XML_DECL + `<w:footnotes ${W}>`
  + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
  + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  + `<w:footnote w:id="1"><w:p><w:r>${NOTE_REF}<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> The note of the last row.</w:t></w:r></w:p></w:footnote>`
  + '</w:footnotes>';
const smokeRows = Array.from({ length: 58 }, (_, i) => ({
  lines: 1,
  cells: [{ runs: [r(`S row ${String(i + 1).padStart(2, '0')}`), ...(i === 57 ? [`<w:r>${NOTE_REF}<w:footnoteReference w:id="1"/></w:r>`] : [])] }],
}));

const parts = () => {
  nextId = 1;
  // A page-end case must start a page and close with a paragraph block, so
  // its section break carries the next case's properties.
  return docxPackage([
    ...header({
      id: 'docx-058',
      title: 'Merged borders, page end',
      verify: [
        'Verify: merged cells keep room for their bottom border;',
        'B markers sit as in the PDF; in R1 and R2 the empty',
        'paragraph stays on its page, in R3 it moves.',
      ],
    }),
    caption('A. Cell merged down three rows, bottom borders', { at: DOCX.rowLine(2), before: 0 }),
    raw(mergedTable(4)), line(''), raw(mergedTable(12)), line(''), raw(mergedTable(24)),
    caption('B. Empty paragraphs on a 20 pt grid'),
    ...empties('B1', 12), ...empties('B2', 18),
    { kind: 'section', type: 'continuous' },
    caption('B. Without a grid'),
    ...empties('B3', 12), ...empties('B4', 18),
    { kind: 'section', type: 'nextPage', docGrid: GRID },
    ...exactFirst(pageEnd('R1', false, false)),
    { kind: 'section', type: 'nextPage', docGrid: GRID },
    ...exactFirst(pageEnd('R2', true, false)),
    { kind: 'section', type: 'nextPage', docGrid: GRID },
    ...exactFirst(pageEnd('R3', false, true)),
    { kind: 'section', type: 'nextPage' },
    { kind: 'table', grid: [span(1, 2)], rows: smokeRows },
    line('S: after the table.'),
  ], {
    firstSection: { docGrid: GRID },
    settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>',
    parts: { 'word/footnotes.xml': FOOTNOTES },
    rels: [['rId3', 'footnotes', 'footnotes.xml']],
    overrides: { 'word/footnotes.xml': `${WML}.footnotes+xml` },
    ns: [R_NS, WP_NS, A_NS, WPS_NS, MC_NS],
  });
};

export default {
  id: 'docx-058',
  slug: 'merged-borders-page-end',
  format: 'docx',
  title: 'Merged borders, page end',
  covers: [
    'word/document.xml|w:tcPr>w:vMerge',
    'word/document.xml|w:tcBorders>w:bottom',
    'word/document.xml|w:docGrid@w:linePitch',
    'word/document.xml|w:pPr>w:rPr',
    'word/document.xml|wp:anchor>wp:wrapSquare',
    'word/document.xml|w:r>w:footnoteReference',
  ],
  parts,
};
