import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Shared table edges where the two cells, and the table, disagree. Each case
// is a table of two rows; the target is the edge between them (in V, the
// vertical edge between the first two cells). One parameter changes at a
// time:
//   the upper cell's bottom border: nil, none, omitted, or explicit;
//   the lower cell's top border: omitted or explicit;
//   the table's insideH: thin black, or absent;
//   the topology: three cells over three, one spanning cell over three,
//   one over one plus a span of two, and three cells over one span.
// Explicit cell borders are the accent at 2.25 pt, so on every segment of
// the target edge the reader can tell a cell border (thick accent), the
// table's insideH (thin black) and no line apart.
const { LINE, colStart, span } = DOCX;
const GRID = [colStart(2), colStart(3) - colStart(2), span(3, 4)];
const THIN = (n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`;
const SIDE = {
  nil: (n) => `<w:${n} w:val="nil"/>`,
  none: (n) => `<w:${n} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
  explicit: (n) => `<w:${n} w:val="single" w:sz="18" w:space="0" w:color="${COLOR.accent}"/>`,
  omitted: () => '',
};
const cell = (cols, side, label, extra = '') => {
  const w = GRID.slice(cols[0], cols[1] + 1).reduce((a, b) => a + b, 0);
  const n = cols[1] - cols[0] + 1;
  const b = side ? side() : '';
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${n > 1 ? `<w:gridSpan w:val="${n}"/>` : ''}`
    + `${b ? `<w:tcBorders>${b}</w:tcBorders>` : ''}${extra}</w:tcPr><w:p><w:r><w:t xml:space="preserve">${esc(label)}</w:t></w:r></w:p></w:tc>`;
};
const row = (cells) => `<w:tr><w:trPr><w:trHeight w:val="${LINE}" w:hRule="exact"/></w:trPr>${cells}</w:tr>`;
const table = (insideH, upper, lower, insideV = true) => '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${GRID.reduce((a, b) => a + b, 0)}" w:type="dxa"/>`
  + `<w:tblBorders>${THIN('top')}${THIN('left')}${THIN('bottom')}${THIN('right')}${insideH ? THIN('insideH') : ''}${insideV ? THIN('insideV') : ''}</w:tblBorders>`
  + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid>${GRID.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${row(upper)}${row(lower)}</w:tbl>`;

const bottom = (k) => () => SIDE[k]('bottom');
const top = (k) => () => SIDE[k]('top');
const three = (sides, pos, verb) => sides.map((k, i) => cell([i, i], pos(k), `${verb} ${k}`)).join('');
const spanned = (k, pos, verb) => cell([0, 2], pos(k), `${verb} ${k} (spans 3)`);

// [label, insideH, upper cells, lower cells]
const CASES = [
  ['A1 three over three: nil / omitted, insideH', true, three(['nil', 'nil', 'nil'], bottom, 'bottom'), three(['omitted', 'omitted', 'omitted'], top, 'top')],
  ['A2 three over three: nil / explicit, insideH', true, three(['nil', 'nil', 'nil'], bottom, 'bottom'), three(['explicit', 'explicit', 'explicit'], top, 'top')],
  ['A3 three over three: none / omitted, insideH', true, three(['none', 'none', 'none'], bottom, 'bottom'), three(['omitted', 'omitted', 'omitted'], top, 'top')],
  ['A4 three over three: omitted / omitted, insideH (control)', true, three(['omitted', 'omitted', 'omitted'], bottom, 'bottom'), three(['omitted', 'omitted', 'omitted'], top, 'top')],
  ['A5 three over three: nil / omitted, no insideH', false, three(['nil', 'nil', 'nil'], bottom, 'bottom'), three(['omitted', 'omitted', 'omitted'], top, 'top')],
  ['B1 span over three: nil / omitted, insideH', true, spanned('nil', bottom, 'bottom'), three(['omitted', 'omitted', 'omitted'], top, 'top')],
  ['B2 span over three: nil / nil, omitted, omitted', true, spanned('nil', bottom, 'bottom'), three(['nil', 'omitted', 'omitted'], top, 'top')],
  ['B3 span over three: nil / explicit, insideH', true, spanned('nil', bottom, 'bottom'), three(['explicit', 'explicit', 'explicit'], top, 'top')],
  ['C1 span over one + span of two: nil / nil, explicit', true, spanned('nil', bottom, 'bottom'), cell([0, 0], top('nil'), 'top nil') + cell([1, 2], top('explicit'), 'top explicit (spans 2)')],
  ['D1 three over span: omitted / nil, insideH', true, three(['omitted', 'omitted', 'omitted'], bottom, 'bottom'), spanned('nil', top, 'top')],
  ['D2 three over span: explicit / nil, insideH', true, three(['explicit', 'explicit', 'explicit'], bottom, 'bottom'), spanned('nil', top, 'top')],
  ['V1 vertical: right nil / left omitted, insideV', true,
    cell([0, 0], () => SIDE.nil('right'), 'right nil') + cell([1, 1], null, 'left omitted') + cell([2, 2], null, 'third cell'),
    cell([0, 0], () => SIDE.nil('right'), 'right nil') + cell([1, 1], () => SIDE.explicit('left'), 'left explicit') + cell([2, 2], null, 'third cell')],
];

const blocks = CASES.flatMap(([label, insideH, upper, lower], i) => [
  { kind: 'p', style: 'caption', bold: true, lines: [label], ...(i === 0 ? { at: DOCX.rowLine(2) } : {}) },
  { kind: 'raw', xml: table(insideH, upper, lower), height: 2 },
]);

export default {
  id: 'docx-040',
  slug: 'border-conflicts',
  format: 'docx',
  title: 'Table border conflicts',
  covers: [
    'word/document.xml|w:tcPr>w:tcBorders',
    'word/document.xml|w:bottom@w:val=nil',
    'word/document.xml|w:bottom@w:val=none',
    'word/document.xml|w:top@w:val=single',
    'word/document.xml|w:tblBorders>w:insideH',
    'word/document.xml|w:tblBorders>w:insideV',
    'word/document.xml|w:tcPr>w:gridSpan',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-040',
      title: 'Table border conflicts',
      verify: [
        'Verify: each edge between the rows shows the thick',
        'accent, the thin black insideH or no line, as Word',
        'resolves the cells and the table; see the PDF.',
      ],
    }),
    ...blocks,
  ]),
};
