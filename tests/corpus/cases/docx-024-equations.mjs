import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';
import {
  M_NS, MATH, SP, acc, aln, bar, borderBox, box, d, eqArr, f, func, groupChr, limLow, limUpp, matrix, nary, oMath, oMathPara,
  phant, pre, r, rad, sub, subSup, sup,
} from '../lib/omml.mjs';

// Every Office Math object, operator family, accent, delimiter pair, math
// alphabet and display alignment Word's equation editor writes, one family
// per block: a caption naming the objects, then display equations (one
// m:oMath per line). Inline equations, alignment of whole displays and math
// in centred table cells follow. The formulas are standard identities.
MATH.sz = 24; // 12 pt, so accents and stretched symbols are legible
const x = 'x';
const n = 'n';
const caption = (text) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1 });
// One equation per paragraph: each display is its own m:oMathPara.
const display = (lines, jc) => lines.map((line) => ({ kind: 'p', lineAuto: true, runs: [oMathPara([line], jc)] }));
const ABC = 'abc';

const BLOCKS = [
  ['Fractions: m:f bar, skw, lin, noBar; nested; long numerator', [
    [f('a', 'b'), r(SP), f('a', 'b', 'skw'), r(SP), f('a', 'b', 'lin'), r(SP), d(f('n', 'k', 'noBar'))],
    [r('x='), f([r('−b±'), rad([sup('b', '2'), r('−4ac')])], '2a')],
    [f('1', [r('1+'), f('1', [r('1+'), f('1', 'x')])])],
  ]],
  ['Scripts: m:sSub, m:sSup, m:sSubSup, m:sPre; nested scripts', [
    [sub('a', 'n'), r(SP), sup('x', '2'), r(SP), subSup('x', 'i', '2'), r(SP), pre('1', '2', 'F'), r(SP), sup('e', sup('x', '2'))],
  ]],
  ['Radicals: m:rad with and without degree; nested; over a tall fraction', [
    [rad('2'), r(SP), rad('8', '3'), r(SP), rad([r('1+'), rad([r('1+'), rad('x')])]), r(SP), rad(f('a', 'b'))],
  ]],
  ['Large operators: m:nary, limits above and below (undOvr)', [
    ['∑', '∏', '∐', '⋃', '⋂', '⋀', '⋁', '⨀', '⨁', '⨂'].flatMap((chr) => [
      nary(chr, sub('a', 'k'), { limLoc: 'undOvr', sub: 'k=1', sup: 'n' }), r(SP),
    ]),
  ]],
  ['Integrals: m:nary, limits as scripts (subSup); hidden limits', [
    [nary('∫', [sup('x', '2'), r('dx')], { limLoc: 'subSup', sub: '0', sup: '1' }), r('=') , f('1', '3'), r(SP),
      ...['∬', '∭', '∮', '∯', '∰'].flatMap((chr) => [nary(chr, 'f', { limLoc: 'subSup', sub: 'D' }), r(SP)]),
      nary('∫', [r('f'), d('x'), r('dx')], {})],
  ]],
  ['Delimiters: m:d with every bracket pair, one-sided, separators, growing', [
    [d('x'), d('x', '[', ']'), d('x', '{', '}'), d('x', '⟨', '⟩'), d('x', '|', '|'), d('x', '‖', '‖'),
      d('x', '⌊', '⌋'), d('x', '⌈', '⌉'), d('x', '⟦', '⟧')],
    [r('f'), d(x), r('='), d([[r('x'), r('if '), r('x≥0')]], '{', ''), r(SP), d([['a'], ['b'], ['c']], '⟨', '⟩', { sep: '|' }),
      r(SP), d(f([r('a+b')], 'c'), '[', ']')],
  ]],
  ['Accents: m:acc (dot, two and three dots, hat, caron, tilde, breve, acute, grave, bar, double bar, vector); m:bar', [
    [['̇', '̈', '⃛', '̂', '̌', '̃', '̆', '́', '̀', '̅', '̿', '⃗'].flatMap((chr) => [acc(chr, 'a'), r(SP)]),
      bar('AB'), r(SP), bar('x', 'bot')],
  ]],
  ['Wide symbols over long bases: m:groupChr braces and arrows, wide hat and bar', [
    [groupChr('⏞', [r('a+b+c+d')], 'top', 'bot'), r(SP), groupChr('⏟', [r('a+b+c+d')], 'bot', 'top'), r(SP),
      groupChr('⏜', [r('x+y+z')], 'top', 'bot'), r(SP), groupChr('⏝', [r('x+y+z')], 'bot', 'top')],
    [groupChr('→', [r('x+y+z+w')], 'top', 'bot'), r(SP), groupChr('←', [r('x+y+z+w')], 'bot', 'top'), r(SP),
      acc('̂', [r('x+y+z')]), r(SP), bar([r('x+y+z')]), r(SP), acc('⃗', [r('AB+BC')])],
  ]],
  ['Functions and limits: m:func, m:limLow, m:limUpp', [
    [func('sin', sup('θ', '2')), r('+'), func('cos', sup('θ', '2')), r('=1'), r(SP), func(limLow(r('lim', { sty: 'p' }), [r('n→∞')]), sup(d([r('1+'), f('1', 'n')]), 'n')), r('=e')],
    [func(limLow(r('max', { sty: 'p' }), [r('x∈S')]), [r('f'), d(x)]), r(SP), limUpp([r('A')], [r('def')]), r(SP), func('log', sub('b', 'x'))],
  ]],
  ['Matrices and arrays: m:m with column alignment, m:eqArr with alignment points', [
    [d(matrix([['a', 'b'], ['c', 'd']])), r(SP), d(matrix([['1', '0', '0'], ['0', '1', '0'], ['0', '0', '1']], ['left', 'center', 'right']), '[', ']'), r(SP),
      d(eqArr([[r('x+y'), aln(), r('3')], [r('x−y'), aln(), r('1')]]), '{', '')],
  ]],
  ['Boxes and phantoms: m:box, m:borderBox, m:phant', [
    [box([r('a+b')]), r(SP), borderBox([r('E=m'), sup('c', '2')]), r(SP), r('a'), phant([r('xxx')]), r('b')],
  ]],
  ['Math alphabets and styles: m:scr, m:sty, m:nor', [
    [['roman', 'script', 'fraktur', 'double-struck', 'sans-serif', 'monospace'].flatMap((scr) => [r(ABC, { scr }), r(SP)])],
    [r('x', { sty: 'p' }), r(SP), r('x', { sty: 'i' }), r(SP), r('x', { sty: 'b' }), r(SP), r('x', { sty: 'bi' }), r(SP), r('for all real x', { nor: true })],
  ]],
];

const CELL = (content) => `<w:tc><w:tcPr><w:tcW w:w="${DOCX.span(1, 2) + DOCX.gutterAfter(2)}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>`
  + `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${content}</w:p></w:tc>`;
const BORDER = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const TABLE = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${DOCX.LIVE_W}" w:type="dxa"/>${BORDER}<w:tblLayout w:type="fixed"/></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="${DOCX.span(1, 2) + DOCX.gutterAfter(2)}"/><w:gridCol w:w="${DOCX.span(3, 4)}"/></w:tblGrid>`
  + `<w:tr><w:trPr><w:trHeight w:val="${4 * DOCX.LINE}" w:hRule="atLeast"/></w:trPr>`
  + CELL(oMath(f('a', 'b'))) + CELL(oMath(nary('∑', sub('a', 'k'), { limLoc: 'undOvr', sub: 'k=1', sup: 'n' }))).replace(`w:w="${DOCX.span(1, 2) + DOCX.gutterAfter(2)}" w:type="dxa"/><w:vAlign`, `w:w="${DOCX.span(3, 4)}" w:type="dxa"/><w:vAlign`)
  + '</w:tr></w:tbl>';

export default {
  id: 'docx-024',
  slug: 'equations',
  format: 'docx',
  title: 'Equations',
  covers: [
    'word/document.xml|w:p>m:oMathPara',
    'word/document.xml|w:p>m:oMath',
    'word/document.xml|m:fPr>m:type',
    'word/document.xml|m:oMath>m:sPre',
    'word/document.xml|m:radPr>m:degHide',
    'word/document.xml|m:naryPr>m:limLoc',
    'word/document.xml|m:naryPr>m:subHide',
    'word/document.xml|m:dPr>m:sepChr',
    'word/document.xml|m:accPr>m:chr',
    'word/document.xml|m:groupChrPr>m:vertJc',
    'word/document.xml|m:oMath>m:func',
    'word/document.xml|m:fName>m:limLow',
    'word/document.xml|m:oMath>m:limUpp',
    'word/document.xml|m:mcPr>m:mcJc',
    'word/document.xml|m:e>m:eqArr',
    'word/document.xml|m:oMath>m:box',
    'word/document.xml|m:oMath>m:borderBox',
    'word/document.xml|m:oMath>m:phant',
    'word/document.xml|m:rPr>m:scr',
    'word/document.xml|m:rPr>m:sty',
    'word/document.xml|m:rPr>m:nor',
    'word/document.xml|m:rPr>m:aln',
    'word/document.xml|m:oMathParaPr>m:jc',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-024',
      title: 'Equations',
      verify: [
        'Verify: each caption names the objects in the display below',
        'it; no symbol is missing or drawn as a box; wide braces and',
        'arrows stretch along their bases, not upwards.',
      ],
    }),
    ...BLOCKS.flatMap(([title, lines]) => [caption(title), ...display(lines)]),
    caption('Inline equations: m:oMath inside a paragraph'),
    {
      kind: 'p',
      lineAuto: true,
      runs: [run('The roots of '), oMath(sup('x', '2'), r('+px+q=0')), run(' are '), oMath(r('x='), f([r('−p±'), rad([sup('p', '2'), r('−4q')])], '2')),
        run('; the line containing a built-up fraction grows to fit it, and the next line returns to the normal spacing of the paragraph.')],
    },
    caption('Display alignment: m:oMathParaPr m:jc left, center, right, centerGroup'),
    ...['left', 'center', 'right', 'centerGroup'].flatMap((jc) => display([[r(`${jc}: `, { nor: true }), sup('e', [r('iπ')]), r('+1=0')]], jc)),
    caption('Equations in table cells centred both ways'),
    { kind: 'raw', xml: TABLE },
  ], { ns: [M_NS] }),
};
