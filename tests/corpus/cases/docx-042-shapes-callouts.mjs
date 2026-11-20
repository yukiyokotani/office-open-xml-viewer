import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { A_NS, wpsShape } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// DrawingML shapes on a page, where geometry and stacking are the feature:
//   A. Callouts whose leader lines run far outside the box (adj values
//      below 0 and above 100000): borderCallout1, accentBorderCallout2 and
//      wedgeRectCallout, pointing left, up and down-right.
//   B. Two boxes joined by a straight and a curved connector
//      (straightConnector1, curvedConnector3), and the plus and minus signs
//      (mathPlus, mathMinus).
//   C. A number line drawn from shapes: a rule, ticks and labels in text
//      boxes, with three overlapping squares whose stacking order
//      (wp:anchor relativeHeight) is the reverse of their order in the
//      document. The square written last must be at the back.
//   D. A text box holding an equation (m:oMath) and ruby (w:ruby).
// Every shape is anchored to the margin and the paragraph of its block, in
// whole grid fields, wrapNone, so the text flow is unaffected.
const { LINE, colStart, span, emu } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
let nextId = 1;
const line = (n) => n * LINE;
const ln = `<a:ln w="12700"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln>`;
const p = (runs) => `<w:p>${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;

/** A shape anchored at (x, y) twips from the margin / paragraph, size cx × cy twips. */
function shape({ x, y, cx, cy, prst, adj = '', fillColor = null, outline = true, text = '', relH = 1, connector = false }) {
  const inset = text && outline ? 152400 : 0;
  return wpsShape({
    id: nextId++, x: emu(x), y: emu(y), cx: emu(cx), cy: emu(cy), prst, adj, fill: fillColor,
    outline: outline ? ln : '', text, insets: [inset, 0, inset, 0], relH, connector,
  });
}
const gd = (pairs) => pairs.map(([n, v]) => `<a:gd name="${n}" fmla="val ${v}"/>`).join('');
const COL = colStart(2);
const BOXW = span(1, 1);

// Anchor paragraphs: one per block, each holding its shapes, followed by
// empty lines that reserve the block's height.
const block = (shapes, lines) => [
  { kind: 'p', lines: [''], runs: [shapes.join('')], height: 1 },
  ...(lines > 1 ? [{ kind: 'p', lines: Array(lines - 1).fill(''), height: lines - 1 }] : []),
];
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

const parts = () => {
  nextId = 1;
  const A = [
    shape({ x: colStart(2), y: line(1), cx: BOXW, cy: line(2), prst: 'borderCallout1', text: p(r('leader to the left')),
      adj: gd([['adj1', 18750], ['adj2', -8333], ['adj3', 50000], ['adj4', -60000]]) }),
    shape({ x: colStart(3), y: line(2), cx: BOXW, cy: line(2), prst: 'accentBorderCallout2', text: p(r('leader upwards')),
      adj: gd([['adj1', 18750], ['adj2', -2129], ['adj3', -12506], ['adj4', -9355], ['adj5', -38227], ['adj6', 12678]]) }),
    shape({ x: colStart(4), y: line(0), cx: BOXW, cy: line(2), prst: 'wedgeRectCallout', text: p(r('wedge down-right')),
      adj: gd([['adj1', 20000], ['adj2', 150000]]) }),
  ];
  const B = [
    shape({ x: 0, y: 0, cx: BOXW, cy: line(2), prst: 'rect', text: p(r('box one')) }),
    shape({ x: colStart(3), y: line(2), cx: BOXW, cy: line(2), prst: 'rect', text: p(r('box two')) }),
    shape({ x: BOXW, y: line(1), cx: colStart(3) - BOXW, cy: line(2), prst: 'curvedConnector3', connector: true }),
    shape({ x: BOXW, y: line(1), cx: colStart(3) - BOXW, cy: 0, prst: 'straightConnector1', connector: true }),
    shape({ x: colStart(4), y: 0, cx: line(2), cy: line(2), prst: 'mathPlus', fillColor: COLOR.ink, outline: false }),
    shape({ x: colStart(4) + line(3), y: 0, cx: line(2), cy: line(2), prst: 'mathMinus', fillColor: COLOR.ink, outline: false }),
  ];
  const ticks = [-2, -1, 0, 1, 2].flatMap((v, i) => {
    const x = (i * DOCX.LIVE_W) / 4;
    return [
      shape({ x, y: line(1), cx: 0, cy: line(1), prst: 'line' }),
      shape({ x: x - line(1), y: line(2), cx: line(2), cy: line(1), prst: 'rect', outline: false, text: `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${r(String(v))}</w:p>` }),
    ];
  });
  const C = [
    shape({ x: 0, y: line(1), cx: DOCX.LIVE_W, cy: 0, prst: 'line' }),
    ...ticks,
    // Written front-to-back in the document, stacked back-to-front.
    shape({ x: colStart(2), y: line(4), cx: line(6), cy: line(3), prst: 'rect', fillColor: COLOR.accent, relH: 3, text: p(r('top')) }),
    shape({ x: colStart(2) + line(2), y: line(5), cx: line(6), cy: line(3), prst: 'rect', fillColor: COLOR.paper, relH: 2, text: p(r('middle')) }),
    shape({ x: colStart(2) + line(4), y: line(6), cx: line(6), cy: line(3), prst: 'rect', fillColor: COLOR.ink, relH: 1, text: p(r('back', `<w:color w:val="${COLOR.paper}"/>`)) }),
  ];
  const MATH = '<m:oMathPara><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:sSup><m:e><m:r><m:t>2</m:t></m:r></m:e><m:sup><m:r><m:t>n</m:t></m:r></m:sup></m:sSup></m:den></m:f>'
    + '<m:r><m:t>=</m:t></m:r><m:sSub><m:e><m:r><m:t>A</m:t></m:r></m:e><m:sub><m:r><m:t>n</m:t></m:r></m:sub></m:sSub></m:oMath></m:oMathPara>';
  const RUBY = `<w:p><w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="9"/><w:hpsRaise w:val="16"/><w:hpsBaseText w:val="18"/><w:lid w:val="ja-JP"/></w:rubyPr>`
    + `<w:rt><w:r><w:rPr>${JA}<w:sz w:val="9"/></w:rPr><w:t>ようし</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr>${JA}</w:rPr><w:t>用紙</w:t></w:r></w:rubyBase></w:ruby></w:r>${r('の面積', JA)}</w:p>`;
  const D = [shape({ x: 0, y: 0, cx: span(1, 2), cy: line(5), prst: 'rect', text: `${p(r('Area of An in m²:'))}${MATH}${RUBY}` })];

  return docxPackage([
    ...header({
      id: 'docx-042',
      title: 'Shapes, callouts and stacking',
      verify: [
        'Verify: A leaders reach far outside their boxes; B connectors',
        'join the boxes; C squares stack top, middle, back against',
        'document order; D shows the equation and ruby in the box.',
      ],
    }),
    caption('A. Callouts with leaders outside the box', { at: DOCX.rowLine(2), before: 0 }),
    ...block(A, 5),
    caption('B. Connectors and math signs'),
    ...block(B, 4),
    caption('C. Number line; relativeHeight against document order'),
    ...block(C, 10),
    caption('D. Text box with an equation and ruby'),
    ...block(D, 6),
  ], {
    ns: [A_NS,
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'],
  });
};

export default {
  id: 'docx-042',
  slug: 'shapes-callouts',
  format: 'docx',
  title: 'Shapes, callouts and stacking',
  covers: [
    'word/document.xml|a:prstGeom@prst=wedgeRectCallout',
    'word/document.xml|a:prstGeom@prst=mathPlus',
    'word/document.xml|a:prstGeom@prst=mathMinus',
    'word/document.xml|wps:wsp>wps:cNvCnPr',
    'word/document.xml|wp:anchor@relativeHeight',
    'word/document.xml|w:txbxContent>m:oMathPara',
    'word/document.xml|w:r>w:ruby',
  ],
  parts,
};
