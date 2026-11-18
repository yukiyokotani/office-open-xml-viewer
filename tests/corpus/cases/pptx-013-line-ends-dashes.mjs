import { COLOR, PPTX } from '../lib/grid.mjs';
import { field, header, pptxPackage, solid, textBox, xfrm } from '../lib/pptx.mjs';

// Each rule runs over columns 1-6 on the baseline that closes its caption
// line in columns 7-12, so rule and caption share a line of the grid.
const W = 25400; // 2 pt
const lineShape = (id, y, lnInner, attrs = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xfrm({ x: PPTX.x(1), y, cx: PPTX.w(1, 6), cy: 0 })}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`
  + `<a:ln w="${W}"${attrs}>${solid(COLOR.ink)}${lnInner}</a:ln></p:spPr></p:sp>`;

const RULES = [
  ['prstDash dash', '<a:prstDash val="dash"/>', ''],
  ['prstDash sysDot, cap rnd', '<a:prstDash val="sysDot"/>', ' cap="rnd"'],
  ['prstDash lgDashDot', '<a:prstDash val="lgDashDot"/>', ''],
  ['tailEnd triangle, w med, len med', '<a:tailEnd type="triangle" w="med" len="med"/>', ''],
  ['headEnd oval, tailEnd stealth, w lg', '<a:headEnd type="oval" w="lg" len="lg"/><a:tailEnd type="stealth" w="lg" len="lg"/>', ''],
  ['headEnd diamond, tailEnd arrow, w sm', '<a:headEnd type="diamond" w="sm" len="sm"/><a:tailEnd type="arrow" w="sm" len="sm"/>', ''],
];

export default {
  id: 'pptx-013',
  slug: 'line-ends-dashes',
  format: 'pptx',
  title: 'Line ends and dashes',
  covers: [
    'ppt/slides/slide1.xml|a:ln>a:prstDash',
    'ppt/slides/slide1.xml|a:prstDash@val=dash',
    'ppt/slides/slide1.xml|a:prstDash@val=sysDot',
    'ppt/slides/slide1.xml|a:prstDash@val=lgDashDot',
    'ppt/slides/slide1.xml|a:ln@cap=rnd',
    'ppt/slides/slide1.xml|a:ln>a:headEnd',
    'ppt/slides/slide1.xml|a:headEnd@type=oval',
    'ppt/slides/slide1.xml|a:headEnd@type=diamond',
    'ppt/slides/slide1.xml|a:tailEnd@type=stealth',
    'ppt/slides/slide1.xml|a:tailEnd@type=arrow',
    'ppt/slides/slide1.xml|a:tailEnd@w=lg',
    'ppt/slides/slide1.xml|a:tailEnd@len=sm',
    'ppt/slides/slide1.xml|a:prstGeom@prst=line',
  ],
  parts: () => {
    const shapes = [];
    RULES.forEach(([label, inner, attrs], i) => {
      const top = PPTX.y(2) + 3 * i * PPTX.LINE; // a caption line every three baselines
      shapes.push(lineShape(3 + 2 * i, top + PPTX.LINE, inner, attrs));
      shapes.push(textBox(4 + 2 * i, `Caption ${i + 1}`, { x: PPTX.x(7), y: top, cx: PPTX.w(7, 12), cy: PPTX.LINE },
        [{ style: 'caption', text: `${label}; 2 pt` }]));
    });
    return pptxPackage([
      header({
        id: 'pptx-013',
        title: 'Line ends and dashes',
        verify: [
          'Verify: six 2 pt rules over columns 1-6, each dash pattern and end as its caption names;',
          'the rule lies on the baseline under its caption.',
        ],
      }),
      ...shapes,
    ]);
  },
};
