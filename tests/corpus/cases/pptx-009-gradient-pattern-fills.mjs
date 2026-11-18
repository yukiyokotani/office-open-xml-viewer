import { COLOR } from '../lib/grid.mjs';
import { NO_LINE, field, header, pptxPackage, shape, textBox } from '../lib/pptx.mjs';

const stop = (pos, rgb) => `<a:gs pos="${pos}"><a:srgbClr val="${rgb}"/></a:gs>`;
const FILLS = [
  {
    name: 'Linear gradient',
    xml: `<a:gradFill rotWithShape="1"><a:gsLst>${stop(0, COLOR.ink)}${stop(100000, COLOR.paper)}</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>`,
    lines: ['gradFill, a:lin ang 0', `stops 0 % ${COLOR.ink}, 100 % ${COLOR.paper}`, 'black at the left edge'],
  },
  {
    name: 'Linear gradient, 90°',
    xml: `<a:gradFill rotWithShape="1"><a:gsLst>${stop(0, COLOR.accent)}${stop(50000, COLOR.paper)}${stop(100000, COLOR.ink)}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`,
    lines: ['gradFill, a:lin ang 5400000 (90°)', `stops 0 % ${COLOR.accent}, 50 % ${COLOR.paper},`, `100 % ${COLOR.ink}: orange at the top`],
  },
  {
    name: 'Radial gradient',
    xml: `<a:gradFill rotWithShape="1"><a:gsLst>${stop(0, COLOR.paper)}${stop(100000, COLOR.ink)}</a:gsLst>`
      + '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>',
    lines: ['gradFill, a:path circle', 'fillToRect at the centre', 'white centre, black corners'],
  },
  {
    name: 'Pattern',
    xml: `<a:pattFill prst="dkHorz"><a:fgClr><a:srgbClr val="${COLOR.ink}"/></a:fgClr><a:bgClr><a:srgbClr val="${COLOR.paper}"/></a:bgClr></a:pattFill>`,
    lines: ['pattFill prst dkHorz', `fgClr ${COLOR.ink}, bgClr ${COLOR.paper}`, 'dark horizontal lines'],
  },
];

export default {
  id: 'pptx-009',
  slug: 'gradient-pattern-fills',
  format: 'pptx',
  title: 'Gradient fills',
  covers: [
    'ppt/slides/slide1.xml|p:spPr>a:gradFill',
    'ppt/slides/slide1.xml|a:gradFill>a:gsLst',
    'ppt/slides/slide1.xml|a:gsLst>a:gs',
    'ppt/slides/slide1.xml|a:gs@pos',
    'ppt/slides/slide1.xml|a:gradFill>a:lin',
    'ppt/slides/slide1.xml|a:lin@ang',
    'ppt/slides/slide1.xml|a:gradFill>a:path',
    'ppt/slides/slide1.xml|a:path@path=circle',
    'ppt/slides/slide1.xml|a:path>a:fillToRect',
    'ppt/slides/slide1.xml|p:spPr>a:pattFill',
    'ppt/slides/slide1.xml|a:pattFill@prst=dkHorz',
    'ppt/slides/slide1.xml|a:pattFill>a:fgClr',
    'ppt/slides/slide1.xml|a:pattFill>a:bgClr',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-009',
      title: 'Gradient fills',
      verify: [
        'Verify: three gradients and one pattern fill, each as its caption states, in a field of 3 × 2;',
        'gradients reach their end colours exactly at the edges of the shape.',
      ],
    }),
    ...FILLS.flatMap((f, i) => {
      const c1 = 1 + 3 * i;
      return [
        shape(3 + 2 * i, f.name, field(c1, c1 + 2, 2, 3), 'rect', f.xml, NO_LINE),
        textBox(4 + 2 * i, `Caption ${i + 1}`, field(c1, c1 + 2, 4, 4), f.lines.map((text, j) => ({ style: 'caption', bold: j === 0, text }))),
      ];
    }),
  ]),
};
