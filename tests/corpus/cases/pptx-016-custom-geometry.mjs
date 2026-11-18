import { COLOR, PPTX } from '../lib/grid.mjs';
import { NO_LINE, field, header, line, pptxPackage, solid, textBox, xfrm } from '../lib/pptx.mjs';

// Path coordinates are in the path's own space (w × h); the shape stretches
// it over its field. Each path below is drawn in a 12 × 18 space: one unit is
// a quarter column wide and a sixth of two rows high, so vertices fall on the
// grid lines inside the field.
const custom = (id, name, box, path, fill, ln) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xfrm(box)}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>`
  + `<a:pathLst>${path}</a:pathLst></a:custGeom>${fill}${ln}</p:spPr></p:sp>`;
const pt = (x, y) => `<a:pt x="${x}" y="${y}"/>`;
const moveTo = (x, y) => `<a:moveTo>${pt(x, y)}</a:moveTo>`;
const lnTo = (x, y) => `<a:lnTo>${pt(x, y)}</a:lnTo>`;

const STEPS = `<a:path w="12" h="18">${moveTo(0, 18)}${lnTo(0, 12)}${lnTo(4, 12)}${lnTo(4, 6)}${lnTo(8, 6)}${lnTo(8, 0)}${lnTo(12, 0)}${lnTo(12, 18)}<a:close/></a:path>`;
// Up the left edge to the top left corner, which is the 270° point of an
// ellipse centred on the bottom left corner; sweeping 90° clockwise ends the
// arc at the bottom right corner (12, 18). A quarter ellipse.
const ARC = `<a:path w="12" h="18">${moveTo(0, 18)}${lnTo(0, 0)}<a:arcTo wR="12" hR="18" stAng="16200000" swAng="5400000"/><a:close/></a:path>`;
const WAVE = `<a:path w="12" h="18" fill="none">${moveTo(0, 9)}<a:cubicBezTo>${pt(4, 0)}${pt(8, 18)}${pt(12, 9)}</a:cubicBezTo></a:path>`;

const SHAPES = [
  ['Steps', STEPS, solid(COLOR.ink), NO_LINE, ['custGeom: moveTo, lnTo, close', 'three steps of a third each', 'solidFill 000000']],
  ['Quarter', ARC, solid(COLOR.accent), NO_LINE, ['custGeom: arcTo', 'wR 12, hR 18, stAng 270°, swAng 90°', `solidFill ${COLOR.accent}`]],
  ['Wave', WAVE, '<a:noFill/>', line(25400), ['custGeom: cubicBezTo, fill none', 'from mid-left to mid-right', 'ln 2 pt 000000']],
];

export default {
  id: 'pptx-016',
  slug: 'custom-geometry',
  format: 'pptx',
  title: 'Custom geometry',
  covers: [
    'ppt/slides/slide1.xml|p:spPr>a:custGeom',
    'ppt/slides/slide1.xml|a:custGeom>a:pathLst',
    'ppt/slides/slide1.xml|a:pathLst>a:path',
    'ppt/slides/slide1.xml|a:path@w',
    'ppt/slides/slide1.xml|a:path>a:moveTo',
    'ppt/slides/slide1.xml|a:path>a:lnTo',
    'ppt/slides/slide1.xml|a:path>a:arcTo',
    'ppt/slides/slide1.xml|a:path>a:cubicBezTo',
    'ppt/slides/slide1.xml|a:path>a:close',
    'ppt/slides/slide1.xml|a:path@fill=none',
    'ppt/slides/slide1.xml|a:custGeom>a:rect',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-016',
      title: 'Custom geometry',
      verify: [
        'Verify: black steps, an orange quarter ellipse and a wave, each filling a field of three columns',
        'by two rows; corners and ends meet the field edges exactly.',
      ],
    }),
    ...SHAPES.flatMap(([name, path, fill, ln, lines], i) => {
      const c1 = 1 + 3 * i;
      return [
        custom(3 + 2 * i, name, field(c1, c1 + 2, 2, 3), path, fill, ln),
        textBox(4 + 2 * i, `Caption ${i + 1}`, field(c1, c1 + 2, 4, 4), lines.map((text, j) => ({ style: 'caption', bold: j === 0, text }))),
      ];
    }),
  ]),
};
