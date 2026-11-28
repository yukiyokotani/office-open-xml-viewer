import { DOCX } from '../lib/grid.mjs';
import { A_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header, run } from '../lib/docx.mjs';
import { THEME } from '../lib/pptx.mjs';

// A figure drawn from Word shapes in one group (wpg:wgp), as reports build
// diagrams: rectangles rotated by angles that are not multiples of 90°, filled
// with theme colour accent1 at four luminance settings, and straight arrow
// connectors in all four flip combinations. The rectangles are twice as wide
// as high, so a turn of 136° and one of 45° look different. Each arrowhead is at the end
// point, so the flips decide which corner it points to.
const { emu, colStart, span, LINE } = DOCX;
const NS = [A_NS, R_NS,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'];
const W = emu(span(1, 4));
const H = emu(15 * LINE);
const SIDE = emu(4 * LINE);
const THIN = emu(2 * LINE); // rectangles 2:1, so each angle can be read
const LN = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
const accent = (mods) => `<a:solidFill><a:schemeClr val="accent1">${mods}</a:schemeClr></a:solidFill>`;
const xfrm = (x, y, cx, cy, extra = '') => `<a:xfrm${extra}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const centre = (c) => emu(colStart(c) + span(c, c) / 2);

const SQUARES = [
  [45, ''], [136, '<a:lumMod val="75000"/>'], [224, '<a:lumMod val="50000"/>'], [318, '<a:lumMod val="40000"/><a:lumOff val="60000"/>'],
];
const ARROWS = [['', 'end at bottom right'], [' flipH="1"', 'flipH: bottom left'], [' flipV="1"', 'flipV: top right'], [' flipH="1" flipV="1"', 'both: top left']];

const square = ([deg, mods], i) => `<wps:wsp><wps:cNvPr id="${10 + i}" name="Rectangle ${i + 1}"/><wps:cNvSpPr/><wps:spPr>`
  + xfrm(centre(i + 1) - SIDE / 2, emu(2 * LINE), SIDE, THIN, ` rot="${deg * 60000}"`)
  + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${accent(mods)}${LN}</wps:spPr><wps:bodyPr/></wps:wsp>`;
const arrow = ([flip], i) => `<wps:wsp><wps:cNvPr id="${20 + i}" name="Arrow ${i + 1}"/><wps:cNvCnPr/><wps:spPr>`
  + xfrm(centre(i + 1) - SIDE / 2, emu(9 * LINE), SIDE, SIDE, flip)
  + '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
  + '<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln>'
  + '</wps:spPr><wps:bodyPr/></wps:wsp>';

const GROUP = '<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
  + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
  + `<wp:extent cx="${W}" cy="${H}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>`
  + '<wp:docPr id="1" name="Group 1"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/>'
  + `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></wpg:grpSpPr>`
  + SQUARES.map(square).join('') + ARROWS.map(arrow).join('')
  + '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';

const TABS = [2, 3, 4].map((c) => ['left', colStart(c)]);

export default {
  id: 'docx-027',
  slug: 'shape-group',
  format: 'docx',
  title: 'Shape group',
  covers: [
    'word/document.xml|a:graphicData>wpg:wgp',
    'word/document.xml|wpg:wgp>wps:wsp',
    'word/document.xml|wpg:grpSpPr>a:xfrm',
    'word/document.xml|a:xfrm>a:chOff',
    'word/document.xml|a:xfrm@rot',
    'word/document.xml|a:xfrm@flipH',
    'word/document.xml|a:xfrm@flipV',
    'word/document.xml|wps:wsp>wps:cNvCnPr',
    'word/document.xml|a:ln>a:tailEnd',
    'word/document.xml|a:solidFill>a:schemeClr',
    'word/document.xml|a:schemeClr>a:lumMod',
    'word/document.xml|a:schemeClr>a:lumOff',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-027',
      title: 'Shape group',
      verify: [
        'Verify: four 2:1 rectangles turned 45°, 136°, 224°, 318°',
        '(accent, darker, darkest, pale); below them four arrows',
        'pointing to the corners named. All eight are one group.',
      ],
    }),
    { kind: 'p', before: 1, runs: [GROUP, run('The group sits above this paragraph, wrapped top and bottom.')] },
    { kind: 'p', style: 'caption', tabs: TABS, lines: ['rot 45°, accent1\trot 136°, lumMod 75 %\trot 224°, lumMod 50 %\trot 318°, lumMod 40 % lumOff 60 %'] },
    { kind: 'p', style: 'caption', tabs: TABS, lines: [ARROWS.map(([, label]) => label).join('\t')] },
  ], {
    ns: NS,
    parts: { 'word/theme/theme1.xml': THEME },
    rels: [['rId3', 'theme', 'theme/theme1.xml']],
    overrides: { 'word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml' },
  }),
};
