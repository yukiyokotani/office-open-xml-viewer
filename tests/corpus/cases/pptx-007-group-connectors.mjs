import { COLOR, PPTX } from '../lib/grid.mjs';
import { NO_LINE, field, header, line, pptxPackage, shape, solid, textBox, xfrm } from '../lib/pptx.mjs';

const LINE_W = 12700;
const A = field(1, 2, 2, 2);
const B = field(5, 6, 2, 2);
const groupBox = { x: A.x, y: A.y, cx: B.x + B.cx - A.x, cy: A.cy };
// Child coordinates are twice the slide coordinates: the group scales its
// children by chExt → ext, here one half, and moves them by chOff → off.
const toChild = (b) => ({ x: 2 * b.x, y: 2 * b.y, cx: 2 * b.cx, cy: 2 * b.cy });
const groupXfrm = `<a:xfrm><a:off x="${groupBox.x}" y="${groupBox.y}"/><a:ext cx="${groupBox.cx}" cy="${groupBox.cy}"/>`
  + `<a:chOff x="${2 * groupBox.x}" y="${2 * groupBox.y}"/><a:chExt cx="${2 * groupBox.cx}" cy="${2 * groupBox.cy}"/></a:xfrm>`;

const C = field(7, 8, 2, 2);
const D = field(11, 12, 3, 3);

/** Connector from shape `st` (site stIdx) to `end` (site endIdx). */
const connector = (id, name, prst, box, st, end, flipV) => `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}"/>`
  + `<p:cNvCxnSpPr><a:stCxn id="${st[0]}" idx="${st[1]}"/><a:endCxn id="${end[0]}" idx="${end[1]}"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>`
  + `<p:spPr>${xfrm(box).replace('<a:xfrm>', `<a:xfrm${flipV ? ' flipV="1"' : ''}>`)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`
  + `<a:ln w="${LINE_W}">${solid(COLOR.ink)}<a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>`;

// Connection sites of a rectangle: 0 top, 1 left, 2 bottom, 3 right.
const midY = (b) => b.y + b.cy / 2;

export default {
  id: 'pptx-007',
  slug: 'group-connectors',
  format: 'pptx',
  title: 'Groups and connectors',
  covers: [
    'ppt/slides/slide1.xml|p:spTree>p:grpSp',
    'ppt/slides/slide1.xml|p:grpSpPr>a:xfrm',
    'ppt/slides/slide1.xml|a:xfrm>a:chOff',
    'ppt/slides/slide1.xml|a:xfrm>a:chExt',
    'ppt/slides/slide1.xml|p:spTree>p:cxnSp',
    'ppt/slides/slide1.xml|p:cNvCxnSpPr>a:stCxn',
    'ppt/slides/slide1.xml|p:cNvCxnSpPr>a:endCxn',
    'ppt/slides/slide1.xml|p:cxnSp>p:spPr',
    'ppt/slides/slide1.xml|a:ln>a:tailEnd',
    'ppt/slides/slide1.xml|a:tailEnd@type=triangle',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-007',
      title: 'Groups and connectors',
      verify: [
        'Verify: the grouped rectangles fill columns 1-2 and 5-6 of row 2 at half their child size;',
        'arrows run from edge to edge: straight in columns 3-4, elbowed from column 8 to 11.',
      ],
    }),
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
      + `<p:grpSpPr>${groupXfrm}</p:grpSpPr>`
      + shape(4, 'A', toChild(A), 'rect', solid(COLOR.ink), NO_LINE)
      + shape(5, 'B', toChild(B), 'rect', '<a:noFill/>', line(LINE_W))
      + '</p:grpSp>',
    connector(6, 'Straight', 'straightConnector1', { x: A.x + A.cx, y: midY(A), cx: B.x - A.x - A.cx, cy: 0 }, [4, 3], [5, 1]),
    shape(7, 'C', C, 'rect', solid(COLOR.ink), NO_LINE),
    shape(8, 'D', D, 'rect', '<a:noFill/>', line(LINE_W)),
    connector(9, 'Elbow', 'bentConnector3', { x: C.x + C.cx, y: midY(C), cx: D.x - C.x - C.cx, cy: midY(D) - midY(C) }, [7, 3], [8, 1]),
    textBox(10, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', text: `Group: off ${groupBox.x}, ${groupBox.y}; ext ${groupBox.cx} × ${groupBox.cy}; chOff and chExt twice as large.` },
      { style: 'caption', text: 'Connectors name their end shapes and sites (a:stCxn, a:endCxn): site 3 is the right edge, site 1 the left.' },
    ]),
  ]),
};
