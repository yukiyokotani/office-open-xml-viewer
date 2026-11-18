import { DOCX } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS, pictureGraphic } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header, run } from '../lib/docx.mjs';
import { gridPicture } from '../lib/png.mjs';

// Floating pictures as reports and forms use them: anchored to the column
// and paragraph rather than the margin, side by side, with each wrap type,
// pulled into the margin by negative offsets, and reaching into the bottom
// margin. Every picture is the same PNG: one outlined field.
const { emu, span, colStart, LINE } = DOCX;
const PX = 20;
const png = gridPicture({ cols: 1, rows: 1, field: Math.round(span(1, 1) / PX), gutterX: 0, fieldH: 4 * LINE / PX, gutterY: 0 });
const SQUARE = { cx: emu(span(1, 1)), cy: emu(4 * LINE) };
const WIDE = { cx: emu(span(1, 2)), cy: emu(4 * LINE) };
const POLYGON = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/>'
  + '<wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
const WRAP = {
  topAndBottom: '<wp:wrapTopAndBottom/>',
  square: '<wp:wrapSquare wrapText="bothSides"/>',
  tight: `<wp:wrapTight wrapText="bothSides">${POLYGON}</wp:wrapTight>`,
  through: `<wp:wrapThrough wrapText="bothSides">${POLYGON}</wp:wrapThrough>`,
  none: '<wp:wrapNone/>',
};
let nextId = 1;
function anchor({ x, y, size, h = 'column', v = 'paragraph', wrap, behind = false, dist = emu(DOCX.GUTTER) }) {
  const id = nextId++;
  return '<w:r><w:drawing>'
    + `<wp:anchor distT="0" distB="0" distL="${dist}" distR="${dist}" simplePos="0" relativeHeight="${id}" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="${h}"><wp:posOffset>${x}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="${v}"><wp:posOffset>${y}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${size.cx}" cy="${size.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${WRAP[wrap]}`
    + `<wp:docPr id="${id}" name="Picture ${id}"/>${pictureGraphic(id, `Picture ${id}`, 'rId3', size.cx, size.cy)}`
    + '</wp:anchor></w:drawing></w:r>';
}
const caption = (text) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1 });
const FILL = 'The text of this paragraph flows round the picture according to its wrap type. It is long enough to run past the '
  + 'picture, so the lines beside it and the lines below it can both be seen: beside the picture they stop one gutter '
  + 'from it (distL and distR), below it they take the full measure again. The picture is four baselines high and one '
  + 'grid column wide; the paragraph runs to six or more lines, so the next caption always starts below the picture.';

export default {
  id: 'docx-026',
  slug: 'floating-pictures',
  format: 'docx',
  title: 'Floating pictures',
  covers: [
    'word/document.xml|wp:positionH@relativeFrom=column',
    'word/document.xml|wp:positionV@relativeFrom=paragraph',
    'word/document.xml|wp:positionV@relativeFrom=page',
    'word/document.xml|wp:anchor>wp:wrapTopAndBottom',
    'word/document.xml|wp:anchor>wp:wrapTight',
    'word/document.xml|wp:anchor>wp:wrapThrough',
    'word/document.xml|wp:anchor>wp:wrapNone',
    'word/document.xml|wp:wrapTight>wp:wrapPolygon',
  ],
  parts: () => build(),
};

function build() {
  nextId = 1;
  return docxPackage([
    ...header({
      id: 'docx-026',
      title: 'Floating pictures',
      verify: [
        'Verify: A sets two pictures side by side above its text; B-D',
        'wrap text square, tight and over the picture; E pulls a',
        'picture into the left margin; F ends in the bottom margin.',
      ],
    }),
    caption('A. Side by side: two pictures anchored to one paragraph, top and bottom wrapping'),
    { kind: 'p', runs: [anchor({ x: 0, y: 0, size: WIDE, wrap: 'topAndBottom' }), anchor({ x: emu(colStart(3)), y: 0, size: WIDE, wrap: 'topAndBottom' }),
      run('Both pictures are anchored to this paragraph at its top; the text continues below them.')] },
    caption('B. Square wrapping, picture in column 4'),
    { kind: 'p', runs: [anchor({ x: emu(colStart(4)), y: 0, size: SQUARE, wrap: 'square' }), run(FILL)] },
    caption('C. Tight wrapping (rectangular polygon), picture in column 1'),
    { kind: 'p', runs: [anchor({ x: 0, y: 0, size: SQUARE, wrap: 'tight' }), run(FILL)] },
    caption('D. Through wrapping, and a picture behind the text (wrapNone, behindDoc)'),
    { kind: 'p', runs: [anchor({ x: emu(colStart(2)), y: 0, size: SQUARE, wrap: 'through' }), anchor({ x: emu(colStart(4)), y: 0, size: SQUARE, wrap: 'none', behind: true }), run(FILL)] },
    caption('E. Negative offsets: the picture starts one gutter left of the column and one line above its paragraph'),
    { kind: 'p', runs: [anchor({ x: -emu(DOCX.GUTTER) - SQUARE.cx, y: -emu(LINE), size: SQUARE, wrap: 'none' }), run('The picture sits in the left margin, beside this paragraph; it does not move the text.')] },
    caption('F. Into the bottom margin: anchored to the page, two baselines above the foot of the live area'),
    { kind: 'p', runs: [anchor({ x: emu(colStart(3)), y: emu(DOCX.MARGIN.top + (DOCX.LINES - 2) * LINE), size: SQUARE, v: 'page', wrap: 'none' }),
      run('The picture in column 3 begins two baselines above the bottom of the live area and ends in the bottom margin; it is drawn whole, not cut at the margin.')] },
  ], {
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
    parts: { 'word/media/image1.png': png.bytes },
    rels: [['rId3', 'image', 'media/image1.png']],
    defaults: { png: 'image/png' },
  });
}
