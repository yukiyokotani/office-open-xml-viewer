import { COLOR, DOCX } from '../lib/grid.mjs';
import { A_NS } from '../lib/drawing.mjs';
import { docxPackage, run } from '../lib/docx.mjs';

// A cover page as Word's cover templates build it: a rectangle anchored to the
// page and as large as the page, behind the text (behindDoc), coloured with
// the accent; over it, white title text in content controls (w:sdt), the way
// templates leave fields for a title and a date. A page break leads to an
// ordinary white page 2.
const { emu, PAGE_W, PAGE_H } = DOCX;
const BACKGROUND = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
  + '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
  + `<wp:extent cx="${emu(PAGE_W)}" cy="${emu(PAGE_H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
  + '<wp:docPr id="1" name="Cover background"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(PAGE_W)}" cy="${emu(PAGE_H)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  + `<a:solidFill><a:srgbClr val="${COLOR.accent}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
const WHITE = `<w:color w:val="${COLOR.paper}"/>`;
const sdt = (alias, tag, content) => `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${tag}"/><w:id w:val="${tag.length}"/><w:text/></w:sdtPr>`
  + `<w:sdtContent>${content}</w:sdtContent></w:sdt>`;

export default {
  id: 'docx-030',
  slug: 'cover-page',
  format: 'docx',
  title: 'Cover page',
  covers: [
    'word/document.xml|wp:anchor@behindDoc',
    'word/document.xml|wp:positionH@relativeFrom=page',
    'word/document.xml|w:p>w:sdt',
    'word/document.xml|w:sdtPr>w:alias',
    'word/document.xml|w:sdtPr>w:text',
  ],
  parts: () => docxPackage([
    { kind: 'p', style: 'caption', bold: true, runs: [BACKGROUND, run('docx-030', `<w:b/>${WHITE}<w:sz w:val="16"/>`)], height: 1 },
    { kind: 'p', at: DOCX.rowLine(3), style: 'title', runs: [sdt('Title', 'title', run('Cover page', `<w:b/>${WHITE}<w:sz w:val="48"/>`))], height: 2 },
    { kind: 'p', before: 1, runs: [sdt('Date', 'date', run('2026', WHITE))], height: 1 },
    {
      kind: 'p',
      before: 1,
      cols: [1, 2],
      runs: [run('Verify: page 1 is cyan to every edge, the rectangle behind the text; the title and the year, held in content '
        + 'controls, are white; page 2 is white.', `${WHITE}<w:sz w:val="16"/>`)],
    },
    { kind: 'p', pageBreakBefore: true, text: 'Page 2: the cover rectangle is anchored to page 1 only, so this page has no background.' },
  ], {
    ns: [A_NS,
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'],
  }),
};
