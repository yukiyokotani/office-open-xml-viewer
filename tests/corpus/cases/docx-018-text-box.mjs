import { DOCX } from '../lib/grid.mjs';
import { A_NS } from '../lib/drawing.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const { emu, colStart, span, lineTop, LINE } = DOCX;
const X = colStart(3);
const Y = lineTop(at(2));
const W = span(3, 4);
const H = 9 * LINE;
const pt = (twips) => twips / 20;

const NS = [
  A_NS,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
];

// Text box content: a 1 pt top rule taking one baseline marks the box top.
const RULE = '<w:pBdr><w:top w:val="single" w:sz="8" w:space="11" w:color="000000"/></w:pBdr>';
const CONTENT = '<w:txbxContent>'
  + `<w:p><w:pPr>${RULE}</w:pPr>${run('Text box: wps:wsp inside mc:AlternateContent.')}</w:p>`
  + `<w:p>${run('Its frame fills grid columns 3-4 of row 2;')}</w:p>`
  + `<w:p>${run('wps:bodyPr insets are 0, so this text starts')}</w:p>`
  + `<w:p>${run('on the column edge, one baseline below the rule.')}</w:p>`
  + `<w:p>${run('The mc:Fallback carries the same content as a')}</w:p>`
  + `<w:p>${run('VML v:shape with a v:textbox for older readers.')}</w:p>`
  + '</w:txbxContent>';

const TEXT_BOX = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${emu(X)}</wp:posOffset></wp:positionH>`
  + `<wp:positionV relativeFrom="margin"><wp:posOffset>${emu(Y)}</wp:posOffset></wp:positionV>`
  + `<wp:extent cx="${emu(W)}" cy="${emu(H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
  + '<wp:docPr id="1" name="Text box 1"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(W)}" cy="${emu(H)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>`
  + `<wps:txbx>${CONTENT}</wps:txbx>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
  + '<mc:Fallback><w:pict>'
  + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
  + `<v:shape id="Text Box 1" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:${pt(X)}pt;margin-top:${pt(Y)}pt;width:${pt(W)}pt;height:${pt(H)}pt;z-index:1;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin" filled="f" stroked="f">`
  + `<v:textbox inset="0,0,0,0">${CONTENT}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

export default {
  id: 'docx-018',
  slug: 'text-box',
  format: 'docx',
  title: 'Text box',
  covers: [
    'word/document.xml|w:r>mc:AlternateContent',
    'word/document.xml|mc:AlternateContent>mc:Choice',
    'word/document.xml|mc:AlternateContent>mc:Fallback',
    'word/document.xml|a:graphicData>wps:wsp',
    'word/document.xml|wps:wsp>wps:txbx',
    'word/document.xml|wps:txbx>w:txbxContent',
    'word/document.xml|wps:wsp>wps:bodyPr',
    'word/document.xml|wp:anchor>wp:wrapNone',
    'word/document.xml|w:pict>v:shapetype',
    'word/document.xml|w:pict>v:shape',
    'word/document.xml|v:shape>v:textbox',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-018',
      title: 'Text box',
      verify: [
        'Verify: the text box rule sits on the top of row 2 over',
        'columns 3-4, its text one baseline below, beside this text.',
      ],
    }),
    {
      kind: 'p', at: at(2), cols: [1, 2], height: 4,
      runs: [TEXT_BOX, run('This paragraph anchors the text box. The box'), '<w:r><w:br/></w:r>',
        run('does not wrap the text (wp:wrapNone): the text'), '<w:r><w:br/></w:r>',
        run('keeps to columns 1-2 by its own right indent,'), '<w:r><w:br/></w:r>',
        run('and the box stands beside it in columns 3-4.')],
    },
  ], { ns: NS.map((n) => n) }),
};
