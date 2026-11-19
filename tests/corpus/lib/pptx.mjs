// PresentationML specimen composer: one 16:9 slide on the 12-column grid of
// grid.mjs PPTX, with the minimum master, layout and theme PowerPoint needs.

import { COLOR, FONT, LANG, PPTX } from './grid.mjs';
import { contentTypes, relationships } from './opc.mjs';
import { XML_DECL, attrs, esc } from './xml.mjs';

export const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
export const EMPTY_TREE = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';

export const solid = (rgb) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
export const line = (w, rgb = COLOR.ink) => `<a:ln w="${w}">${solid(rgb)}</a:ln>`;
export const NO_LINE = '<a:ln><a:noFill/></a:ln>';
export const xfrm = ({ x, y, cx, cy }) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;

/** One baseline (12 pt) in hundredths of a point, the unit of a:spcPts. */
const BASELINE_PTS = PPTX.LINE / 127;

/** Grid field: columns c1..c2, rows r1..r2 → offset and extent in EMU. */
export const field = (c1, c2, r1, r2) => ({ x: PPTX.x(c1), y: PPTX.y(r1), cx: PPTX.w(c1, c2), cy: PPTX.h(r1, r2) });

/** Paragraph: { style?: caption|text|title, bold?, text, lvl?, before? (lines),
 * algn?, own?: false to inherit line spacing and size from a list style,
 * themeFonts?: name the theme minor font explicitly }. */
export function paragraph(p) {
  if (p.xml) return p.xml;
  const type = PPTX.TYPE[p.style ?? 'text'];
  const own = p.own !== false;
  const inner = (own ? `<a:lnSpc><a:spcPts val="${type.leading}"/></a:lnSpc>` : '')
    + (p.before ? `<a:spcBef><a:spcPts val="${p.before * BASELINE_PTS}"/></a:spcBef>` : '');
  const pAttrs = attrs({ lvl: p.lvl || undefined, algn: p.algn });
  const pPr = inner ? `<a:pPr${pAttrs}>${inner}</a:pPr>` : pAttrs ? `<a:pPr${pAttrs}/>` : '';
  const rAttrs = attrs({ lang: p.lang ?? LANG.latin, altLang: p.altLang, sz: own ? type.sz : undefined, b: p.bold ? 1 : undefined });
  const rPr = p.themeFonts
    ? `<a:rPr${rAttrs}><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr>`
    : `<a:rPr${rAttrs}/>`;
  return `<a:p>${pPr}<a:r>${rPr}<a:t>${esc(p.text)}</a:t></a:r></a:p>`;
}

/** Text box with zero insets, no autofit, top anchored. body: extra a:bodyPr
 * attributes (numCol, spcCol, vert, ...). */
export function textBox(id, name, box, paragraphs, lstStyle = '', body = {}) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
    + `<p:txBody>${bodyPr(body)}<a:lstStyle>${lstStyle}</a:lstStyle>${paragraphs.map(paragraph).join('')}</p:txBody></p:sp>`;
}

/** a:bodyPr with the grid's zero insets; extra attributes go after them. */
export const bodyPr = (body = {}) => `<a:bodyPr${attrs({ ...(body.vert ? { vert: body.vert } : {}), wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0, numCol: body.numCol, spcCol: body.spcCol, anchor: 't' })}><a:noAutofit/></a:bodyPr>`;

/** Preset-geometry shape without text. */
export function shape(id, name, box, prst, fill, ln) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrm(box)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}${ln}</p:spPr></p:sp>`;
}

/** Table frame. columns: widths; rows: [{ h, cells: [{ text, style?, bold?,
 * algn?, anchor?, marR?, borders?: { L|R|T|B: width } }] }]. */
export function table(id, name, box, columns, rows) {
  const trs = rows.map((row) => `<a:tr h="${row.h}">${row.cells.map((cell) => {
    // Every edge is written: a table without a table style has no defined
    // borders, and consumers differ on what they draw for an unspecified edge.
    // CT_TableCellProperties child order: lnL, lnR, lnT, lnB.
    const edges = ['L', 'R', 'T', 'B'].map((e) => {
      const w = cell.borders?.[e];
      return w ? `<a:ln${e} w="${w}">${solid(COLOR.ink)}</a:ln${e}>` : `<a:ln${e}><a:noFill/></a:ln${e}>`;
    });
    const tcPr = `<a:tcPr${attrs({ marL: 0, marR: cell.marR ?? PPTX.GUTTER, marT: 0, marB: 0, anchor: cell.anchor })}>${edges.join('')}</a:tcPr>`;
    // Table text names the theme font itself: without a table style, table text
    // has no inherited font, and consumers differ (LibreOffice falls back to
    // its own default face).
    return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph({ ...cell, themeFonts: true })}</a:txBody>${tcPr}</a:tc>`;
  }).join('')}</a:tr>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/>`
    + '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
    + `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></p:xfrm>`
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
    + `<a:tbl><a:tblPr/><a:tblGrid>${columns.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${trs}</a:tbl>`
    + '</a:graphicData></a:graphic></p:graphicFrame>';
}

/** Header furniture in the top-left field: columns 1-6, row 1. */
export function header({ id, title, verify }) {
  return textBox(2, 'Header', field(1, 12, 1, 1), [
    { style: 'caption', bold: true, text: id },
    { style: 'title', bold: true, text: title },
    ...verify.map((text, i) => ({ style: 'caption', text, before: i === 0 ? 1 : 0 })),
  ]);
}

const LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9];
const DEFAULT_TEXT_STYLE = '<p:defaultTextStyle>'
  + `<a:defPPr><a:defRPr lang="${LANG.latin}"/></a:defPPr>`
  + LEVELS.map((n) => `<a:lvl${n}pPr marL="0" indent="0" algn="l"><a:defRPr sz="${PPTX.TYPE.text.sz}">`
    + '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
    + '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl' + n + 'pPr>').join('')
  + '</p:defaultTextStyle>';

const fontSet = `<a:latin typeface="${FONT.latin}"/><a:ea typeface="${FONT.eastAsia}"/><a:cs typeface=""/>`;
const phFill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
export const THEME = XML_DECL + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Corpus"><a:themeElements>'
  + '<a:clrScheme name="Corpus">'
  + `<a:dk1><a:srgbClr val="${COLOR.ink}"/></a:dk1><a:lt1><a:srgbClr val="${COLOR.paper}"/></a:lt1>`
  + `<a:dk2><a:srgbClr val="${COLOR.ink}"/></a:dk2><a:lt2><a:srgbClr val="${COLOR.paper}"/></a:lt2>`
  + `<a:accent1><a:srgbClr val="${COLOR.accent}"/></a:accent1>`
  + [2, 3, 4, 5, 6].map((n) => `<a:accent${n}><a:srgbClr val="${COLOR.ink}"/></a:accent${n}>`).join('')
  + `<a:hlink><a:srgbClr val="${COLOR.ink}"/></a:hlink><a:folHlink><a:srgbClr val="${COLOR.ink}"/></a:folHlink>`
  + '</a:clrScheme>'
  + `<a:fontScheme name="Corpus"><a:majorFont>${fontSet}</a:majorFont><a:minorFont>${fontSet}</a:minorFont></a:fontScheme>`
  + '<a:fmtScheme name="Corpus">'
  + `<a:fillStyleLst>${phFill.repeat(3)}</a:fillStyleLst>`
  + `<a:lnStyleLst>${[6350, 12700, 19050].map((w) => `<a:ln w="${w}">${phFill}</a:ln>`).join('')}</a:lnStyleLst>`
  + `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst>`
  + `<a:bgFillStyleLst>${phFill.repeat(3)}</a:bgFillStyleLst>`
  + '</a:fmtScheme></a:themeElements></a:theme>';

const masterXml = ({ shapes = '', txStyles = '' } = {}) => XML_DECL + `<p:sldMaster ${NS}><p:cSld>`
  + '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
  + `${EMPTY_TREE}${shapes}</p:spTree></p:cSld>`
  + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
  + 'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
  + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>${txStyles}</p:sldMaster>`;

const layoutXml = ({ shapes = '', type = 'blank', name = 'Blank' } = {}) => XML_DECL
  + `<p:sldLayout ${NS} type="${type}" preserve="1"><p:cSld name="${name}">${EMPTY_TREE}${shapes}</p:spTree></p:cSld>`
  + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';

const PRESENTATION = XML_DECL + `<p:presentation ${NS}>`
  + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
  + '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
  + `<p:sldSz cx="${PPTX.SLIDE_W}" cy="${PPTX.SLIDE_H}"/><p:notesSz cx="6858000" cy="9144000"/>`
  + `${DEFAULT_TEXT_STYLE}</p:presentation>`;

const PML = 'application/vnd.openxmlformats-officedocument.presentationml';

/** extra: { parts?: { path: xml|bytes }, rels?: [[id, type, target]] from the
 * slide (rId1 is the layout), overrides?: { part: contentType }, defaults?:
 * { ext: contentType }, master?: { shapes, txStyles }, layout?: { shapes,
 * type, name }, slideBg?: '<p:bg>…', clrMapOvr?: '<a:overrideClrMapping …/>' } */
export function pptxPackage(shapes, extra = {}) {
  const slide = XML_DECL + `<p:sld ${NS}><p:cSld>${extra.slideBg ?? ''}${EMPTY_TREE}${shapes.join('')}</p:spTree></p:cSld>`
    + `<p:clrMapOvr>${extra.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr></p:sld>`;
  const types = contentTypes({
    'ppt/presentation.xml': `${PML}.presentation.main+xml`,
    'ppt/slideMasters/slideMaster1.xml': `${PML}.slideMaster+xml`,
    'ppt/slideLayouts/slideLayout1.xml': `${PML}.slideLayout+xml`,
    'ppt/slides/slide1.xml': `${PML}.slide+xml`,
    'ppt/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
    ...(extra.overrides ?? {}),
  });
  const defaults = Object.entries(extra.defaults ?? {})
    .map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('');
  return {
    '[Content_Types].xml': defaults ? types.replace('<Default Extension="rels"', `${defaults}<Default Extension="rels"`) : types,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/_rels/presentation.xml.rels': relationships([
      ['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'],
      ['rId2', 'slide', 'slides/slide1.xml'],
      ['rId3', 'theme', 'theme/theme1.xml'],
    ]),
    'ppt/presentation.xml': PRESENTATION,
    'ppt/slideMasters/slideMaster1.xml': masterXml(extra.master),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': relationships([
      ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
      ['rId2', 'theme', '../theme/theme1.xml'],
    ]),
    'ppt/slideLayouts/slideLayout1.xml': layoutXml(extra.layout),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relationships([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/slides/slide1.xml': slide,
    'ppt/slides/_rels/slide1.xml.rels': relationships([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ...(extra.rels ?? [])]),
    'ppt/theme/theme1.xml': THEME,
    ...(extra.parts ?? {}),
  };
}
