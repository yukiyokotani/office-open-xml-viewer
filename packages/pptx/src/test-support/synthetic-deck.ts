import { storedZip } from '@silurus/ooxml-core/testing';

// Synthetic one-slide PPTX decks for tests (no third-party content): a
// minimal PresentationML package (ECMA-376 Part 1 §19.2 presentation, §19.3
// slide/layout/master, §20.1.6 theme) around caller-supplied slide shapes.

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A one-slide deck whose slide tree is `shapes` (p:sp / p:graphicFrame XML),
 * with `media` parts (`ppt/media/<name>`) related to the slide as `rIdM<n>`. */
export function deck(shapes: string, media: ReadonlyMap<string, Uint8Array> = new Map()): Uint8Array {
  const xml = (body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
  const files: Array<[string, string | Uint8Array]> = [
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="emf" ContentType="image/x-emf"/>'
      + '<Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
      + '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
      + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
      + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
      + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
      + '</Types>')],
    ['_rels/.rels', xml(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`)],
    ['ppt/presentation.xml', xml(`<p:presentation xmlns:p="${P}" xmlns:r="${R}">`
      + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
      + '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
      + '<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')],
    ['ppt/_rels/presentation.xml.rels', xml(`<Relationships xmlns="${REL}">`
      + `<Relationship Id="rId1" Type="${OFFICE_REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`
      + `<Relationship Id="rId2" Type="${OFFICE_REL}/slide" Target="slides/slide1.xml"/>`
      + `<Relationship Id="rId3" Type="${OFFICE_REL}/theme" Target="theme/theme1.xml"/></Relationships>`)],
    ['ppt/slides/slide1.xml', xml(`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>`
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
      + `${shapes}</p:spTree></p:cSld></p:sld>`)],
    ['ppt/slides/_rels/slide1.xml.rels', xml(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`
      + [...media.keys()].map((name, index) => `<Relationship Id="rIdM${index}" Type="${OFFICE_REL}/image" Target="../media/${name}"/>`).join('')
      + '</Relationships>')],
    ['ppt/slideLayouts/slideLayout1.xml', xml(`<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>`
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>')],
    ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', xml(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`)],
    ['ppt/slideMasters/slideMaster1.xml', xml(`<p:sldMaster xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>`
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>'
      + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
      + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', xml(`<Relationships xmlns="${REL}">`
      + `<Relationship Id="rId1" Type="${OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`
      + `<Relationship Id="rId2" Type="${OFFICE_REL}/theme" Target="../theme/theme1.xml"/></Relationships>`)],
    ['ppt/theme/theme1.xml', xml(`<a:theme xmlns:a="${A}" name="T"><a:themeElements>`
      + '<a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
      + '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
      + '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
      + '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
      + '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
      + '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
      + '<a:fontScheme name="F"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
      + '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
      + '<a:fmtScheme name="M"><a:fillStyleLst><a:noFill/><a:noFill/><a:noFill/></a:fillStyleLst>'
      + '<a:lnStyleLst><a:ln/><a:ln/><a:ln/></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
      + '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
      + '<a:bgFillStyleLst><a:noFill/><a:noFill/><a:noFill/></a:bgFillStyleLst></a:fmtScheme>'
      + '</a:themeElements></a:theme>')],
  ];
  return storedZip([...files, ...[...media].map(([name, bytes]) => [`ppt/media/${name}`, bytes] as [string, Uint8Array])]);
}

let shapeId = 2;
/** One text box at a fixed frame with the given `a:bodyPr` and paragraphs. */
export const textBox = (paragraphs: string, bodyPr = '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId++}" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + '<p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="3200400" cy="2743200"/></a:xfrm>'
  + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${bodyPr}<a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
export const para = (runs: string, pPr = '<a:pPr><a:buNone/></a:pPr>') => `<a:p>${pPr}${runs}</a:p>`;
export const run = (text: string, rPr = '') => `<a:r><a:rPr lang="en-US" sz="2400"${rPr}/><a:t>${text}</a:t></a:r>`;


/** A picture of the `index`th `deck` media part (§19.3.1.37 pic). */
export const picture = (index: number) => `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId++}" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="rIdM${index}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
  + '<p:spPr><a:xfrm><a:off x="457200" y="3657600"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
