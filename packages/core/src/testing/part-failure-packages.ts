import { storedZip } from './non-ooxml-fixture';

// Synthetic OOXML packages with one malformed part each (#1974), authored here
// for tests — no third-party content. Browser-safe: no Node imports.

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
/** A mismatched end tag: not well-formed XML (XML 1.0 §3, element type match). */
const MALFORMED = '<broken><a></b></broken>';

function contentTypes(overrides: ReadonlyArray<readonly [string, string]>): string {
  return `${XML}<Types xmlns="${CT}">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')
    + '</Types>';
}

function rels(targets: ReadonlyArray<readonly [string, string]>): string {
  return `${XML}<Relationships xmlns="${REL}">`
    + targets.map(([type, target], index) => (
      `<Relationship Id="rId${index + 1}" Type="${OFFICE_REL}/${type}" Target="${target}"/>`
    )).join('')
    + '</Relationships>';
}

/** A WordprocessingML package whose `word/document.xml` is malformed. */
export function brokenDocx(): Uint8Array<ArrayBuffer> {
  return storedZip([
    ['[Content_Types].xml', contentTypes([[
      '/word/document.xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    ]])],
    ['_rels/.rels', rels([['officeDocument', 'word/document.xml']])],
    ['word/document.xml', `${XML}${MALFORMED}`],
  ]);
}

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = OFFICE_REL;
const SLIDE_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SP_TREE_ROOT = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';

/** A two-slide PresentationML package whose second slide part is malformed. */
export function brokenPptx(): Uint8Array<ArrayBuffer> {
  const slide = `${XML}<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>${SP_TREE_ROOT}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="3200400" cy="914400"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></p:spPr></p:sp>'
    + '</p:spTree></p:cSld></p:sld>';
  const slideRels = rels([['slideLayout', '../slideLayouts/slideLayout1.xml']]);
  return storedZip([
    ['[Content_Types].xml', contentTypes([
      ['/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
      ['/ppt/slides/slide1.xml', SLIDE_TYPE],
      ['/ppt/slides/slide2.xml', SLIDE_TYPE],
      ['/ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
      ['/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
      ['/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ])],
    ['_rels/.rels', rels([['officeDocument', 'ppt/presentation.xml']])],
    ['ppt/presentation.xml', `${XML}<p:presentation xmlns:p="${P}" xmlns:r="${R}">`
      + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
      + '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>'
      + '<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'],
    ['ppt/_rels/presentation.xml.rels', rels([
      ['slideMaster', 'slideMasters/slideMaster1.xml'],
      ['slide', 'slides/slide1.xml'],
      ['slide', 'slides/slide2.xml'],
      ['theme', 'theme/theme1.xml'],
    ])],
    ['ppt/slides/slide1.xml', slide],
    ['ppt/slides/_rels/slide1.xml.rels', slideRels],
    ['ppt/slides/slide2.xml', `${XML}${MALFORMED}`],
    ['ppt/slides/_rels/slide2.xml.rels', slideRels],
    ['ppt/slideLayouts/slideLayout1.xml', `${XML}<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${SP_TREE_ROOT}</p:spTree></p:cSld></p:sldLayout>`],
    ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([['slideMaster', '../slideMasters/slideMaster1.xml']])],
    ['ppt/slideMasters/slideMaster1.xml', `${XML}<p:sldMaster xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>${SP_TREE_ROOT}</p:spTree></p:cSld>`
      + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
      + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([
      ['slideLayout', '../slideLayouts/slideLayout1.xml'],
      ['theme', '../theme/theme1.xml'],
    ])],
    ['ppt/theme/theme1.xml', `${XML}<a:theme xmlns:a="${A}" name="T"><a:themeElements>`
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
      + '</a:themeElements></a:theme>'],
  ]);
}

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

/** A two-sheet SpreadsheetML package whose second worksheet part is malformed. */
export function brokenXlsx(): Uint8Array<ArrayBuffer> {
  return storedZip([
    ['[Content_Types].xml', contentTypes([
      ['/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
      ['/xl/worksheets/sheet1.xml', SHEET_TYPE],
      ['/xl/worksheets/sheet2.xml', SHEET_TYPE],
      ['/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    ])],
    ['_rels/.rels', rels([['officeDocument', 'xl/workbook.xml']])],
    ['xl/workbook.xml', `${XML}<workbook xmlns="${S}" xmlns:r="${R}"><sheets>`
      + '<sheet name="Healthy" sheetId="1" r:id="rId1"/><sheet name="Broken" sheetId="2" r:id="rId2"/>'
      + '</sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', rels([
      ['worksheet', 'worksheets/sheet1.xml'],
      ['worksheet', 'worksheets/sheet2.xml'],
      ['styles', 'styles.xml'],
    ])],
    ['xl/worksheets/sheet1.xml', `${XML}<worksheet xmlns="${S}"><sheetData>`
      + '<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'],
    ['xl/worksheets/sheet2.xml', `${XML}${MALFORMED}`],
    ['xl/styles.xml', `${XML}<styleSheet xmlns="${S}">`
      + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
      + '</styleSheet>'],
  ]);
}
