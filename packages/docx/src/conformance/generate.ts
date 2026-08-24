import type { ConformanceAxisValues, ConformanceCase } from './cases.js';

const encoder = new TextEncoder();
const FIXED_DOS_TIME = 0;
const FIXED_DOS_DATE = 0x0021; // 1980-01-01

const CONTENT_TYPES_NS =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const PACKAGE_REL_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const WORD_NS =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const WORD_DRAWING_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PICTURE_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/picture';

const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_STYLES =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const REL_THEME =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const REL_HEADER =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const REL_IMAGE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const ONE_PIXEL_PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41,
  0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
  0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99,
  0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
  0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

function xml(value: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${value}`);
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function paragraphProperties(axes: ConformanceAxisValues): string {
  const style = axes.styleSource === 'paragraphStyle'
    ? '<w:pStyle w:val="CaseStyle"/>'
    : '';
  if (axes.styleSource !== 'direct') return `<w:pPr>${style}</w:pPr>`;
  return `<w:pPr>${style}${paragraphPropertyValues(axes)}</w:pPr>`;
}

function paragraphPropertyValues(axes: ConformanceAxisValues): string {
  const bidi = axes.direction === 'rtl' ? '<w:bidi/>' : '';
  const spacing = axes.spacing === 'exact'
    ? '<w:spacing w:after="120" w:line="240" w:lineRule="exact"/>'
    : '<w:spacing w:after="120" w:line="240" w:lineRule="auto"/>';
  return `${bidi}${spacing}`;
}

function runProperties(axes: ConformanceAxisValues): string {
  if (axes.fontSource === 'documentDefault') return '';
  const fonts = axes.fontSource === 'theme'
    ? '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
    : '<w:rFonts w:ascii="Ahem" w:hAnsi="Ahem"/>';
  return `<w:rPr>${fonts}<w:sz w:val="20"/></w:rPr>`;
}

function targetText(testCase: ConformanceCase): string {
  const base = testCase.expected.targetText;
  return testCase.axes.paragraph === 'wrapped'
    ? `${base} ${'PAIRWISE CONFORMANCE '.repeat(18).trim()}`
    : base;
}

function drawing(axes: ConformanceAxisValues): string {
  if (axes.object === 'none') return '';
  const extent = '<wp:extent cx="457200" cy="274320"/>';
  const graphic = `<wp:docPr id="1" name="Synthetic conformance object"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic xmlns:a="${DRAWING_NS}">
      <a:graphicData uri="${PICTURE_NS}">
        <pic:pic xmlns:pic="${PICTURE_NS}">
          <pic:nvPicPr>
            <pic:cNvPr id="1" name="pixel.png"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="rIdImage"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="274320"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>`;
  if (axes.object === 'inline') {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
      ${extent}${graphic}
    </wp:inline></w:drawing></w:r>`;
  }
  const reference = axes.anchorReference;
  if (reference === 'none') throw new Error('floating conformance object requires a reference');
  // `paragraph` is an ST_RelFromV value, not ST_RelFromH (§20.4.3.2 versus
  // §20.4.3.3). Keep the horizontal axis schema-valid while the case exercises
  // the requested vertical reference frame.
  const horizontalReference = reference === 'paragraph' ? 'column' : reference;
  return `<w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="251658240"
      behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"
      distT="0" distB="0" distL="0" distR="0">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="${horizontalReference}"><wp:posOffset>914400</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="${reference}"><wp:posOffset>0</wp:posOffset></wp:positionV>
    ${extent}<wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapSquare wrapText="bothSides"/>
    ${graphic}
  </wp:anchor></w:drawing></w:r>`;
}

function paragraph(testCase: ConformanceCase, text = targetText(testCase)): string {
  return `<w:p>
    ${paragraphProperties(testCase.axes)}
    <w:r>${runProperties(testCase.axes)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
    ${drawing(testCase.axes)}
  </w:p>`;
}

function bodyContent(testCase: ConformanceCase): string {
  if (testCase.axes.story !== 'body') {
    return '<w:p><w:r><w:t>BODY_WITNESS</w:t></w:r></w:p>';
  }
  const target = paragraph(testCase);
  if (testCase.axes.container === 'paragraph') return target;
  if (testCase.axes.container === 'table') {
    return table(target, false);
  }
  const inner = table(target, true);
  return table(`<w:p><w:r><w:t>NESTED_TABLE_PREFIX</w:t></w:r></w:p>${inner}
    <w:p><w:r><w:t>NESTED_TABLE_SUFFIX</w:t></w:r></w:p>`, false);
}

function table(cellContent: string, inner: boolean): string {
  const width = inner ? 3600 : 7200;
  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${width}" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>
      </w:tblBorders>
      <w:tblLayout w:type="fixed"/>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid>
    <w:tr><w:tc>
      <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>
      ${cellContent}
    </w:tc></w:tr>
  </w:tbl>`;
}

function styles(testCase: ConformanceCase): Uint8Array {
  const axes = testCase.axes;
  const defaultParagraph = axes.styleSource === 'documentDefault'
    ? `<w:pPrDefault><w:pPr>${paragraphPropertyValues(axes)}</w:pPr></w:pPrDefault>`
    : '<w:pPrDefault><w:pPr/></w:pPrDefault>';
  const defaultRun = axes.fontSource === 'documentDefault'
    ? '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Ahem" w:hAnsi="Ahem"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>'
    : '<w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>';
  const styleParagraph = axes.styleSource === 'paragraphStyle'
    ? paragraphPropertyValues(axes)
    : '';
  return xml(`<w:styles xmlns:w="${WORD_NS}">
    <w:docDefaults>${defaultRun}${defaultParagraph}</w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="CaseStyle">
      <w:name w:val="Case Style"/>
      <w:basedOn w:val="Normal"/>
      <w:pPr>${styleParagraph}</w:pPr>
    </w:style>
  </w:styles>`);
}

function theme(): Uint8Array {
  return xml(`<a:theme xmlns:a="${DRAWING_NS}" name="Synthetic Conformance">
    <a:themeElements>
      <a:clrScheme name="Synthetic">
        <a:dk1><a:srgbClr val="000000"/></a:dk1>
        <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
        <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
        <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
      </a:clrScheme>
      <a:fontScheme name="Synthetic">
        <a:majorFont><a:latin typeface="Ahem"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
        <a:minorFont><a:latin typeface="Ahem"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
      </a:fontScheme>
      <a:fmtScheme name="Synthetic">
        <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
        <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
        <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
        <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
      </a:fmtScheme>
    </a:themeElements>
  </a:theme>`);
}

function contentTypes(testCase: ConformanceCase): Uint8Array {
  const storyOverride = testCase.axes.story === 'header'
    ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    : testCase.axes.story === 'footer'
      ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      : '';
  return xml(`<Types xmlns="${CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    ${storyOverride}
  </Types>`);
}

function documentRelationships(testCase: ConformanceCase): Uint8Array {
  const storyRelationship = testCase.axes.story === 'header'
    ? `<Relationship Id="rIdHeader" Type="${REL_HEADER}" Target="header1.xml"/>`
    : testCase.axes.story === 'footer'
      ? `<Relationship Id="rIdFooter" Type="${REL_FOOTER}" Target="footer1.xml"/>`
      : '';
  const imageRelationship = testCase.axes.object === 'none'
    ? ''
    : `<Relationship Id="rIdImage" Type="${REL_IMAGE}" Target="media/pixel.png"/>`;
  return xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
    <Relationship Id="rIdStyles" Type="${REL_STYLES}" Target="styles.xml"/>
    <Relationship Id="rIdTheme" Type="${REL_THEME}" Target="theme/theme1.xml"/>
    ${storyRelationship}${imageRelationship}
  </Relationships>`);
}

function documentXml(testCase: ConformanceCase): Uint8Array {
  const storyReference = testCase.axes.story === 'header'
    ? '<w:headerReference w:type="default" r:id="rIdHeader"/>'
    : testCase.axes.story === 'footer'
      ? '<w:footerReference w:type="default" r:id="rIdFooter"/>'
      : '';
  return xml(`<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"
      xmlns:wp="${WORD_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:pic="${PICTURE_NS}">
    <w:body>
      ${bodyContent(testCase)}
      <w:sectPr>
        ${storyReference}
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
          w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr>
    </w:body>
  </w:document>`);
}

function storyXml(testCase: ConformanceCase): Uint8Array {
  const root = testCase.axes.story === 'header' ? 'hdr' : 'ftr';
  return xml(`<w:${root} xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"
      xmlns:wp="${WORD_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:pic="${PICTURE_NS}">
    ${paragraph(testCase)}
  </w:${root}>`);
}

export function generateConformanceParts(
  testCase: ConformanceCase,
): ReadonlyMap<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', contentTypes(testCase)],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="word/document.xml"/>
    </Relationships>`)],
    ['word/_rels/document.xml.rels', documentRelationships(testCase)],
    ['word/document.xml', documentXml(testCase)],
    ['word/styles.xml', styles(testCase)],
    ['word/theme/theme1.xml', theme()],
  ]);
  if (testCase.axes.story !== 'body') {
    parts.set(`word/${testCase.axes.story}1.xml`, storyXml(testCase));
  }
  if (testCase.axes.object !== 'none') {
    parts.set('word/media/pixel.png', ONE_PIXEL_PNG);
  }
  return new Map([...parts].sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0));
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value: number): Uint8Array {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function u32(value: number): Uint8Array {
  return Uint8Array.of(
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  );
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const length = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const output = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

/**
 * Minimal deterministic ZIP writer using stored entries.
 *
 * Avoiding DEFLATE makes the byte stream independent of native zlib versions.
 * Entry order, timestamps, flags, creator version and attributes are all fixed.
 */
export function storeZip(parts: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const localChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  let offset = 0;

  for (const [name, data] of [...parts].sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0)) {
    const encodedName = encoder.encode(name);
    const checksum = crc32(data);
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0x0800),
      u16(0),
      u16(FIXED_DOS_TIME),
      u16(FIXED_DOS_DATE),
      u32(checksum),
      u32(data.byteLength),
      u32(data.byteLength),
      u16(encodedName.byteLength),
      u16(0),
      encodedName,
      data,
    ]);
    localChunks.push(local);
    centralChunks.push(concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0x0800),
      u16(0),
      u16(FIXED_DOS_TIME),
      u16(FIXED_DOS_DATE),
      u32(checksum),
      u32(data.byteLength),
      u32(data.byteLength),
      u16(encodedName.byteLength),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      encodedName,
    ]));
    offset += local.byteLength;
  }

  const central = concat(centralChunks);
  return concat([
    ...localChunks,
    central,
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(parts.size),
    u16(parts.size),
    u32(central.byteLength),
    u32(offset),
    u16(0),
  ]);
}

export function generateConformanceDocx(testCase: ConformanceCase): Uint8Array {
  return storeZip(generateConformanceParts(testCase));
}

// ── Redistributable tracked-changes / comments fixtures ─────────────────────
//
// Deterministic synthetic documents for the ECMA-376 §17.13.5 (revisions) and
// §17.13.4 (comments) features: no real-document content, fixed authors and
// dates, stored-zip bytes stable across runs. Used by the WASM-backed layout
// tests and the Storybook demo; nothing binary is committed.

const REL_COMMENTS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_COMMENTS_EXTENDED =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const REL_FOOTNOTES =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const REL_ENDNOTES =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

const FIXTURE_SECTION = `<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
    w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>`;

const FIXTURE_STYLES = `<w:styles xmlns:w="${WORD_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Ahem" w:hAnsi="Ahem"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr/></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

function fixtureContentTypes(extra: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const commentOverrides = `${extra.has('word/comments.xml')
    ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    : ''}${extra.has('word/commentsExtended.xml')
    ? '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
    : ''}`;
  const storyOverrides = [
    extra.has('word/header1.xml')
      ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      : '',
    extra.has('word/footer1.xml')
      ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      : '',
    extra.has('word/footnotes.xml')
      ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      : '',
    extra.has('word/endnotes.xml')
      ? '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
      : '',
  ].join('');
  return xml(`<Types xmlns="${CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    ${commentOverrides}${storyOverrides}
  </Types>`);
}

function fixtureParts(
  documentBody: string,
  extra: ReadonlyMap<string, Uint8Array> = new Map(),
  extraRelationships = '',
  section = FIXTURE_SECTION,
): ReadonlyMap<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', fixtureContentTypes(extra)],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="word/document.xml"/>
    </Relationships>`)],
    ['word/_rels/document.xml.rels', xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdStyles" Type="${REL_STYLES}" Target="styles.xml"/>
      ${extraRelationships}
    </Relationships>`)],
    ['word/document.xml', xml(`<w:document xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"
      xmlns:r="${OFFICE_REL_NS}" xmlns:v="urn:schemas-microsoft-com:vml">
      <w:body>${documentBody}${section}</w:body>
    </w:document>`)],
    ['word/styles.xml', xml(FIXTURE_STYLES)],
    ...extra,
  ]);
  return new Map([...parts].sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0));
}

/**
 * A document exercising every §17.13.5 run-revision kind: an insertion and a
 * deletion by Alice, and a Bob move represented by its moveFrom/moveTo pair.
 * The rendered final state is "Kept inserted moved-in tail"; all four revision
 * kinds remain available in the parsed model for consumer-owned review UI.
 */
export function generateTrackedChangesDocx(): Uint8Array {
  const body = `
    <w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t xml:space="preserve">inserted </w:t></w:r>
      </w:ins>
      <w:del w:id="2" w:author="Alice" w:date="2024-01-02T00:00:00Z">
        <w:r><w:delText xml:space="preserve">deleted </w:delText></w:r>
      </w:del>
      <w:moveFromRangeStart w:id="30" w:name="move1" w:author="Bob" w:date="2024-01-03T00:00:00Z"/>
      <w:moveFrom w:id="3" w:author="Bob" w:date="2024-01-03T00:00:00Z">
        <w:r><w:t xml:space="preserve">moved-in </w:t></w:r>
      </w:moveFrom>
      <w:moveFromRangeEnd w:id="30"/>
      <w:moveToRangeStart w:id="40" w:name="move1" w:author="Bob" w:date="2024-01-03T00:00:00Z"/>
      <w:moveTo w:id="4" w:author="Bob" w:date="2024-01-03T00:00:00Z">
        <w:r><w:t xml:space="preserve">moved-in </w:t></w:r>
      </w:moveTo>
      <w:moveToRangeEnd w:id="40"/>
      <w:r><w:t>tail</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Unrevised second paragraph.</w:t></w:r></w:p>`;
  return storeZip(fixtureParts(body));
}

/**
 * A document exercising §17.13.4 comment anchors with commentsExtended
 * threading: comment 1 (Alice) with a reply (Bob) anchored on a mid-paragraph
 * range, and a resolved comment 3 (Carol).
 */
export function generateCommentedDocx(): Uint8Array {
  const body = `
    <w:p>
      <w:r><w:t xml:space="preserve">Before </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t xml:space="preserve">annotated text</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t xml:space="preserve"> after.</w:t></w:r>
    </w:p>
    <w:p>
      <w:commentRangeStart w:id="3"/>
      <w:r><w:t>Resolved-thread anchor.</w:t></w:r>
      <w:commentRangeEnd w:id="3"/>
      <w:r><w:commentReference w:id="3"/></w:r>
    </w:p>`;
  const comments = xml(`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}">
    <w:comment w:id="1" w:author="Alice" w:initials="A" w:date="2024-03-01T10:00:00Z">
      <w:p w14:paraId="00000A01"><w:r><w:t>Please review this wording.</w:t></w:r></w:p>
    </w:comment>
    <w:comment w:id="2" w:author="Bob" w:initials="B" w:date="2024-03-01T11:00:00Z">
      <w:p w14:paraId="00000B01"><w:r><w:t>Agreed, second sentence reads better.</w:t></w:r></w:p>
    </w:comment>
    <w:comment w:id="3" w:author="Carol" w:initials="C" w:date="2024-03-02T09:00:00Z">
      <w:p w14:paraId="00000C01"><w:r><w:t>Old resolved note.</w:t></w:r></w:p>
    </w:comment>
  </w:comments>`);
  const commentsExtended = xml(`<w15:commentsEx xmlns:w15="${W15_NS}">
    <w15:commentEx w15:paraId="00000A01" w15:done="0"/>
    <w15:commentEx w15:paraId="00000B01" w15:paraIdParent="00000A01" w15:done="0"/>
    <w15:commentEx w15:paraId="00000C01" w15:done="1"/>
  </w15:commentsEx>`);
  return storeZip(fixtureParts(
    body,
    new Map([
      ['word/comments.xml', comments],
      ['word/commentsExtended.xml', commentsExtended],
    ]),
    `<Relationship Id="rIdComments" Type="${REL_COMMENTS}" Target="comments.xml"/>
     <Relationship Id="rIdCommentsExtended" Type="${REL_COMMENTS_EXTENDED}" Target="commentsExtended.xml"/>`,
  ));
}

/** A real-parser fixture with one valid comment anchor in every retained DOCX
 * story: body, header, footer, footnote, endnote, and a VML text box. */
export function generateAllStoryCommentsDocx(): Uint8Array {
  const anchored = (id: number, text: string): string => `<w:p>
    <w:commentRangeStart w:id="${id}"/>
    <w:r><w:t>${text}</w:t></w:r>
    <w:commentRangeEnd w:id="${id}"/>
    <w:r><w:commentReference w:id="${id}"/></w:r>
  </w:p>`;
  const body = `${anchored(1, 'Body anchor')}
    <w:p>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t xml:space="preserve"> </w:t></w:r>
      <w:r><w:endnoteReference w:id="1"/></w:r>
    </w:p>
    <w:p><w:r><w:pict>
      <v:shape id="review-textbox" type="#_x0000_t202"
        style="position:relative;width:180pt;height:36pt" filled="f" stroked="f">
        <v:textbox><w:txbxContent>${anchored(6, 'Text box anchor')}</w:txbxContent></v:textbox>
      </v:shape>
    </w:pict></w:r></w:p>`;
  const comments = xml(`<w:comments xmlns:w="${WORD_NS}">
    ${[1, 2, 3, 4, 5, 6].map((id) => `<w:comment w:id="${id}" w:author="Reviewer">
      <w:p><w:r><w:t>Comment ${id}</w:t></w:r></w:p>
    </w:comment>`).join('')}
  </w:comments>`);
  const extra = new Map<string, Uint8Array>([
    ['word/comments.xml', comments],
    ['word/header1.xml', xml(`<w:hdr xmlns:w="${WORD_NS}">${anchored(2, 'Header anchor')}</w:hdr>`)],
    ['word/footer1.xml', xml(`<w:ftr xmlns:w="${WORD_NS}">${anchored(3, 'Footer anchor')}</w:ftr>`)],
    ['word/footnotes.xml', xml(`<w:footnotes xmlns:w="${WORD_NS}">
      <w:footnote w:id="1">${anchored(4, 'Footnote anchor')}</w:footnote>
    </w:footnotes>`)],
    ['word/endnotes.xml', xml(`<w:endnotes xmlns:w="${WORD_NS}">
      <w:endnote w:id="1">${anchored(5, 'Endnote anchor')}</w:endnote>
    </w:endnotes>`)],
  ]);
  const relationships = `
    <Relationship Id="rIdComments" Type="${REL_COMMENTS}" Target="comments.xml"/>
    <Relationship Id="rIdHeader" Type="${REL_HEADER}" Target="header1.xml"/>
    <Relationship Id="rIdFooter" Type="${REL_FOOTER}" Target="footer1.xml"/>
    <Relationship Id="rIdFootnotes" Type="${REL_FOOTNOTES}" Target="footnotes.xml"/>
    <Relationship Id="rIdEndnotes" Type="${REL_ENDNOTES}" Target="endnotes.xml"/>`;
  const section = `<w:sectPr>
    <w:headerReference w:type="default" r:id="rIdHeader"/>
    <w:footerReference w:type="default" r:id="rIdFooter"/>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>`;
  return storeZip(fixtureParts(body, extra, relationships, section));
}
