import { readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { build } from 'vite';

const root = resolve(new URL('..', import.meta.url).pathname);
const outDir = join(tmpdir(), 'ooxml-worker-consumer-dist');
const entry = (name) => resolve(root, `dist/${name}.mjs`);

const crc32Table = Array.from({ length: 256 }, (_, value) => {
  let crc = value;
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  return crc >>> 0;
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = (crc >>> 8) ^ crc32Table[(crc ^ byte) & 0xff];
  return (crc ^ 0xffffffff) >>> 0;
}

function storedZip(entries) {
  const encoder = new TextEncoder();
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, contents] of entries) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(contents);
    const checksum = crc32(data);
    const local = Buffer.alloc(30 + nameBytes.length + data.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.set(nameBytes, 30);
    local.set(data, 30 + nameBytes.length);
    locals.push(local);

    const directory = Buffer.alloc(46 + nameBytes.length);
    directory.writeUInt32LE(0x02014b50, 0);
    directory.writeUInt16LE(20, 4);
    directory.writeUInt16LE(20, 6);
    directory.writeUInt32LE(checksum, 16);
    directory.writeUInt32LE(data.length, 20);
    directory.writeUInt32LE(data.length, 24);
    directory.writeUInt16LE(nameBytes.length, 28);
    directory.writeUInt32LE(offset, 42);
    directory.set(nameBytes, 46);
    central.push(directory);
    offset += local.length;
  }
  const directoryOffset = offset;
  const directorySize = central.reduce((sum, part) => sum + part.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directorySize, 12);
  end.writeUInt32LE(directoryOffset, 16);
  return Buffer.concat([...locals, ...central, end]);
}

await build({
  configFile: false,
  root: resolve(root, 'tests/worker-dist/consumer'),
  base: './',
  resolve: {
    alias: {
      '@silurus/ooxml/docx': entry('docx'),
      '@silurus/ooxml/xlsx': entry('xlsx'),
      '@silurus/ooxml/pptx': entry('pptx'),
      '@silurus/ooxml/math': entry('math'),
      '@silurus/ooxml/three-d': entry('three-d'),
      '@silurus/ooxml/region-map': entry('region-map'),
      '@silurus/ooxml/chart-ex': entry('chart-ex'),
    },
  },
  build: {
    outDir,
    emptyOutDir: true,
    target: 'esnext',
  },
  logLevel: 'warn',
});

// A small self-authored package forces the production render worker to execute
// MathJax. The ordinary public demo has no equation and would only prove that
// the renderer descriptor was reconstructed, not that its external engine URL
// survived a consumer rebundle.
writeFileSync(join(outDir, 'equation.docx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
  ['word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:body>
        <w:p><w:r><w:t>Production worker equation</w:t></w:r></w:p>
        <m:oMathPara><m:oMath><m:f>
          <m:num><m:r><m:t>x+1</m:t></m:r></m:num>
          <m:den><m:r><m:t>y−1</m:t></m:r></m:den>
        </m:f></m:oMath></m:oMathPara>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      </w:body>
    </w:document>`],
]));

// A self-authored text-only presentation exercises the production PPTX worker
// without optional renderer descriptors. Keeping this separate from the public
// demo catches worker-bundle initialization bugs that optional chart renderers
// can otherwise mask by initializing shared DrawingML unit constants first.
writeFileSync(join(outDir, 'text.pptx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`],
  ['ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="5143500"/>
    </p:presentation>`],
  ['ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    </Relationships>`],
  ['ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Text Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
          </p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>
            <a:rPr lang="en-US" sz="2800" b="1"><a:latin typeface="Arial"/></a:rPr>
            <a:t>Production worker text</a:t>
          </a:r><a:endParaRPr lang="en-US" sz="2800"/></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`],
]));

// A self-authored bordered workbook exercises the production XLSX worker without
// optional renderer descriptors. Every border edge goes through the shared
// border dash lookup, so a worker bundle that strands the shared draw module's
// initializer throws on the first bordered cell. The public demo and the
// chart-ex workbook cannot catch that: one has no borders, and the other pulls
// in an optional renderer that initializes the shared draw module as a side
// effect. `thin` is deliberate — it is the style Excel emits most, and it is
// absent from the dash table (it is solid), so it reaches the lookup's miss path.
writeFileSync(join(outDir, 'bordered.xlsx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`],
  ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Bordered" sheetId="1" r:id="rIdSheet"/></sheets>
    </workbook>`],
  ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    </Relationships>`],
  ['xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="2">
        <border/>
        <border>
          <left style="thin"><color rgb="FF000000"/></left>
          <right style="thin"><color rgb="FF000000"/></right>
          <top style="thin"><color rgb="FF000000"/></top>
          <bottom style="thin"><color rgb="FF000000"/></bottom>
          <diagonal/>
        </border>
      </borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
      </cellXfs>
    </styleSheet>`],
  ['xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:B2"/>
      <sheetData>
        <row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>
        <row r="2"><c r="A2" s="1"><v>3</v></c><c r="B2" s="1"><v>4</v></c></row>
      </sheetData>
    </worksheet>`],
]));

const chartExXml = `<?xml version="1.0" encoding="UTF-8"?>
  <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    <cx:chartData><cx:data id="0">
      <cx:strDim type="cat"><cx:lvl ptCount="3">
        <cx:pt idx="0">Start</cx:pt><cx:pt idx="1">Change</cx:pt><cx:pt idx="2">End</cx:pt>
      </cx:lvl></cx:strDim>
      <cx:numDim type="val"><cx:lvl ptCount="3">
        <cx:pt idx="0">50</cx:pt><cx:pt idx="1">-15</cx:pt><cx:pt idx="2">35</cx:pt>
      </cx:lvl></cx:numDim>
    </cx:data></cx:chartData>
    <cx:chart><cx:plotArea><cx:plotAreaRegion>
      <cx:series layoutId="waterfall"/>
    </cx:plotAreaRegion></cx:plotArea></cx:chart>
  </cx:chartSpace>`;

// Self-authored ChartEx packages exercise the renderer after descriptor
// reconstruction. Colored waterfall bars distinguish the optional painter
// from the default renderer's grayscale unsupported-chart placeholder.
writeFileSync(join(outDir, 'chart-ex.xlsx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`],
  ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="ChartEx" sheetId="1" r:id="rIdSheet"/></sheets>
    </workbook>`],
  ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    </Relationships>`],
  ['xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
    </styleSheet>`],
  ['xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <dimension ref="A1:J20"/><sheetData/><drawing r:id="rIdDrawing"/>
    </worksheet>`],
  ['xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
    </Relationships>`],
  ['xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <xdr:twoCellAnchor>
        <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="ChartEx"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:chart r:id="rIdChart"/>
          </a:graphicData></a:graphic>
        </xdr:graphicFrame><xdr:clientData/>
      </xdr:twoCellAnchor>
    </xdr:wsDr>`],
  ['xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/>
    </Relationships>`],
  ['xl/charts/chartEx1.xml', chartExXml],
]));

writeFileSync(join(outDir, 'chart-ex.docx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
  ['word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="5000000" cy="3000000"/><wp:docPr id="1" name="ChartEx"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <c:chart r:id="rIdChart"/>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="9000"/></w:sectPr></w:body>
    </w:document>`],
  ['word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"/>
    </Relationships>`],
  ['word/charts/chartEx1.xml', chartExXml],
]));

writeFileSync(join(outDir, 'chart-ex.pptx'), storedZip([
  ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`],
  ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`],
  ['ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/>
    </p:presentation>`],
  ['ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    </Relationships>`],
  ['ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <p:cSld><p:spTree><p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="2" name="ChartEx"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="500000" y="500000"/><a:ext cx="8000000" cy="5500000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart r:id="rIdChart"/>
        </a:graphicData></a:graphic>
      </p:graphicFrame></p:spTree></p:cSld>
    </p:sld>`],
  ['ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/>
    </Relationships>`],
  ['ppt/charts/chartEx1.xml', chartExXml],
]));

const workers = readdirSync(join(outDir, 'assets'))
  .filter((name) => /^render-worker-[\w-]+\.js$/.test(name)
    && !name.startsWith('render-worker-host-'));
if (workers.length !== 3) {
  throw new Error(`Vite consumer output must contain 3 render workers, found ${workers.length}`);
}
console.log(`Vite consumer bundle: ${workers.length} self-contained render workers`);
