import { readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { build } from 'vite';
import {
  chartExDocxBytes,
  chartExPptxBytes,
  chartExXlsxBytes,
  storedZip,
} from '../tests/fixtures/chart-ex-packages.mjs';

import { pathGradientPptxBytes } from '../tests/fixtures/path-gradient-package.mjs';
// Node strips the fixture module's type annotations; it has no imports.
import { buildDocFixture } from '../packages/legacy-office/src/test-fixtures.ts';

const root = resolve(new URL('..', import.meta.url).pathname);
const outDir = join(tmpdir(), 'ooxml-worker-consumer-dist');
const entry = (name) => resolve(root, `dist/${name}.mjs`);


await build({
  configFile: false,
  root: resolve(root, 'tests/worker-dist/consumer'),
  base: './',
  resolve: {
    alias: {
      // Nuxt (and many other apps) alias `assets` to their own folder. A
      // bundled sibling URL without `./` would resolve there and the render
      // worker would silently not be emitted (#2134).
      assets: resolve(root, 'tests/worker-dist/consumer/app-assets'),
      '@silurus/ooxml/docx': entry('docx'),
      '@silurus/ooxml/xlsx': entry('xlsx'),
      '@silurus/ooxml/pptx': entry('pptx'),
      '@silurus/ooxml/math': entry('math'),
      '@silurus/ooxml/three-d': entry('three-d'),
      '@silurus/ooxml/region-map': entry('region-map'),
      '@silurus/ooxml/chart-ex': entry('chart-ex'),
      '@silurus/ooxml-legacy-office/doc': resolve(root, 'packages/legacy-office/dist/doc.mjs'),
      '@silurus/ooxml/internal/companion': entry('internal-companion'),
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

// The legacy readers' own synthetic Word 97 fixture (canonical FIB, one
// Unicode piece, root-linked CFB streams): it proves that a consumer bundle
// keeps the legacy DOC source module and its WASM loadable by URL.
writeFileSync(join(outDir, 'legacy.doc'), buildDocFixture({ text: 'Legacy worker source\rSecond paragraph\r' }));

// A tracked insertion and deletion make the final and markup views differ, so a
// model source's view default is observable in the rendered page.
writeFileSync(join(outDir, 'tracked.docx'), storedZip([
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
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r><w:t xml:space="preserve">Kept text </w:t></w:r>
          <w:del w:id="1" w:author="Reviewer" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>deleted words that only the markup view shows</w:delText></w:r></w:del>
          <w:ins w:id="2" w:author="Reviewer" w:date="2020-01-01T00:00:00Z"><w:r><w:t>inserted words</w:t></w:r></w:ins>
        </w:p>
        <w:sectPr><w:pgSz w:w="12240" w:h="4000"/></w:sectPr>
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

// Self-authored ChartEx packages exercise the renderer after descriptor
// reconstruction; see tests/fixtures/chart-ex-packages.mjs.
writeFileSync(join(outDir, 'chart-ex.xlsx'), chartExXlsxBytes());
writeFileSync(join(outDir, 'chart-ex.docx'), chartExDocxBytes());
writeFileSync(join(outDir, 'chart-ex.pptx'), chartExPptxBytes());

const workers = readdirSync(join(outDir, 'assets'))
  .filter((name) => /^render-worker-[\w-]+\.js$/.test(name)
    && !name.includes('-host-'));
if (workers.length !== 3) {
  throw new Error(`Vite consumer output must contain exactly 3 shared render workers, found ${workers.length}`);
}
// The legacy DOC model source is loaded by URL inside the parser worker; the
// consumer bundler must copy both its self-contained module and its WASM.
const legacyAssets = readdirSync(join(outDir, 'assets'))
  .filter((name) => /^legacy-doc-source-module[\w-]*\.m?js$/.test(name) || /^legacy_doc_direct_bg[\w-]*\.wasm$/.test(name));
if (legacyAssets.length !== 2) {
  throw new Error(`Vite consumer output must copy the legacy DOC source module and WASM, found ${legacyAssets.join(', ') || 'none'}`);
}
console.log(`Vite consumer bundle: ${workers.length} self-contained render workers, legacy DOC source assets`);

writeFileSync(join(outDir, 'path-gradient.pptx'), pathGradientPptxBytes());
