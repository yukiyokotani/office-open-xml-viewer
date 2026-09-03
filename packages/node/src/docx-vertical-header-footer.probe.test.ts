/**
 * Vertical (tbRl) header/footer render probe — ECMA-376 §17.6.20 + §17.10.1,
 * issue #988 batch-3 adjudication.
 *
 * Word ground truth: a section with `<w:textDirection w:val="tbRl"/>` flows its
 * BODY vertically (glyphs stack top→bottom, columns right→left) but keeps its
 * header/footer HORIZONTAL at the PHYSICAL top/bottom margins, centred, with the
 * PAGE field horizontal — the section direction does NOT propagate to them.
 *
 * The synthetic `.docx` (built in-memory, no file committed) is a tbRl section
 * with CJK body text plus a centred header ("HEADERMARK") and footer
 * ("FOOTERMARK"). We capture the renderer's `onTextRun` reports and assert:
 *   - the header run is HORIZONTAL (no `rotate(90deg)` transform) at the physical
 *     top band, horizontally centred on the page;
 *   - the footer run is HORIZONTAL at the physical bottom band, centred;
 *   - the BODY runs ARE vertical (carry the `rotate(90deg)` transform), proving
 *     the section really is vertical — so the header/footer horizontality is the
 *     tested behaviour, not a non-vertical page.
 *
 * CI-safe: gated on the docx WASM (built by `pnpm build:wasm`) + skia-canvas
 * (devDependency). Skips when either is absent; hard-fails under OOXML_REQUIRE_SKIA=1.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function storedZip(files: Record<string, string>): Uint8Array {
  const enc = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...enc.encode(name)];
    const data = [...enc.encode(content)];
    const crc = crc32(Uint8Array.from(data)) >>> 0;
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    ];
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
      ...nameBytes,
    );
    chunks.push(...local);
    offset += local.length;
  }
  const centralOffset = offset;
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(central.length), ...u32(centralOffset), ...u16(0),
  ];
  return Uint8Array.from([...chunks, ...central, ...end]);
}

interface DocxOpts {
  /** Include a header + footer part & references. */
  hf: boolean;
  /** Body paragraph repeat count (drives page count). */
  paras?: number;
  /** pgMar attributes — default asymmetric so a logical-axis reserve would show. */
  pgMar?: string;
}

function verticalHeaderFooterDocx(opts: DocxOpts): Uint8Array {
  const hf = opts.hf;
  const paras = opts.paras ?? 2;
  const pgMar =
    opts.pgMar ??
    'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"';
  const hfOverrides = hf
    ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    : '';
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    hfOverrides +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    (hf
      ? '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
      : '') +
    '</Relationships>';
  const bodyPara =
    '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr>' +
    '<w:t>縦書き本文の段落。縦書き本文の段落。縦書き本文の段落。</w:t></w:r></w:p>';
  const hfRefs = hf
    ? '<w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/>'
    : '';
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    bodyPara.repeat(paras) +
    // sectPr in XSD sequence order: headerReference → footerReference → pgSz →
    // pgMar → textDirection.
    '<w:sectPr>' +
    hfRefs +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    `<w:pgMar ${pgMar}/>` +
    '<w:textDirection w:val="tbRl"/>' +
    '</w:sectPr></w:body></w:document>';
  const header =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>` +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>HEADERMARK</w:t></w:r></w:p></w:hdr>';
  const footer =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}>` +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>FOOTERMARK</w:t></w:r></w:p></w:ftr>';
  const files: Record<string, string> = {
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
  };
  if (hf) {
    files['word/header1.xml'] = header;
    files['word/footer1.xml'] = footer;
  }
  return storedZip(files);
}

interface Run { t: string; x: number; y: number; w: number; h: number; tr?: string }

async function renderRuns(): Promise<{ runs: Run[] }> {
  const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
  const { renderDocumentToCanvas } = rendererMod!;
  const doc = await materializeDocxDocument(verticalHeaderFooterDocx({ hf: true }));
  // Physical letter portrait: width 612 (=pageWidth), height 792 (=pageHeight).
  const canvas = new Canvas(Math.round(doc.section.pageWidth), Math.round(doc.section.pageHeight));
  const runs: Run[] = [];
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as Any, 0, {
      dpr: 1,
      width: doc.section.pageWidth, // scale 1 px/pt on the physical page
      onTextRun: (r: Any) =>
        runs.push({ t: r.text, x: r.x, y: r.y, w: r.w, h: r.h, tr: r.transform }),
    });
  } finally {
    restoreOff();
    restoreImg();
  }
  return { runs };
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx vertical header/footer stay horizontal at physical top/bottom (§17.6.20 + §17.10.1)',
  () => {
    it('draws the header horizontal at the physical top, the footer at the bottom, body vertical', async () => {
      const { runs } = await renderRuns();
      const pageW = 612; // physical letter width (pt == px at scale 1)
      const pageH = 792;
      const center = pageW / 2;

      const header = runs.find((r) => r.t.includes('HEADERMARK'));
      const footer = runs.find((r) => r.t.includes('FOOTERMARK'));
      const body = runs.find((r) => r.t.includes('縦書き'));

      expect(header, 'header run reported').toBeTruthy();
      expect(footer, 'footer run reported').toBeTruthy();
      expect(body, 'body run reported').toBeTruthy();
      if (!header || !footer || !body) return;

      const isRotated = (tr?: string) => !!tr && /rotate/.test(tr);

      // Header: horizontal (no rotate), physical top band, centred.
      expect(isRotated(header.tr), 'header must be horizontal (no rotate)').toBe(false);
      expect(header.y, 'header near physical top').toBeLessThan(pageH * 0.12);
      const headerMid = header.x + header.w / 2;
      expect(Math.abs(headerMid - center), 'header centred horizontally').toBeLessThan(pageW * 0.1);

      // Footer: horizontal, physical bottom band, centred.
      expect(isRotated(footer.tr), 'footer must be horizontal (no rotate)').toBe(false);
      expect(footer.y, 'footer near physical bottom').toBeGreaterThan(pageH * 0.85);
      const footerMid = footer.x + footer.w / 2;
      expect(Math.abs(footerMid - center), 'footer centred horizontally').toBeLessThan(pageW * 0.1);

      // Body: vertical (rotate transform present) — the section really is tbRl.
      expect(isRotated(body.tr), 'body must be vertical (rotate present)').toBe(true);
    });

    it('paginates a vertical section independently of its header/footer (asymmetric margins)', async () => {
      // The header/footer occupy the PHYSICAL top/bottom band and reserve no body
      // space (issue #988). The paginator runs on the swapped logical geometry, so a
      // logical-axis reserve — comparing header height against the logical marginTop
      // (physical RIGHT) — would spuriously fire with asymmetric margins and add
      // pages the paint (reserve 0) does not expect. Assert the page count is the
      // same with and without the header/footer, on an asymmetric-margin page.
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
      const { layoutDocument } = rendererMod!;
      const margins =
        'w:top="2880" w:right="720" w:bottom="2880" w:left="720" w:header="360" w:footer="360" w:gutter="0"';
      // Keep the corpus comfortably beyond a single logical page. A smaller
      // paragraph count can fit when the selected CJK fallback has compact
      // metrics, making this pagination invariant depend on the host fonts.
      const withHF = await materializeDocxDocument(verticalHeaderFooterDocx({ hf: true, paras: 80, pgMar: margins }));
      const noHF = await materializeDocxDocument(verticalHeaderFooterDocx({ hf: false, paras: 80, pgMar: margins }));
      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      let withPages = 0;
      let noPages = 0;
      try {
        withPages = layoutDocument(withHF).pages.length;
        noPages = layoutDocument(noHF).pages.length;
      } finally {
        rOff();
        rImg();
      }
      expect(withPages, 'multi-page body').toBeGreaterThan(1);
      expect(withPages, 'header/footer do not perturb vertical pagination').toBe(noPages);
    }, 15_000);
  },
);
