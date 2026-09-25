/**
 * ECMA-376 §17.6.5 docGrid line/column pitch through the REAL parse→render
 * path — the issue #1013 adjudication (sample-58 sweep, Word PDF ground truth).
 *
 * Word measurement (pdftotext -bbox over a 19-section synthetic matrix of
 * {10.5,12,14,16,20}pt × linePitch {18,24}pt × {lrTb,tbRl} × {lines,
 * linesAndChars,none}, all Yu Mincho): a single-spaced East Asian line
 * occupies ceil(designSingleLineHeight / pitch) whole grid cells, where the
 * design height for Yu Mincho is 1.3 × its hhea box = 1.43267 em (core
 * line-metrics). The measured matrix:
 *
 *   pitch 18pt: 10.5/12pt → 18pt (1 cell);  14/16/20pt → 36pt (2 cells)
 *   pitch 24pt: 12/16pt   → 24pt (1 cell);  20pt       → 48pt (2 cells)
 *   no grid   : natural 1.43267 em (e.g. 16pt → 22.92pt)
 *
 * HORIZONTAL and VERTICAL (tbRl) sections measured IDENTICAL pitches — the
 * tbRl column pitch is the same grid cell height (the vertical page is the
 * horizontal layout rotated), so these probes pin BOTH directions through
 * renderDocumentToCanvas. The pre-#1013 em-based rule (floor(em/pitch)+1)
 * rendered every 14–16pt line 1 cell (18pt) — 2× too tight vs Word's 36pt.
 *
 * CI-safe: gated on docx WASM + skia-canvas; skips when absent.
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

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

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

/** A one-section docx: a single Yu Mincho paragraph of 4 CJK lines joined by
 *  <w:br/>. Line spacing is single via docDefaults ONLY (inherited-only, the
 *  state that grid-snaps per §17.6.5 — an explicit per-pPr w:line would take
 *  the multiplier path instead, matching Word). */
function docxWith(sizePt: number, sectPr: string, lineText?: string): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';
  const styles =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles ${NS}><w:docDefaults>` +
    '<w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Yu Mincho" w:hAnsi="Yu Mincho" w:eastAsia="游明朝"/>' +
    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults></w:styles>';
  const sz = String(Math.round(sizePt * 2));
  const rpr = `<w:rPr><w:rFonts w:ascii="Yu Mincho" w:hAnsi="Yu Mincho" w:eastAsia="游明朝"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`;
  const line = lineText ?? '国境の長いトンネルを抜けると雪国であった';
  const runs = Array.from({ length: 4 }, (_, i) =>
    `<w:r>${rpr}<w:t xml:space="preserve">${line}</w:t></w:r>` +
    (i < 3 ? `<w:r>${rpr}<w:br/></w:r>` : ''),
  ).join('');
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    `<w:p><w:pPr>${rpr}</w:pPr>${runs}</w:p>` +
    `<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`;
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
    'word/styles.xml': styles,
  });
}

const A4 = '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

function sectPr(opts: { vertical?: boolean; pitchTw?: number; gridType?: string }): string {
  return A4 +
    (opts.vertical ? '<w:textDirection w:val="tbRl"/>' : '') +
    (opts.pitchTw ? `<w:docGrid w:type="${opts.gridType ?? 'lines'}" w:linePitch="${opts.pitchTw}"/>` : '');
}

/** Render page 0 and return the mean gap between consecutive line positions —
 *  physical y for horizontal pages, physical x (column advance, right→left
 *  page rotated to +x order) for vertical ones. */
async function measurePitch(bytes: Uint8Array, axis: 'y' | 'x', marker = '国境'): Promise<number> {
  const { materializeDocxDocument } = docxMod!;
  const { createLayoutServices, renderDocumentToCanvas } = rendererMod!;
  const doc = await materializeDocxDocument(bytes);
  const canvas = new Canvas(10, 10);
  const lineHeightRatio = (2257 * 1.3) / 2048;
  const localMetrics = Object.fromEntries(['Yu Mincho', '游明朝'].flatMap((family) => {
    const key = family.toLowerCase();
    const metric = {
      family: 'serif', requestedFamily: family, weight: 400, style: 'normal' as const,
      // Both authored aliases resolve to the same CSS face in this test
      // backend, so they must agree on its vertical geometry.
      lineHeightRatio,
      sourceIdentity: 'test-fixture:node-skia-generic-serif',
      synthesized: false,
    };
    return [[key, metric]];
  }));
  const layoutServices = createLayoutServices(doc, {
    localMetrics,
    measureContext: canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
  });
  const pos: number[] = [];
  const rImg = installImageBitmapShim(factory);
  const rOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr: 1,
      width: doc.section.pageWidth,
      onTextRun: (r: Any) => {
        if (String(r.text).includes(marker)) pos.push(axis === 'y' ? r.y : r.x);
      },
      layoutServices,
    });
  } finally {
    rOff();
    rImg();
  }
  const uniq = [...new Set(pos.map((v) => Math.round(v * 100) / 100))].sort((a, b) => a - b);
  // Cluster to line centers (per-glyph runs share the line position ±ε).
  const centers: number[] = [];
  for (const v of uniq) {
    if (centers.length === 0 || v - centers[centers.length - 1] > 2) centers.push(v);
  }
  expect(centers.length).toBe(4);
  const gaps = centers.slice(1).map((v, i) => v - centers[i]);
  return gaps.reduce((a, b) => a + b, 0) / gaps.length;
}

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx docGrid line/column pitch (§17.6.5, issue #1013 sample-58 adjudication)',
  () => {
    it('HORIZONTAL: a 16pt line on an 18pt grid takes 2 cells (36pt, was 18pt)', async () => {
      const pitch = await measurePitch(docxWith(16, sectPr({ pitchTw: 360 })), 'y');
      expect(pitch).toBeCloseTo(36, 1);
    });
    it('HORIZONTAL: a 12pt line on an 18pt grid stays 1 cell (18pt)', async () => {
      const pitch = await measurePitch(docxWith(12, sectPr({ pitchTw: 360 })), 'y');
      expect(pitch).toBeCloseTo(18, 1);
    });
    it('VERTICAL (tbRl): a 16pt column on an 18pt grid takes 2 cells (36pt) — the issue #1013 case', async () => {
      const pitch = await measurePitch(docxWith(16, sectPr({ vertical: true, pitchTw: 360 })), 'x');
      expect(pitch).toBeCloseTo(36, 1);
    });
    it('VERTICAL (tbRl): a 12pt column on an 18pt grid stays 1 cell (18pt)', async () => {
      const pitch = await measurePitch(docxWith(12, sectPr({ vertical: true, pitchTw: 360 })), 'x');
      expect(pitch).toBeCloseTo(18, 1);
    });
    it('a 16pt line on a 24pt grid stays 1 cell (24pt) — pitch-scaled, not size-thresholded', async () => {
      const pitch = await measurePitch(docxWith(16, sectPr({ pitchTw: 480 })), 'y');
      expect(pitch).toBeCloseTo(24, 1);
    });
    it('no docGrid: the Yu Mincho design single line (1.43267 em; 16pt → 22.92pt)', async () => {
      const pitch = await measurePitch(docxWith(16, sectPr({})), 'y');
      expect(pitch).toBeCloseTo(16 * ((2257 * 1.3) / 2048), 1);
    });
    it('a 20pt line on a 24pt grid takes 2 cells (48pt) — sample-58 C3', async () => {
      const pitch = await measurePitch(docxWith(20, sectPr({ vertical: true, pitchTw: 480 })), 'x');
      expect(pitch).toBeCloseTo(48, 1);
    });
    it('linesAndChars behaves like lines (grid-type-agnostic; 16pt/18pt → 36pt) — sample-58 D2', async () => {
      const pitch = await measurePitch(
        docxWith(16, sectPr({ vertical: true, pitchTw: 360, gridType: 'linesAndChars' })), 'x');
      expect(pitch).toBeCloseTo(36, 1);
    });
    it('a LATIN-only Yu Mincho line keeps its natural box above a one-pitch floor, never FE cells', async () => {
      // §17.6.5 leaves Latin lines at default spacing above a one-pitch floor
      // (max(natural, pitch)), and the eaOnly Yu Mincho FE height must NOT
      // fire for a Latin line (demo/sample-1 footnote adjudication: Word gives
      // Latin Yu Mincho the win box, not 1.43267 em). The Latin natural is the
      // SUBSTITUTED font's box (environment-dependent in this headless run),
      // so calibrate it from the no-grid render and assert the on-grid value
      // is exactly max(natural, pitch) — an FE regression would instead
      // cell-round 16pt to 2 cells (36pt).
      const latin = 'The quick brown fox jumps over the dog';
      const natural = await measurePitch(docxWith(16, sectPr({}), latin), 'y', 'quick');
      const pitch = await measurePitch(docxWith(16, sectPr({ pitchTw: 360 }), latin), 'y', 'quick');
      expect(pitch).toBeCloseTo(Math.max(natural, 18), 1);
      expect(pitch).toBeLessThan(36);
    });
  },
);
