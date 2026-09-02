/**
 * Per-section text-direction mixing — ECMA-376 §17.6.20, issue #1000.
 *
 * The authored package has a non-final `btLr` section followed by a final
 * horizontal section. This parser-boundary probe deliberately creates the
 * smallest deterministic OOXML package that exercises the real WASM parser,
 * paginator, and renderer. It does not depend on a mutable private corpus.
 *
 * Both sections use a physical Letter portrait page. The first page must use
 * the vertical page paint transform and right-to-left column progression; the
 * second page must retain ordinary horizontal top-left flow.
 */
import { describe, expect, it } from 'vitest';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { minimalDocx } from './test-ooxml-package';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
} from './test-imports';

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

interface RunInfo { text: string; x: number; y: number; transform?: unknown }

const SECTION_GEOMETRY = `
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
           w:header="720" w:footer="720" w:gutter="0"/>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Vertical heading</w:t></w:r></w:p>
    <w:p><w:r><w:t>First vertical paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second vertical paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Third vertical paragraph</w:t></w:r></w:p>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:type w:val="nextPage"/>
          ${SECTION_GEOMETRY}
          <w:textDirection w:val="btLr"/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p><w:r><w:t>Horizontal control paragraph</w:t></w:r></w:p>
    <w:sectPr>
      ${SECTION_GEOMETRY}
      <w:textDirection w:val="lrTb"/>
    </w:sectPr>
  </w:body>
</w:document>`;

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx per-section text-direction mixing (§17.6.20, issue #1000)',
  () => {
    it('renders the non-final section vertically and the final section horizontally', async () => {
      const { materializeDocxDocument } = docxMod!;
      const { createLayoutServices, layoutDocument, renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(minimalDocx(DOCUMENT_XML));

      const restoreImage = installImageBitmapShim(factory);
      const restoreOffscreen = installOffscreenCanvasShim(factory);
      try {
        const layoutServices = createLayoutServices(doc);
        const layout = layoutDocument(doc, layoutServices, { currentDateMs: 0 });
        expect(layout.pages).toHaveLength(2);
        expect(layout.pages.map((page: Any) => ({
          widthPt: page.geometry.widthPt,
          heightPt: page.geometry.heightPt,
          direction: page.section.textDirection,
        }))).toEqual([
          { widthPt: 612, heightPt: 792, direction: 'btLr' },
          { widthPt: 612, heightPt: 792, direction: 'lrTb' },
        ]);

        const renderPage = async (pageIndex: number) => {
          const runs: RunInfo[] = [];
          const canvas = new Canvas(10, 10);
          await renderDocumentToCanvas(doc, canvas as Any, pageIndex, {
            dpr: 1,
            width: 612,
            layoutServices,
            currentDate: 0,
            defaultCurrentDateMs: 0,
            onTextRun: (run: RunInfo) => runs.push(run),
          });
          return { runs, canvas };
        };

        const vertical = await renderPage(0);
        expect(vertical.canvas.width).toBe(612);
        expect(vertical.canvas.height).toBe(792);
        expect(vertical.runs.length).toBeGreaterThan(0);
        expect(vertical.runs.every((run) => run.transform !== undefined)).toBe(true);

        const verticalXs = vertical.runs.map((run) => run.x);
        expect(Math.max(...verticalXs)).toBeGreaterThan(500);
        expect(Math.max(...verticalXs)).toBeLessThanOrEqual(540);
        expect(Math.min(...verticalXs)).toBeLessThan(Math.max(...verticalXs));
        expect(Math.min(...vertical.runs.map((run) => run.y))).toBeCloseTo(72, 0);

        const horizontal = await renderPage(1);
        expect(horizontal.canvas.width).toBe(612);
        expect(horizontal.canvas.height).toBe(792);
        expect(horizontal.runs.length).toBeGreaterThan(0);
        expect(horizontal.runs.every((run) => run.transform === undefined)).toBe(true);
        expect(Math.min(...horizontal.runs.map((run) => run.x))).toBeCloseTo(72, 0);
        const minY = Math.min(...horizontal.runs.map((run) => run.y));
        expect(minY).toBeGreaterThan(70);
        expect(minY).toBeLessThan(95);
      } finally {
        restoreOffscreen();
        restoreImage();
      }
    });
  },
);
