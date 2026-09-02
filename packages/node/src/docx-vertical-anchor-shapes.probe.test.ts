/**
 * positionV/positionH anchors under a vertical (tbRl) section — ECMA-376
 * §17.6.20 + §20.4.3.x, issue #988 batch-3 adjudication ②.
 *
 * The synthetic package below encodes the three anchor reference frames from
 * the Office-observed compatibility case. It traverses the real WASM parser
 * and renderer while remaining deterministic and independent of private sample
 * filenames.
 *
 *   | fill    | positionH (page) | positionV        | physical box                |
 *   |---------|------------------|------------------|-----------------------------|
 *   | #FCE4D6 | 72.0             | paragraph + 21.6 | (72.0, 93.6)–(172.8, 180.0) |
 *   | #E2EFDA | 230.4            | margin + 108     | (230.4, 180.0)–(331.2,266.4)|
 *   | #DDEBF7 | 388.8            | page + 216       | (388.8, 216.0)–(489.6,302.4)|
 */
import { describe, expect, it } from 'vitest';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { minimalDocx } from './test-ooxml-package';
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

interface Box { x0: number; y0: number; x1: number; y1: number }

const EXPECTED: Array<{ name: string; rgb: [number, number, number]; box: Box }> = [
  { name: 'V=paragraph', rgb: [0xfc, 0xe4, 0xd6], box: { x0: 72.0, y0: 93.6, x1: 172.8, y1: 180.0 } },
  { name: 'V=margin', rgb: [0xe2, 0xef, 0xda], box: { x0: 230.4, y0: 180.0, x1: 331.2, y1: 266.4 } },
  { name: 'V=page', rgb: [0xdd, 0xeb, 0xf7], box: { x0: 388.8, y0: 216.0, x1: 489.6, y1: 302.4 } },
];

function anchoredRectangle(
  id: number,
  horizontalOffset: number,
  verticalReference: 'paragraph' | 'margin' | 'page',
  verticalOffset: number,
  fill: string,
): string {
  return `<w:r><w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0"
               simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0"
               layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>${horizontalOffset}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="${verticalReference}"><wp:posOffset>${verticalOffset}</wp:posOffset></wp:positionV>
      <wp:extent cx="1280160" cy="1097280"/>
      <wp:wrapNone/>
      <wp:docPr id="${id}" name="Anchor ${id}"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1280160" cy="1097280"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>
              <a:ln><a:noFill/></a:ln>
            </wps:spPr>
            <wps:bodyPr vert="horz"/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing></w:r>`;
}

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      ${anchoredRectangle(1, 914400, 'paragraph', 274320, 'FCE4D6')}
      ${anchoredRectangle(2, 2926080, 'margin', 1371600, 'E2EFDA')}
      ${anchoredRectangle(3, 4937760, 'page', 2743200, 'DDEBF7')}
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:textDirection w:val="tbRl"/>
    </w:sectPr>
  </w:body>
</w:document>`;

describe.skipIf(!skia || !docxMod || !rendererMod)(
  'docx vertical anchored shapes resolve physically (§20.4.3.x, #988 ②)',
  () => {
    it('lands all positionV reference frames on their physical boxes', async () => {
      const { materializeDocxDocument } = docxMod as {
        materializeDocxDocument: (bytes: Uint8Array) => Promise<Any>;
      };
      const { renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(minimalDocx(DOCUMENT_XML));
      const canvas = new Canvas(612, 792);
      const restoreImage = installImageBitmapShim(factory);
      const restoreOffscreen = installOffscreenCanvasShim(factory);
      try {
        await renderDocumentToCanvas(doc, canvas as Any, 0, {
          dpr: 1,
          width: 612,
          currentDate: 0,
          defaultCurrentDateMs: 0,
        });
      } finally {
        restoreOffscreen();
        restoreImage();
      }

      expect(canvas.width).toBe(612);
      expect(canvas.height).toBe(792);
      const context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
      const { data } = context.getImageData(0, 0, canvas.width, canvas.height);

      for (const { name, rgb, box } of EXPECTED) {
        let x0 = Infinity;
        let y0 = Infinity;
        let x1 = -Infinity;
        let y1 = -Infinity;
        for (let y = 0; y < canvas.height; y += 1) {
          for (let x = 0; x < canvas.width; x += 1) {
            const index = (y * canvas.width + x) * 4;
            if (
              Math.abs(data[index] - rgb[0]) <= 4
              && Math.abs(data[index + 1] - rgb[1]) <= 4
              && Math.abs(data[index + 2] - rgb[2]) <= 4
            ) {
              x0 = Math.min(x0, x);
              y0 = Math.min(y0, y);
              x1 = Math.max(x1, x);
              y1 = Math.max(y1, y);
            }
          }
        }

        expect(x0, `${name} fill found`).toBeLessThan(Infinity);
        const tolerance = 2.5;
        expect(Math.abs(x0 - box.x0), `${name} left`).toBeLessThanOrEqual(tolerance);
        expect(Math.abs(y0 - box.y0), `${name} top`).toBeLessThanOrEqual(tolerance);
        expect(Math.abs(x1 - box.x1), `${name} right`).toBeLessThanOrEqual(tolerance);
        expect(Math.abs(y1 - box.y1), `${name} bottom`).toBeLessThanOrEqual(tolerance);
      }
    }, 120_000);
  },
);
