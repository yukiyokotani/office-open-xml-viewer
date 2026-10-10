import { describe, expect, it } from 'vitest';
import { Canvas } from 'skia-canvas';
import { openDocxDocument } from './docx.ts';
import type { NodeCanvasFactory } from './render.ts';
import { minimalDocx } from './test-ooxml-package';

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: async () => { throw new Error('Unexpected image in line-only fixture'); },
};

const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

function line(id: number, zeroWidth: boolean): string {
  const cx = zeroWidth ? 0 : 1_270_000;
  const cy = zeroWidth ? 254_000 : 0;
  const x0 = zeroWidth ? 10_800 : 0;
  const y0 = zeroWidth ? 0 : 10_800;
  const x1 = zeroWidth ? 10_800 : 21_600;
  const y1 = zeroWidth ? 21_600 : 10_800;
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="1270000" cy="254000"/><wp:docPr id="${id}" name="zero-axis-${id}"/>
    <a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
      <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>
        <a:pathLst><a:path w="21600" h="21600"><a:moveTo><a:pt x="${x0}" y="${y0}"/></a:moveTo>
          <a:lnTo><a:pt x="${x1}" y="${y1}"/></a:lnTo></a:path></a:pathLst>
      </a:custGeom><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
    </wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r></w:p>`;
}

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="${WPS}"><w:body>${line(1, true)}${line(2, false)}
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"
    w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

describe('inline WPS custom paths with zero transform axes', () => {
  it('paints stroked paths when either transform extent is zero', async () => {
    const session = await openDocxDocument(minimalDocx(DOCUMENT), { factory });
    const rendered = await session.renderPage(0, { dpr: 1, width: 612 })
      .finally(() => session.close());
    const canvas = rendered as unknown as Canvas;

    const { data } = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    const rows = new Uint16Array(canvas.height);
    const columns = new Uint16Array(canvas.width);
    for (let y = 0; y < canvas.height; y += 1) {
      for (let x = 0; x < canvas.width; x += 1) {
        const index = (y * canvas.width + x) * 4;
        if (data[index] < 80 && data[index + 1] < 80 && data[index + 2] < 80) {
          rows[y] += 1;
          columns[x] += 1;
        }
      }
    }
    expect(Math.max(...rows), 'horizontal custom stroke pixels in one row').toBeGreaterThan(80);
    expect(Math.max(...columns), 'vertical custom stroke pixels in one column').toBeGreaterThan(15);
  }, 120_000);
});
