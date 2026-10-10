/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';

const encoder = new TextEncoder();
const EMU = 12_700;
const PNG = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==',
  'base64',
));

function shape(color: string): string {
  return `<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${36 * EMU}" cy="${20 * EMU}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr></wps:wsp>`;
}

function inlineGroup(): string {
  return `<w:r><w:drawing><wp:inline><wp:extent cx="${36 * EMU}" cy="${20 * EMU}"/><wp:docPr id="1" name="inline group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${36 * EMU}" cy="${20 * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${36 * EMU}" cy="${20 * EMU}"/></a:xfrm></wpg:grpSpPr>${shape('0000FF')}<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="middle.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${36 * EMU}" cy="${20 * EMU}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>${shape('00FF00')}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function docx(): Uint8Array {
  const document = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p><w:r><w:t>BEFORE</w:t></w:r>${inlineGroup()}<w:r><w:t>AFTER</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="15840" w:h="12240"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`;
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/one.png"/></Relationships>')],
    ['word/document.xml', encoder.encode(document)],
    ['word/media/one.png', PNG],
  ]));
}

function recordingCanvas(): {
  canvas: HTMLCanvasElement;
  events: string[];
  fillText: Array<{ text: string; x: number }>;
} {
  let font = '10px serif';
  let fillStyle = '#000000';
  const events: string[] = [];
  const fillText: Array<{ text: string; x: number }> = [];
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    get fillStyle() { return fillStyle; }, set fillStyle(value: string) { fillStyle = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const px = Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 10);
      return { width: [...text].length * px * 0.5, fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2, actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2 } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() { if (String(fillStyle).startsWith('rgba(')) events.push(String(fillStyle)); },
    fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() { events.push('image'); },
    fillText(text: string, x: number) { fillText.push({ text, x }); }, strokeText() {},
    strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection, globalAlpha: 1, lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    canvas: { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement,
    events, fillText,
  };
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

afterEach(() => vi.unstubAllGlobals());

describe('inline WPG group through DOCX parse and render', () => {
  it('keeps mixed child paint order and advances paragraph once by wp:extent', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 1, height: 1, close() {} })));
    const archive = new DocxArchive(docx());
    let model;
    try {
      model = normalizeInternalDocumentModel(
        JSON.parse(new TextDecoder().decode(archive.parse())),
      ).document;
    } finally {
      archive.free();
    }
    const paint = recordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(model, paint.canvas, 0, {
      width: 792,
      dpr: 1,
      fetchImage: async (_path, mime) => new Blob([PNG], { type: mime }),
      onTextRun: (run) => runs.push(run),
    });

    expect(paint.events).toEqual(['rgba(0,0,255,1)', 'image', 'rgba(0,255,0,1)']);
    const before = paint.fillText.find((run) => run.text === 'BEFORE');
    const after = paint.fillText.find((run) => run.text === 'AFTER');
    const beforeRun = runs.find((run) => run.text === 'BEFORE');
    if (!before || !after || !beforeRun) throw new Error('rendered text callbacks missing');
    expect(after.x - (before.x + beforeRun.w)).toBeCloseTo(36, 2);
  });
});
