import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxDocumentModel } from './types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function document(indentPt: number): DocxDocumentModel {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const wpg = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const textbox = (text: string, x: number) => `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>
    <a:xfrm><a:off x="${x * 12700}" y="0"/><a:ext cx="254000" cy="254000"/></a:xfrm>
    <a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>
    <wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
    <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>`;
  const files = new Map<string, Uint8Array>();
  const xml = (name: string, value: string) => files.set(name, new TextEncoder().encode(value));
  xml('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}" xmlns:a="${a}"
    xmlns:wp="${wp}" xmlns:wpg="${wpg}" xmlns:wps="${wps}" xmlns:pic="${pic}"><w:body>
    <w:p><w:pPr><w:ind w:left="${indentPt * 20}"/></w:pPr><w:r><w:t>PRE </w:t></w:r>
    <w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="381000"/><wp:docPr id="1" name="group"/>
    <a:graphic><a:graphicData uri="${wpg}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>
    <a:off x="0" y="0"/><a:ext cx="1270000" cy="381000"/><a:chOff x="0" y="0"/>
    <a:chExt cx="1270000" cy="381000"/></a:xfrm></wpg:grpSpPr>${textbox('A', 0)}
    <pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="pixel"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="img"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="762000" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>
    <a:prstGeom prst="rect"/></pic:spPr></pic:pic>${textbox('B', 80)}
    </wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>POST</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr></w:body></w:document>`);
  xml('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  xml('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="img" Type="${r}/image" Target="media/pixel.png"/></Relationships>`);
  xml('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  files.set('word/media/pixel.png', Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64')));
  const archive = new DocxArchive(storeZip(files));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

function canvas() {
  let matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const stack: typeof matrix[] = [];
  const events: { text: string; x: number; y: number }[] = [];
  const record = (text: string, x: number, y: number) => events.push({ text,
    x: matrix.a * x + matrix.c * y + matrix.e, y: matrix.b * x + matrix.d * y + matrix.f });
  const ctx = {
    font: '11px serif', letterSpacing: '0px', fontKerning: 'auto', globalAlpha: 1,
    measureText(text: string) { return { width: text.length * 5, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics; },
    save() { stack.push({ ...matrix }); }, restore() { matrix = stack.pop()!; },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      matrix = { a: matrix.a * a + matrix.c * b, b: matrix.b * a + matrix.d * b,
        c: matrix.a * c + matrix.c * d, d: matrix.b * c + matrix.d * d,
        e: matrix.a * e + matrix.c * f + matrix.e, f: matrix.b * e + matrix.d * f + matrix.f };
    },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) { matrix = { a, b, c, d, e, f }; },
    getTransform() { return matrix; },
    translate(x: number, y: number) { ctx.transform(1, 0, 0, 1, x, y); },
    scale(x: number, y: number) { ctx.transform(x, 0, 0, y, 0, 0); },
    rotate(angle: number) { ctx.transform(Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0); },
    fillText: record, strokeText() {}, clearRect() {}, fillRect() {}, strokeRect() {},
    beginPath() {}, closePath() {}, rect() {}, clip() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, setLineDash() {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, events, record,
    target: { getContext: () => ctx, width: 0, height: 0 } as unknown as HTMLCanvasElement };
}

it('retains inline group flow, interleaved paint order, text geometry and host-relative placement', async () => {
  const captures = [];
  for (const indentPt of [0, 72]) {
    const doc = document(indentPt);
    const surface = canvas();
    const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: surface.ctx }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    const before = JSON.stringify(layout);
    // Exercise serialization of retained layout; actual worker acquisition is a browser check.
    surface.ctx.measureText = () => { throw new Error('paint must not measure'); };
    await paintLayoutPage(structuredClone(layout), 0, surface.target, { scale: 1, dpr: 1 }, {
      paint(_key, _kind, rect) { surface.record('IMAGE', rect.xPt, rect.yPt); },
    });
    expect(JSON.stringify(layout)).toBe(before);
    const events = surface.events.filter(event => ['A', 'B', 'IMAGE'].includes(event.text));
    expect(events.map(event => event.text)).toEqual(['A', 'IMAGE', 'B']);
    expect(events[1]!.x - events[0]!.x).toBeCloseTo(60);
    expect(events[2]!.x - events[0]!.x).toBeCloseTo(80);
    const overlay = textRunsForPage(layout, 0, { scale: 1 });
    expect(overlay.map(run => run.text).join('')).toBe('PRE ABPOST');
    const geometry = textRunGeometryForPage(layout, 0).filter(run => ['A', 'B'].includes(run.placement.text));
    expect(geometry).toHaveLength(2);
    captures.push(events);
  }
  for (let i = 0; i < 3; i++) expect(captures[1]![i]!.x - captures[0]![i]!.x).toBeCloseTo(72);
});
