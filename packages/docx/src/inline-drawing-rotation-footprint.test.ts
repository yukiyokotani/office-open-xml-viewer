/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { hitTestDocxElementContext } from './element-context.js';
import { createPaintResourceRegistry } from './layout/paint-resources.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const EMU = 12700;
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
  + 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

type Kind = 'pic' | 'wps' | 'wpg' | 'vml';
type Case = Readonly<{ kind: Kind; degrees: number; flipH?: boolean; flipV?: boolean; effectPt?: readonly [number, number, number, number] }>;

const xfrm = (degrees: number, flipH = false, flipV = false, off = [0, 0], ext = [200, 100], group = false) =>
  `<a:xfrm${degrees ? ` rot="${Math.round(degrees * 60000)}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}>`
  + `<a:off x="${off[0]! * EMU}" y="${off[1]! * EMU}"/><a:ext cx="${ext[0]! * EMU}" cy="${ext[1]! * EMU}"/>`
  + (group ? `<a:chOff x="0" y="0"/><a:chExt cx="${ext[0]! * EMU}" cy="${ext[1]! * EMU}"/>` : '') + '</a:xfrm>';
const picture = (degrees: number, flipH = false, flipV = false, off = [0, 0], ext = [200, 100]) =>
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="P"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + `<pic:spPr>${xfrm(degrees, flipH, flipV, off, ext)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`;
const shape = (degrees: number, off = [0, 0], ext = [200, 100]) =>
  `<wps:wsp><wps:cNvSpPr/><wps:spPr>${xfrm(degrees, false, false, off, ext)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  + '<a:solidFill><a:srgbClr val="9BD18B"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>';
const group = (degrees: number) => `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm(degrees, false, false, [0, 0], [200, 100], true)}</wpg:grpSpPr>`
  + `${shape(0, [0, 0], [100, 100])}${picture(0, false, false, [100, 0], [100, 100])}</wpg:wgp>`;
const URI: Record<Exclude<Kind, 'vml'>, string> = {
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
};

function docx(test: Case): Uint8Array {
  const body = test.kind === 'pic' ? picture(test.degrees, test.flipH, test.flipV)
    : test.kind === 'wps' ? shape(test.degrees) : test.kind === 'wpg' ? group(test.degrees) : '';
  const [l, t, r, b] = (test.effectPt ?? [0, 0, 0, 0]).map((edge) => Math.round(edge * EMU));
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const drawing = test.kind === 'vml'
    // A legacy inline VML picture: the rotation is its CSS rotation, not a:xfrm.
    ? `<w:r><w:pict><v:shape id="P" type="#_x0000_t75" style="width:200pt;height:100pt;rotation:${test.degrees}">`
      + '<v:imagedata r:id="rIdImg" o:title=""/></v:shape></w:pict></w:r>'
    : `<w:r><w:drawing><wp:inline><wp:extent cx="${200 * EMU}" cy="${100 * EMU}"/>`
      + `<wp:effectExtent l="${l}" t="${t}" r="${r}" b="${b}"/><wp:docPr id="1" name="O"/>`
      + `<a:graphic><a:graphicData uri="${URI[test.kind]}">${body}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const paragraph = (content: string) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${content}</w:p>`;
  const document = `<w:document ${NS}><w:body>${paragraph(run('B') + drawing + run('A'))}${paragraph(run('N'))}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdImg" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/></Relationships>')],
    ['word/document.xml', encode(document)],
    ['word/media/p.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47])],
  ]));
}

function layout(test: Case) {
  const archive = new DocxArchive(docx(test));
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  const page = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 }).pages[0]!;
  const [first, second] = page.layers.body;
  if (first?.kind !== 'paragraph' || second?.kind !== 'paragraph') throw new Error('expected two paragraphs');
  const line = first.lines[0]!;
  const tail = line.placements.filter((placement) => placement.kind === 'text')[1];
  const resource = line.placements.find((placement) => placement.kind === 'resource');
  const command = first.drawings[0]?.commands[0];
  const object = resource?.kind === 'resource' ? resource.bounds
    : command?.kind === 'drawingml-shape' ? { xPt: command.plan.rect.x, yPt: command.plan.rect.y, widthPt: command.plan.rect.w, heightPt: command.plan.rect.h }
      : undefined;
  if (!tail || !object) throw new Error('expected the object and the text after it');
  const hit = (xPt: number, yPt: number) => resource?.kind === 'resource' ? hitTestDocxElementContext(
    { pages: [page] } as never, 0, { xPt, yPt }, createPaintResourceRegistry([{
      kind: 'image', resourceKey: resource.resourceKey, partPath: 'word/media/p.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 200, heightPt: 100 },
    }]),
  ) : undefined;
  return { tailX: tail.bounds.xPt, nextTop: second.lines[0]!.bounds.yPt, object, first, hit,
    lineClip: resource?.kind === 'resource' ? resource.lineClip : undefined,
    rotationDeg: resource?.kind === 'resource' ? resource.rotationDeg : undefined };
}

// Compatibility rule `word-inline-drawing-turned-footprint`. Expected values
// follow the Windows Word 16.0.20527 exports of the same controls: the base box
// is turned for [45, 135) and [225, 315) degrees, effectExtent widens it on its
// page edges, and the unturned 200 x 100pt object is centred in it. The text
// before the object is 5pt wide at the 72pt margin; a line adds 2pt descent.
describe('inline drawing turned by its own xfrm (word-inline-drawing-turned-footprint, word-inline-rotated-picture-line-clip)', () => {
  it.each<Case & { turned: boolean }>([
    { kind: 'pic', degrees: 0, turned: false },
    { kind: 'pic', degrees: 30, turned: false },
    { kind: 'pic', degrees: 30, effectPt: [11.6, 43.3, 11.6, 43.3], turned: false },
    { kind: 'pic', degrees: 45, turned: true },
    { kind: 'pic', degrees: 60, effectPt: [-6.7, 61.6, -6.7, 61.6], turned: true },
    { kind: 'pic', degrees: 90, turned: true },
    { kind: 'pic', degrees: 90, flipH: true, flipV: true, turned: true },
    { kind: 'pic', degrees: 90, effectPt: [10, 20, 30, 40], turned: true },
    { kind: 'pic', degrees: 0, effectPt: [10, 20, 30, 40], turned: false },
    { kind: 'pic', degrees: 135, turned: false },
    { kind: 'pic', degrees: 180, turned: false },
    { kind: 'pic', degrees: 270, turned: true },
    { kind: 'pic', degrees: -90, turned: true },
    { kind: 'wps', degrees: 0, turned: false },
    { kind: 'wps', degrees: 30, turned: false },
    { kind: 'wps', degrees: 45, turned: true },
    { kind: 'wps', degrees: 90, turned: true },
    { kind: 'wpg', degrees: 0, turned: false },
    { kind: 'wpg', degrees: 45, turned: true },
    { kind: 'wpg', degrees: 90, turned: true },
  ])('lays out a $kind at $degrees degrees (flipH $flipH, effect $effectPt) in its Word base box', (test) => {
    const [l, t, r, b] = test.effectPt ?? [0, 0, 0, 0];
    const [baseW, baseH] = test.turned ? [100, 200] : [200, 100];
    const { tailX, nextTop, object, lineClip } = layout(test);
    expect(tailX).toBeCloseTo(77 + l + baseW + r, 6);
    expect(nextTop).toBeCloseTo(72 + t + baseH + b + 2, 6);
    expect([object.widthPt, object.heightPt]).toEqual([200, 100]);
    expect(object.xPt + object.widthPt / 2).toBeCloseTo(77 + l + baseW / 2, 6);
    expect(object.yPt + object.heightPt / 2).toBeCloseTo(72 + t + baseH / 2, 6);
    // word-inline-rotated-picture-line-clip: a turned picture paints inside
    // its line box (72pt to the next line's top); shapes and groups do not.
    if (test.kind === 'pic' && test.degrees % 360 !== 0) {
      expect(lineClip?.topPt).toBeCloseTo(72 - object.yPt, 6);
      expect(lineClip?.bottomPt).toBeCloseTo(nextTop - object.yPt, 6);
    } else {
      expect(lineClip).toBeUndefined();
    }
  });

  it('moves the children of a turned group with its base box', () => {
    // Word paints the group's 100 x 100pt rectangle child (group space 0..100)
    // over the top half of the turned 100 x 200pt base box.
    const { first } = layout({ kind: 'wpg', degrees: 90 });
    const rects: { x: number; y: number; w: number; h: number }[] = [];
    const visit = (value: unknown): void => {
      if (Array.isArray(value)) { value.forEach(visit); return; }
      if (value === null || typeof value !== 'object') return;
      const node = value as { kind?: string; plan?: { rect?: { x: number; y: number; w: number; h: number }; geometry?: { kind?: string } } };
      if (node.kind === 'drawingml-shape' && node.plan?.geometry?.kind === 'preset' && node.plan.rect) rects.push(node.plan.rect);
      Object.values(value).forEach(visit);
    };
    visit(first.textBoxes);
    expect(rects).toHaveLength(1);
    expect(rects[0]!.x + rects[0]!.w / 2).toBeCloseTo(77 + 50, 6);
    expect(rects[0]!.y + rects[0]!.h / 2).toBeCloseTo(72 + 50, 6);
  });

  it('hit-tests and bounds a turned picture by its visible outline', () => {
    // 90 degrees: the picture paints over x 77..177, y 72..272.
    const turned = layout({ kind: 'pic', degrees: 90 });
    expect(turned.hit(100, 90)).toMatchObject({ elementType: 'image', bounds: { xPt: 77, yPt: 72, widthPt: 100, heightPt: 200 } });
    expect(turned.hit(40, 150)).toBeNull();
    // 30 degrees, effectExtent 0: the turned picture (centre 177, 122) is cut
    // at the 72..174pt line box.
    const cut = layout({ kind: 'pic', degrees: 30 });
    const context = cut.hit(177, 100);
    expect(context).toMatchObject({ elementType: 'image' });
    expect(context!.bounds.yPt).toBeCloseTo(72, 6);
    expect(context!.bounds.yPt + context!.bounds.heightPt).toBeCloseTo(174, 6);
    expect(cut.hit(177, 68)).toBeNull();
    expect(cut.hit(177, 178)).toBeNull();
    // Inside the authored box but outside the turned picture.
    expect(cut.hit(80, 75)).toBeNull();
  });

  it('leaves a rotated inline VML picture on its authored box, unclipped', () => {
    // Both rules are DrawingML observations; VML pictures stay unchanged.
    const { tailX, nextTop, object, lineClip, rotationDeg } = layout({ kind: 'vml', degrees: 90 });
    expect(rotationDeg).toBe(90);
    expect(tailX).toBeCloseTo(77 + 200, 6);
    expect(nextTop).toBeCloseTo(72 + 100 + 2, 6);
    expect(object).toMatchObject({ xPt: 77, yPt: 72, widthPt: 200, heightPt: 100 });
    expect(lineClip).toBeUndefined();
  });
});
