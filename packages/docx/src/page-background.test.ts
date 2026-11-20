import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { paintLayoutPageContent } from './paint/canvas-page.js';
import type { DocumentLayout, LayoutRect, PaintResourceKind } from './layout/types.js';
import type { CanvasPaintResourcePainter } from './paint/types.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';
import type { DocxDocumentModel } from './types.js';

// ECMA-376 Part 1 §17.2.1 `w:background` through the production parser,
// layout and canvas paint, gated by §17.15.1.26 `w:displayBackgroundShape`.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="t">${
  [['dk1', '000000'], ['lt1', 'FFFFFF'], ['dk2', '1F497D'], ['lt2', 'EEECE1'], ['accent1', '4F81BD'],
    ['accent2', 'C0504D'], ['accent3', '9BBB59'], ['accent4', '8064A2'], ['accent5', '4BACC6'],
    ['accent6', 'F79646'], ['hlink', '0000FF'], ['folHlink', '800080']]
    .map(([slot, hex]) => `<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`).join('')
}</a:clrScheme></a:themeElements></a:theme>`;

interface Fixture {
  readonly background: string;
  readonly settings?: string;
  /** Body content before the final section properties. */
  readonly body?: string;
  readonly header?: boolean;
}

function parse({ background, settings = '<w:displayBackgroundShape/>', body, header = false }: Fixture): DocxDocumentModel {
  const paragraph = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>';
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">${background}<w:body>${body ?? paragraph}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="rHeader"/>' : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
  const relationships = [
    ['rSettings', 'settings', 'settings.xml'],
    ['rTheme', 'theme', 'theme/theme1.xml'],
    ['rTexture', 'image', 'media/texture.png'],
    ['rChart', 'chart', 'charts/chart1.xml'],
    ...(header ? [['rHeader', 'header', 'header1.xml']] : []),
  ];
  const files = new Map<string, string | Uint8Array>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${
      relationships.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')
    }</Relationships>`],
    ['word/document.xml', document],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`],
    ['word/theme/theme1.xml', THEME],
    ['word/media/texture.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47])],
    ['word/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, data]) => [
    name, typeof data === 'string' ? new TextEncoder().encode(data) : data,
  ]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

function layout(model: DocxDocumentModel): DocumentLayout {
  const ctx = makeRecordingCanvas().canvas.getContext('2d') as CanvasRenderingContext2D;
  return layoutDocument(model, createLayoutServices(model, { measureContext: ctx }), { currentDateMs: 0 });
}

/** A Canvas 2D stub that records the paint sequence: filled rectangles and
 * paths with their fill style, and text. */
function recordingContext(): { ctx: CanvasRenderingContext2D; events: string[] } {
  const events: string[] = [];
  const state: Record<string | symbol, unknown> = { fillStyle: '#000000', globalAlpha: 1 };
  const style = () => {
    const fill = state.fillStyle;
    return typeof fill === 'string' ? fill.toLowerCase() : `gradient:${(fill as { stops?: string[] }).stops?.join(',')}`;
  };
  const gradient = () => {
    const stops: string[] = [];
    return { stops, addColorStop: (_: number, color: string) => { stops.push(color); } };
  };
  const ctx = new Proxy(state, {
    get(target, key) {
      if (key === 'fillRect') return () => { events.push(`rect:${style()}`); };
      if (key === 'fill') return () => { events.push(`fill:${style()}`); };
      if (key === 'fillText') return (text: string) => { events.push(`text:${text}`); };
      if (key === 'measureText') return (text: string) => ({ width: text.length * 5 });
      if (key === 'createLinearGradient' || key === 'createRadialGradient') return gradient;
      if (key === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (key in target) return target[key];
      return () => undefined;
    },
    set(target, key, value) { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, events };
}

function paint(result: DocumentLayout, pageIndex = 0, resources?: CanvasPaintResourcePainter) {
  const { ctx, events } = recordingContext();
  paintLayoutPageContent(result.pages[pageIndex]!, {
    ctx, scale: 1, dpr: 1,
    resources: resources ?? { paint: () => { throw new Error('unexpected resource'); } },
  });
  return events;
}

const THEMED = '<w:background w:color="000000" w:themeColor="accent1" w:themeTint="99"/>';

describe('§17.2.1 document background', () => {
  it.each([
    ['on', '<w:displayBackgroundShape/>', '#95B3D7'],
    ['explicit true', '<w:displayBackgroundShape w:val="true"/>', '#95B3D7'],
    ['absent', '', null],
    ['val=false', '<w:displayBackgroundShape w:val="false"/>', null],
    ['val=0', '<w:displayBackgroundShape w:val="0"/>', null],
  ])('displayBackgroundShape %s decides whether the theme colour is painted', (_label, settings, expected) => {
    const result = layout(parse({ background: THEMED, settings }));
    const page = result.pages[0]!;
    if (expected === null) {
      expect(page.pageBackground).toBeUndefined();
      expect(paint(result).some((event) => event.startsWith('rect:'))).toBe(false);
      return;
    }
    // themeColor accent1 (4F81BD) with themeTint 0x99 overrides w:color.
    expect(page.pageBackground?.commands).toEqual([{
      kind: 'fill-rect',
      rect: { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 },
      fill: expected,
    }]);
  });

  it('resolves background1 through the settings colour scheme mapping', () => {
    const page = layout(parse({
      background: '<w:background w:color="FFFFFF" w:themeColor="background1"/>',
      settings: '<w:displayBackgroundShape/><w:clrSchemeMapping w:bg1="dark2"/>',
    })).pages[0]!;
    expect(page.pageBackground?.commands[0]).toMatchObject({ kind: 'fill-rect', fill: '#1F497D' });
  });

  it('paints the background before the header, behind-text drawings and body text', () => {
    const behindText = '<w:p><w:r><w:pict><v:rect style="position:absolute;margin-left:0;margin-top:0;width:50pt;height:50pt;z-index:-1" fillcolor="#ff0000" stroked="f"/></w:pict></w:r><w:r><w:t>Body</w:t></w:r></w:p>';
    const result = layout(parse({
      background: '<w:background w:color="22551B"/>', body: behindText, header: true,
    }));
    const events = paint(result);
    // The full-page colour precedes every story and anchored drawing.
    expect(events).toEqual(['rect:#22551b', 'text:Header', 'fill:rgba(255,0,0,1)', 'text:Body']);
  });

  it('layers a VML gradient above the solid colour on every page', () => {
    const background = '<w:background w:color="FFF0E1"><v:background fillcolor="#fff0e1 [3201]"><v:fill type="gradient" color2="#4f81bd [3204]"/></v:background></w:background>';
    const twoPages = '<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/><w:t>Two</w:t></w:r></w:p>';
    const result = layout(parse({ background, body: twoPages }));
    expect(result.pages).toHaveLength(2);
    for (const pageIndex of [0, 1]) {
      const commands = result.pages[pageIndex]!.pageBackground!.commands;
      expect(commands.map((command) => command.kind)).toEqual(['fill-rect', 'drawingml-shape']);
      const shape = commands[1];
      if (shape?.kind !== 'drawingml-shape') throw new Error('gradient command');
      expect(shape.plan.rect).toEqual({ x: 0, y: 0, w: 612, h: 792 });
      expect(shape.plan.fill).toMatchObject({ fillType: 'gradient', stops: [{ color: 'fff0e1' }, { color: '4f81bd' }] });
      const events = paint(result, pageIndex);
      expect(events[0]).toBe('rect:#fff0e1');
      expect(events[1]).toMatch(/^fill:gradient:/u);
      expect(events.findIndex((event) => event.startsWith('text:'))).toBeGreaterThan(1);
    }
  });

  it('retains one page-sized VML picture brush per physical page size', () => {
    const background = '<w:background><v:background><v:fill r:id="rTexture" type="tile"/></v:background></w:background>';
    // The first section is landscape: a second physical page size.
    const body = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:pPr><w:r><w:t>Wide</w:t></w:r></w:p><w:p><w:r><w:t>Tall</w:t></w:r></w:p>';
    const model = parse({ background, body });
    const result = layout(model);
    expect(result.pages.map((page) => [page.geometry.widthPt, page.geometry.heightPt])).toEqual([[792, 612], [612, 792]]);
    const keys = result.pages.map((page) => {
      const command = page.pageBackground?.commands[0];
      if (command?.kind !== 'drawingml-image-fill') throw new Error('picture brush');
      expect(command.plan.rect).toEqual({ x: 0, y: 0, w: page.geometry.widthPt, h: page.geometry.heightPt });
      return command.resourceKey;
    });
    expect(new Set(keys).size).toBe(2);
    const descriptors = layoutSourceStore(model).paintResources.descriptors;
    for (const [index, key] of keys.entries()) {
      const descriptor = descriptors.find((candidate) => candidate.resourceKey === key);
      expect(descriptor).toMatchObject({
        kind: 'image',
        partPath: 'word/media/texture.png',
        intrinsicSize: {
          widthPt: result.pages[index]!.geometry.widthPt,
          heightPt: result.pages[index]!.geometry.heightPt,
        },
        vmlFill: { repeat: true },
      });
    }
    const painted: Array<[string, PaintResourceKind, LayoutRect]> = [];
    const events = paint(result, 1, {
      paint: (key: string, kind: PaintResourceKind, bounds: LayoutRect) => { painted.push([key, kind, bounds]); },
    });
    expect(painted).toEqual([[keys[1], 'image', { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 }]]);
    expect(events.findIndex((event) => event.startsWith('text:'))).toBeGreaterThanOrEqual(0);
  });

  const inline = (graphic: string, cx = 914400, cy = 457200) => `<w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="9" name="bg"/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing>`;
  const rect = (x: number, fill: string) => `<wps:wsp><wps:cNvPr id="${x + 10}" name="r"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}<a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>`;

  it('paints a DrawingML group background over the page box beneath every story', () => {
    const group = inline(`<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm></wpg:grpSpPr>${
      rect(0, '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')}${
      rect(457200, '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>')}</wpg:wgp></a:graphicData>`);
    const result = layout(parse({
      background: `<w:background w:color="22551B">${group}</w:background>`, header: true,
    }));
    const page = result.pages[0]!;
    const drawings = page.layers.background.flatMap((node) => (
      node.kind === 'paragraph' ? node.drawings : []));
    // The group occupies the page box; each member scales from the authored
    // 72x36pt group frame.
    expect(drawings.map((drawing) => drawing.flowBounds)).toEqual([
      { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 },
    ]);
    expect(drawings[0]!.commands.map((command) => (
      command.kind === 'drawingml-shape' ? command.plan.rect : command.kind))).toEqual([
      { x: 0, y: 0, w: 306, h: 792 },
      { x: 306, y: 0, w: 306, h: 792 },
    ]);
    expect(result.diagnostics).toEqual([]);
    expect(paint(result)).toEqual([
      'rect:#22551b', 'fill:rgba(255,0,0,1)', 'fill:rgba(0,0,255,1)', 'text:Header', 'text:Body',
    ]);
  });

  it('paints a chart background through the retained chart resource at page size', () => {
    const chart = inline('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData>');
    const model = parse({ background: `<w:background>${chart}</w:background>` });
    const result = layout(model);
    const painted: Array<[PaintResourceKind, LayoutRect]> = [];
    paint(result, 0, { paint: (_key: string, kind: PaintResourceKind, bounds: LayoutRect) => { painted.push([kind, bounds]); } });
    expect(painted).toEqual([['chart', { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 }]]);
    expect(result.pages[0]!.pageBackground).toBeUndefined();
  });

  it('tiles a DrawingML picture fill of a background shape', () => {
    const tiled = inline(`<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${
      rect(0, '<a:blipFill><a:blip r:embed="rTexture"/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="x" algn="tl"/></a:blipFill>')
    }</a:graphicData>`, 457200, 457200);
    const model = parse({ background: `<w:background>${tiled}</w:background>` });
    const result = layout(model);
    const drawing = result.pages[0]!.layers.background.flatMap((node) => (
      node.kind === 'paragraph' ? node.drawings : []))[0]!;
    const command = drawing.commands.find((candidate) => candidate.kind === 'drawingml-image-fill');
    if (command?.kind !== 'drawingml-image-fill') throw new Error('tiled picture fill command');
    expect(command.plan.rect).toEqual({ x: 0, y: 0, w: 612, h: 792 });
    const descriptor = layoutSourceStore(model).paintResources.descriptors
      .find((candidate) => candidate.resourceKey === command.resourceKey);
    expect(descriptor).toMatchObject({ kind: 'image', tile: { sx: 0.5, sy: 0.5, flip: 'x', algn: 'tl' } });
    const painted: Array<[string, LayoutRect]> = [];
    paint(result, 0, { paint: (key: string, _kind: PaintResourceKind, bounds: LayoutRect) => { painted.push([key, bounds]); } });
    expect(painted).toEqual([[command.resourceKey, { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 }]]);
  });
});
