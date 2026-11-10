/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { Canvas, loadImage } from 'skia-canvas';
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas } from './renderer.js';

const encode = (s: string) => new TextEncoder().encode(s);

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});
afterEach(() => vi.unstubAllGlobals());

async function renderVml(content: string, monochrome: boolean | 'transparent' | 'vertical' = false) {
  const source = new Canvas(20, 10);
  const vertical = monochrome === 'vertical';
  source.getContext('2d').fillStyle = monochrome === true ? '#000000' : '#ff0000';
  if (monochrome !== 'transparent') source.getContext('2d').fillRect(0, 0, vertical ? 20 : 10, vertical ? 5 : 10);
  source.getContext('2d').fillStyle = monochrome === true ? '#ffffff' : '#0000ff';
  source.getContext('2d').fillRect(vertical ? 0 : 10, vertical ? 5 : 0, vertical ? 20 : 10, vertical ? 5 : 10);
  const png = new Uint8Array(source.toBufferSync('png'));
  vi.stubGlobal('OffscreenCanvas', Canvas);
  vi.stubGlobal('createImageBitmap', async (source: Blob | Canvas) => loadImage(source instanceof Blob
    ? Buffer.from(await source.arrayBuffer()) : source.toBufferSync('png')));
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="pic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/two.png"/></Relationships>')],
    ['word/media/two.png', png],
    ['word/document.xml', encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:pict>
${content}</w:pict></w:r></w:p><w:sectPr><w:pgSz w:w="4000" w:h="4000"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`)],
  ]));
  const archive = new DocxArchive(bytes);
  let model;
  try { model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document; }
  finally { archive.free(); }
  const canvas = new Canvas(200, 200);
  const diagnostics: unknown[] = [];
  await renderDocumentToCanvas(model, canvas as unknown as HTMLCanvasElement, 0, {
    width: 200, dpr: 1, fetchImage: async () => new Blob([png], { type: 'image/png' }),
    onDiagnostic: d => diagnostics.push(d),
  });
  const pixels = canvas.getContext('2d').getImageData(0, 0, 200, 200).data;
  const layout = layoutDocument(model, createLayoutServices(model, {
    measureContext: canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
  }));
  return { pixels, diagnostics: [...layout.diagnostics, ...diagnostics], model, layout };
}

function counter(pixels: Uint8ClampedArray) {
  const count = (r: number, g: number, b: number) => {
    let n = 0;
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] === r && pixels[i + 1] === g && pixels[i + 2] === b && pixels[i + 3] === 255) n++;
    return n;
  };
  return count;
}

it.each(['shape', 'picture'])('combines standalone %s offsets and margins in root, nested and scaled groups', async kind => {
  const referenceFrame = 'mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-position-horizontal:absolute;mso-position-vertical:absolute';
  const leaf = (position: string, width = 40, height = 20) => `<v:rect style="position:absolute;${referenceFrame};${position};width:${width}pt;height:${height}pt" fillcolor="red" stroked="f">${kind === 'picture' ? '<v:imagedata r:id="pic"/>' : ''}</v:rect>`;
  const margin = 'left:10;top:20;margin-left:60pt;margin-top:50pt';
  const reference = await renderVml(leaf('left:70pt;top:70pt'));
  expect(counter(reference.pixels)(255, 0, 0)).toBeGreaterThan(100);
  expect(reference.diagnostics).toEqual([]);
  for (const content of [
    leaf(margin),
    leaf(margin).replace(referenceFrame, ''),
    leaf('left:70;top:70;margin-left:0;margin-top:0'),
    `<v:group style="position:absolute;${referenceFrame};${margin};width:40pt;height:20pt" coordsize="40,20">${leaf('left:0;top:0')}</v:group>`,
    `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="150,150"><v:group style="${margin};width:40;height:20" coordsize="40,20">${leaf('left:0;top:0')}</v:group></v:group>`,
    `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="150,150">${leaf(margin)}</v:group>`,
    `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="150,150">${leaf('left:70;top:70;margin-left:0;margin-top:0')}</v:group>`,
    `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="300,300">${leaf('left:20;top:40;margin-left:120;margin-top:100', 80, 40)}</v:group>`,
  ]) {
    const actual = await renderVml(content);
    expect(actual.diagnostics).toEqual([]);
    expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
  }
});

it.each(['shape', 'picture'])('composes nested group and %s margins before parent scaling', async kind => {
  const leaf = (position: string) => `<v:rect style="position:absolute;${position};width:10;height:5" fillcolor="red" stroked="f">${kind === 'picture' ? '<v:imagedata r:id="pic"/>' : ''}</v:rect>`;
  const reference = await renderVml(leaf('left:67;top:43'));
  expect(counter(reference.pixels)(255, 0, 0)).toBeGreaterThan(10);
  expect(reference.diagnostics).toEqual([]);
  const actual = await renderVml(`<v:group style="position:absolute;width:150pt;height:150pt" coordsize="300,300" coordorigin="10,20"><v:group style="left:20;top:10;margin-left:100;margin-top:80;width:80;height:40" coordsize="40,20">${leaf('left:2;top:3;margin-left:10;margin-top:5')}</v:group></v:group>`);
  expect(actual.diagnostics).toEqual([]);
  expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
});

it.each(['shape', 'picture'])('uses parent coordinates for grouped %s CSS units and nested extents', async kind => {
  const leaf = (style: string) => `<v:rect style="position:absolute;${style}" fillcolor="red" stroked="f">${kind === 'picture' ? '<v:imagedata r:id="pic"/>' : ''}</v:rect>`;
  const root = (content: string) => `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="300,300" coordorigin="10,20">${content}</v:group>`;
  const reference = await renderVml(leaf('left:67pt;top:43pt;width:30pt;height:20pt'));
  expect(counter(reference.pixels)(255, 0, 0)).toBeGreaterThan(100);
  if (kind === 'picture') expect(counter(reference.pixels)(0, 0, 255)).toBeGreaterThan(100);
  expect(reference.diagnostics).toEqual([]);
  // Exercise all six properties together, including mixed units. The nested
  // group scales its children by two; the outer group scales by one half.
  for (const units of [Array<string>(6).fill('px'), ['in', 'cm', 'mm', 'pc', 'px', 'pt'], ['em', 'ex', 'em', 'ex', 'em', 'ex']]) {
    const style = (values: number[]) => ['left', 'top', 'margin-left', 'margin-top', 'width', 'height']
      .map((property, i) => `${property}:${values[i]}${units[i]}`).join(';');
    for (const content of [
      root(leaf(style([24, 26, 120, 80, 60, 40]))),
      root(`<v:group style="${style([20, 10, 100, 80, 80, 40])}" coordsize="40,20">${leaf(style([2, 3, 10, 5, 30, 20]))}</v:group>`),
    ]) {
      const actual = await renderVml(content);
      expect(actual.diagnostics).toEqual([]);
      expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
    }
  }
});

it('retains physical CSS units for root groups and standalone leaves beside groups', async () => {
  const leaf = (style: string) => `<v:rect style="position:absolute;${style}" fillcolor="red" stroked="f"/>`;
  const root = (size: string) => `<v:group style="position:absolute;width:${size};height:${size}" coordsize="100,100">${leaf('left:10;top:20;width:30;height:40')}</v:group>`;
  const reference = await renderVml(root('75pt'));
  expect(counter(reference.pixels)(255, 0, 0)).toBeGreaterThan(100);
  expect(reference.diagnostics).toEqual([]);
  for (const content of [
    root('100px'),
    leaf('left:10px;top:20px;width:30px;height:40px'),
    `<v:group style="position:absolute;width:100pt;height:100pt"/>${leaf('left:10px;top:20px;width:30px;height:40px')}`,
  ]) {
    const actual = await renderVml(content);
    expect(actual.diagnostics).toEqual([]);
    expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
  }
});

it('paints equivalent textpath font shorthand, longhands and template overrides identically', async () => {
  const normal = "font-family:'Times New Roman';font-size:24pt;font-weight:normal;font-style:normal";
  const draw = (template: string, style: string, grouped = true) => {
    const content = `<v:shapetype id="font"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" style="${template}"/></v:shapetype><v:shape type="#font" style="position:absolute;width:100pt;height:50pt" fillcolor="red" stroked="f"><v:textpath string="Fjp" style="${style}"/></v:shape>`;
    return renderVml(grouped ? `<v:group style="position:absolute;width:100pt;height:50pt" coordsize="100,50">${content}</v:group>` : content);
  };
  const reference = await draw('', normal);
  expect(counter(reference.pixels)(255, 0, 0)).toBeGreaterThan(100);
  expect(reference.diagnostics).toEqual([]);
  for (const [template, style] of [
    ['font-weight:bold;font-style:italic', "font:normal normal normal 24pt 'Times New Roman'"],
    ["font:24pt 'Times New Roman'", ''],
    ['', "font-weight:bold;font-style:italic;font:24pt / 28pt Times New Roman"],
    ["font:italic bold 24pt 'Times New Roman'", 'font-weight:normal;font-style:normal'],
    ['', `font:italic bold 12pt Arial;${normal}`],
  ]) {
    const result = await draw(template, style);
    expect(result.diagnostics, `${template} / ${style}`).toEqual([]);
    expect(result.pixels, `${template} / ${style}`).toEqual(reference.pixels);
  }
  const standalone = await draw('', normal, false);
  expect(standalone.diagnostics).toEqual([]);
  expect(standalone.pixels).toEqual(reference.pixels);
  const bold = await draw('', "font:italic bold 24pt 'Times New Roman'");
  expect(bold.diagnostics).toEqual([]);
  expect(bold.pixels).not.toEqual(reference.pixels);
  // ~3.4 s on an idle machine; the default 5 s budget timed out when
  // the suite ran on a loaded machine.
}, 30_000);

it('paints custom and inherited paths, primitives, cropped pixels and leaf shadows from a VML group', async () => {
  const { pixels, diagnostics } = await renderVml(`      <v:group style="width:140pt;height:120pt" coordsize="140,120">
        <v:shapetype id="triangle" coordsize="100,100" adj="100" path="m0,0l@0,0,0,100xe"><v:formulas><v:f eqn="val #0"/></v:formulas></v:shapetype>
        <v:shape type="#triangle" style="width:40;height:40" fillcolor="#ff00ff" stroked="f"/>
        <v:image src="media/two.png" style="left:50;width:40;height:40" cropleft="50%"/>
        <v:shape style="left:95;width:30;height:40" stroked="f" filled="f"><v:imagedata r:id="pic" cropleft="32768f"/><v:shadow on="t" color="#ffa500" offset="4pt,4pt" obscured="t"/></v:shape>
        <v:rect style="top:50;width:20;height:20" fillcolor="yellow" stroked="f"><v:shadow on="t" color="lime" offset="10pt,10pt"/></v:rect>
        <v:polyline points="50,60 60,80 70,60" filled="f" strokecolor="#800080" strokeweight="3pt"/>
        <v:curve from="80,60" control1="80,80" control2="100,80" to="100,60" filled="f" strokecolor="#008080" strokeweight="3pt"/>
        <v:shape style="top:90;width:40;height:20" coordsize="100,100" path="m0,0l100,0,100,100,0,100xm25,25l75,25,75,75,25,75xe" fillcolor="#a52a2a" stroked="f"/>
        <v:arc style="left:110;top:60;width:20;height:20" filled="f" strokecolor="#808000" strokeweight="3pt"/>
      </v:group>`);
  const count = counter(pixels);
  expect(count(255, 0, 255)).toBeGreaterThan(500);
  expect(count(0, 0, 255)).toBeGreaterThan(2000);
  expect(count(255, 165, 0)).toBeGreaterThan(100);
  expect(count(255, 0, 0)).toBe(0); // The cropped-away left source half cannot leak.
  expect(count(0, 255, 0)).toBeGreaterThan(200); // Visible shadow outside the yellow rectangle.
  for (const color of [[128, 0, 128], [0, 128, 128], [128, 128, 0]]) {
    expect(count(color[0], color[1], color[2])).toBeGreaterThan(10);
  }
  expect(count(165, 42, 42)).toBeGreaterThan(400);
  expect(count(165, 42, 42)).toBeLessThan(700); // Same-winding inner contour is still a hole.
  expect(diagnostics).toEqual([]);
});

it('paints effective picture outlines and inherited opacity independently from the image', async () => {
  const { pixels, diagnostics } = await renderVml(`<v:group style="width:140pt;height:80pt" coordsize="140,80">
    <v:shapetype id="outlined"><v:stroke color="lime" weight="6pt"/></v:shapetype>
    <v:shape style="left:5;top:5;width:40;height:40" stroked="t" strokecolor="red" strokeweight="6pt"><v:imagedata r:id="pic" cropleft="50%"/></v:shape>
    <v:shape type="#outlined" style="left:60;top:5;width:40;height:40"><v:imagedata r:id="pic" cropleft="50%"/></v:shape>
  </v:group>`);
  const count = counter(pixels);
  expect(count(255, 0, 0)).toBeGreaterThan(200);
  expect(count(0, 255, 0)).toBeGreaterThan(200);
  expect(count(0, 0, 255)).toBeGreaterThan(1000);
  expect(diagnostics).toEqual([]);
});

it('paints zero and partial stroke alpha from direct and inherited stroke properties', async () => {
  const hidden = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:rect style="left:10;top:10;width:40;height:40" filled="f"><v:stroke color="red" weight="6pt" opacity="0"/></v:rect></v:group>`);
  expect(counter(hidden.pixels)(255, 255, 255)).toBe(40000);
  const partial = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:shapetype id="alpha"><v:stroke color="red" weight="6pt" opacity="32768f"/></v:shapetype><v:rect type="#alpha" style="left:10;top:10;width:40;height:40" filled="f"/></v:group>`);
  expect(counter(partial.pixels)(255, 127, 127)).toBeGreaterThan(300);
  expect(counter(partial.pixels)(255, 0, 0)).toBe(0);
  expect([...hidden.diagnostics, ...partial.diagnostics]).toEqual([]);
});

it('renders every non-solid VML fill through direct and inherited brushes', async () => {
  for (const fill of [
    '<v:fill type="gradient" color="red" color2="blue" angle="90" colors="50% lime"/>',
    '<v:fill type="gradientradial" color="red" color2="blue" focusposition="50%,50%" focussize="20%,20%"/>',
    '<v:fill type="tile" r:id="pic" size="20pt,20pt" origin="0,0"/>',
    '<v:fill type="pattern" r:id="pic" color="lime" color2="yellow" size="20pt,20pt" origin="0,0"/>',
    '<v:fill type="frame" r:id="pic"/>',
  ]) {
    const direct = await renderVml(`<v:rect style="width:80pt;height:80pt" stroked="f">${fill}</v:rect>`, true);
    const inherited = await renderVml(`<v:shapetype id="brush">${fill}</v:shapetype><v:rect type="#brush" style="width:80pt;height:80pt" stroked="f"/>`, true);
    const grouped = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:shapetype id="brush">${fill}</v:shapetype><v:rect type="#brush" style="width:80;height:80" stroked="f"/></v:group>`, true);
    expect(counter(grouped.pixels)(255, 255, 255), fill).toBeLessThan(38000);
    expect(inherited.pixels, fill).toEqual(direct.pixels);
    const count = counter(direct.pixels);
    if (fill.includes('gradient')) {
      expect(count(255, 255, 255), fill).toBeLessThan(35000);
      // The gradient must contain mixed channels rather than a solid primary.
      let blended = 0;
      for (let i = 0; i < direct.pixels.length; i += 4) {
        const [r, g, b] = direct.pixels.slice(i, i + 3);
        if ([r, g, b].filter(channel => channel > 0 && channel < 255).length >= 2) blended++;
      }
      expect(blended, fill).toBeGreaterThan(100);
    } else if (fill.includes('pattern')) {
      expect(count(0, 255, 0), fill).toBeGreaterThan(1000);
      expect(count(255, 255, 0), fill).toBeGreaterThan(1000);
    } else {
      expect(count(0, 0, 0), fill).toBeGreaterThan(1000);
      expect(count(255, 255, 255), fill).toBeLessThan(38000);
    }
    if (!fill.includes('gradient')) {
      const target = fill.includes('pattern') ? [0, 255, 0] : [0, 0, 0];
      let maximumRuns = 0;
      for (let y = 0; y < 200; y++) {
        let runs = 0, inside = false;
        for (let x = 0; x < 200; x++) {
          const offset = (y * 200 + x) * 4;
          const matches = target.every((channel, i) => direct.pixels[offset + i] === channel);
          if (matches && !inside) runs++;
          inside = matches;
        }
        maximumRuns = Math.max(maximumRuns, runs);
      }
      if (fill.includes('frame')) expect(maximumRuns, fill).toBe(1);
      else expect(maximumRuns, fill).toBeGreaterThanOrEqual(3);
    }
    expect([...direct.diagnostics, ...inherited.diagnostics, ...grouped.diagnostics], fill).toEqual([]);
  }
  // ~3.5 s on an idle machine; the default 5 s budget timed out when
  // the suite ran on a loaded machine.
}, 30_000);

it('paints arc, curve, polyline and image siblings beside a VML group', async () => {
  const { pixels, diagnostics } = await renderVml(`<v:group style="width:140pt;height:100pt"/>
    <v:arc style="left:10pt;top:10pt;width:20pt;height:20pt" filled="f" strokecolor="lime" strokeweight="4pt"/>
    <v:curve from="40,10" control1="40,30" control2="60,30" to="60,10" filled="f" strokecolor="yellow" strokeweight="4pt"/>
    <v:polyline points="70,10 80,30 90,10" filled="f" strokecolor="magenta" strokeweight="4pt"/>
    <v:image src="media/two.png" style="left:100pt;top:10pt;width:20pt;height:20pt" stroked="f"/>`);
  const count = counter(pixels);
  for (const color of [[0, 255, 0], [255, 255, 0], [255, 0, 255], [255, 0, 0], [0, 0, 255]]) {
    expect(count(...color as [number, number, number])).toBeGreaterThan(10);
  }
  expect(diagnostics).toEqual([]);
});

it('keeps a window-aligned frame brush upright through host rotation and reflection', async () => {
  const fill = '<v:fill type="frame" r:id="pic" alignshape="false"/>';
  const ordinary = await renderVml(`<v:rect style="width:80pt;height:80pt" stroked="f">${fill}</v:rect>`, true);
  const transformed = await renderVml(`<v:rect style="width:80pt;height:80pt;rotation:90;flip:x" stroked="f">${fill}</v:rect>`, true);
  expect(transformed.pixels).toEqual(ordinary.pixels);
  expect([...ordinary.diagnostics, ...transformed.diagnostics]).toEqual([]);
});


it('keeps direct, inherited gradient and image backing brushes behind transparent pictures', async () => {
  const { pixels, diagnostics } = await renderVml(`<v:group style="width:150pt;height:100pt" coordsize="150,100">
    <v:shapetype id="gradient"><v:fill type="gradient" color="yellow" color2="lime" angle="90"/></v:shapetype>
    <v:shape style="left:5;top:5;width:40;height:40" fillcolor="lime" strokecolor="magenta" strokeweight="4pt"><v:imagedata r:id="pic"/></v:shape>
    <v:shape type="#gradient" style="left:55;top:5;width:40;height:40" stroked="f"><v:imagedata r:id="pic"/></v:shape>
    <v:shape style="left:105;top:5;width:40;height:40" stroked="f"><v:fill type="tile" r:id="pic" size="20pt,20pt"/><v:imagedata r:id="pic" cropright="50%"/></v:shape>
  </v:group>`, 'transparent');
  const count = counter(pixels);
  expect(count(0, 255, 0)).toBeGreaterThan(500);
  expect(count(255, 0, 255)).toBeGreaterThan(100);
  expect(count(0, 0, 255)).toBeGreaterThan(1000);
  // Only the third picture occupies this region. Its source crop is entirely
  // transparent; blue here must come from the independently acquired tile.
  let tiledBlue = 0;
  for (let i = 0; i < pixels.length; i += 4) {
    if ((i / 4) % 200 >= 100 && pixels[i] === 0 && pixels[i + 1] === 0 && pixels[i + 2] === 255) tiledBlue++;
  }
  expect(tiledBlue).toBeGreaterThan(500);
  let gradient = 0;
  for (let i = 0; i < pixels.length; i += 4) {
    if (pixels[i] > 0 && pixels[i] < 255 && pixels[i + 1] === 255 && pixels[i + 2] === 0) gradient++;
  }
  expect(gradient).toBeGreaterThan(500);
  expect(diagnostics).toEqual([]);
});

it('cascades each picture property across metadata-only and partial imagedata overrides', async () => {
  const expected = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:shape style="width:80;height:80" stroked="f"><v:imagedata r:id="pic" cropleft="50%"/></v:shape></v:group>`);
  for (const override of ['title="metadata"', 'r:id="pic"', 'croptop="0"']) {
    const inherited = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:shapetype id="picture"><v:imagedata r:id="pic" cropleft="50%"/></v:shapetype><v:shape type="#picture" style="width:80;height:80" stroked="f"><v:imagedata ${override}/></v:shape></v:group>`);
    expect(inherited.pixels, override).toEqual(expected.pixels);
    expect(inherited.diagnostics, override).toEqual([]);
  }
  expect(counter(expected.pixels)(255, 0, 0)).toBe(0);
  expect(counter(expected.pixels)(0, 0, 255)).toBeGreaterThan(5000);
});

it('applies effective host and child grayscale/chromakey and discloses unavailable host image effects', async () => {
  for (const picture of [
    '<v:image src="media/two.png" grayscale="t" style="width:80;height:80" stroked="f"/>',
    '<v:shape style="width:80;height:80" stroked="f"><v:imagedata r:id="pic" grayscale="t"/></v:shape>',
    '<v:shapetype id="gray"><v:imagedata r:id="pic" grayscale="t"/></v:shapetype><v:shape type="#gray" style="width:80;height:80" stroked="f"><v:imagedata title="metadata"/></v:shape>',
  ]) {
    const { pixels, diagnostics } = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80">${picture}</v:group>`);
    expect(counter(pixels)(255, 0, 0)).toBe(0);
    expect(counter(pixels)(0, 0, 255)).toBe(0);
    let gray = 0;
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] > 0 && pixels[i] < 255 && pixels[i] === pixels[i + 1] && pixels[i] === pixels[i + 2]) gray++;
    expect(gray).toBeGreaterThan(5000);
    expect(diagnostics).toEqual([]);
  }
  const keyed = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:image src="media/two.png" chromakey="red" fillcolor="lime" style="width:80;height:80" stroked="f"/></v:group>`);
  expect(counter(keyed.pixels)(255, 0, 0)).toBe(0);
  expect(counter(keyed.pixels)(0, 255, 0)).toBeGreaterThan(2500);
  expect(keyed.diagnostics).toEqual([]);
  // gain/blacklevel are the shared luminance correction (#2175): blacklevel
  // .2 is bright .4, which lifts pure red to (255,102,102).
  const lifted = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:image src="media/two.png" blacklevel=".2" style="width:80;height:80" stroked="f"/></v:group>`);
  expect(lifted.diagnostics).toEqual([]);
  expect(counter(lifted.pixels)(255, 0, 0)).toBe(0);
  expect(counter(lifted.pixels)(255, 102, 102)).toBeGreaterThan(2500);
  for (const attribute of ['gamma="2"', 'gain="2" grayscale="f" chromakey="blue"', 'bilevel="t"']) {
    const unavailable = await renderVml(`<v:group style="width:80pt;height:80pt" coordsize="80,80"><v:image src="media/two.png" ${attribute} style="width:80;height:80" stroked="f"/></v:group>`);
    expect(unavailable.diagnostics.length, attribute).toBeGreaterThan(0);
    expect(counter(unavailable.pixels)(255, 0, 0), attribute).toBeGreaterThan(2500);
  }
});

it('applies picture rotation/reflection once and composes nested group transforms', async () => {
  const picture = (transform: string) => `<v:image src="media/two.png" style="left:30;top:30;width:40;height:40;${transform}" stroked="f"/>`;
  const group = (content: string, transform = '') => `<v:group style="width:100pt;height:100pt;${transform}" coordsize="100,100">${content}</v:group>`;
  for (const [content, axis, direction, source] of [
    [group(picture('rotation:90')), 'y', -1, false],
    [group(picture('flip:x')), 'x', 1, false],
    [group(picture('flip:y')), 'y', 1, 'vertical'],
    [group(group(picture('flip:x'), 'flip:x'), 'rotation:90'), 'y', -1, false],
  ] as const) {
    const { pixels, diagnostics } = await renderVml(content, source);
    const centers = [[0, 0], [0, 0]];
    for (let i = 0; i < pixels.length; i += 4) {
      const color = pixels[i] === 255 && pixels[i + 1] === 0 && pixels[i + 2] === 0 ? 0
        : pixels[i] === 0 && pixels[i + 1] === 0 && pixels[i + 2] === 255 ? 1 : -1;
      if (color < 0) continue;
      centers[color][0] += axis === 'x' ? (i / 4) % 200 : Math.floor(i / 4 / 200);
      centers[color][1]++;
    }
    expect(centers[0][1], content).toBeGreaterThan(500);
    expect(centers[1][1], content).toBeGreaterThan(500);
    expect(direction * (centers[0][0] / centers[0][1] - centers[1][0] / centers[1][1]), content).toBeGreaterThan(15);
    expect(diagnostics, content).toEqual([]);
  }
});

it('paints solid fill opacity independently of stroke and shadow and inherits each fill property', async () => {
  const group = (content: string) => `<v:group style="width:100pt;height:100pt" coordsize="100,100">${content}</v:group>`;
  const shape = (type: string, fill: string) => `<v:rect ${type} style="left:10;top:10;width:50;height:50" fillcolor="red" strokecolor="lime" strokeweight="4pt">${fill}<v:shadow on="t" color="blue" opacity="1" offset="12pt,12pt"/></v:rect>`;
  const noStroke = await renderVml(group('<v:rect style="width:50;height:50" fillcolor="red" stroked="f"><v:fill opacity="0"/></v:rect>'));
  expect(counter(noStroke.pixels)(255, 255, 255)).toBe(40000);
  const hidden = await renderVml(group(shape('', '<v:fill opacity="0"/>')));
  expect(counter(hidden.pixels)(255, 0, 0)).toBe(0);
  expect(counter(hidden.pixels)(255, 255, 255)).toBeGreaterThan(35000);
  expect(counter(hidden.pixels)(0, 255, 0)).toBeGreaterThan(500);
  expect(counter(hidden.pixels)(0, 0, 255)).toBeGreaterThan(500);
  const partial = await renderVml(group(shape('', '<v:fill opacity=".5"/>')));
  expect(counter(partial.pixels)(255, 128, 128)).toBeGreaterThan(500);
  expect(counter(partial.pixels)(0, 255, 0)).toBeGreaterThan(500);
  expect(counter(partial.pixels)(0, 0, 255)).toBeGreaterThan(500);
  expect(counter(partial.pixels)(255, 0, 0)).toBe(0);
  const inherited = await renderVml(group(`<v:shapetype id="alpha"><v:fill color="yellow" opacity="32768f"/></v:shapetype>${shape('type="#alpha"', '<v:fill type="solid"/>')}`));
  expect(inherited.pixels).toEqual(partial.pixels);
  expect([...noStroke.diagnostics, ...hidden.diagnostics, ...partial.diagnostics, ...inherited.diagnostics]).toEqual([]);
});

it.each(['rect', 'polyline'])('preserves omitted %s joins and explicit pen inheritance in standalone and grouped VML', async kind => {
  const render = (definition: string, stroke: string, grouped = true) => {
    // The polyline's acute corner distinguishes miter limits 8 and 10;
    // the rectangle catches the square-corner regression at ordinary joins.
    const geometry = kind === 'rect'
      ? '<v:rect type="#pen" style="position:absolute;left:30pt;top:30pt;width:50pt;height:50pt"'
      : '<v:polyline type="#pen" points="34.5pt,80pt 40pt,30pt 45.5pt,80pt"';
    const content = `${definition}${geometry} filled="f" strokecolor="red" strokeweight="8pt">${stroke}</v:${kind}>`;
    return renderVml(grouped
      ? `<v:group style="position:absolute;width:150pt;height:150pt" coordsize="150,150">${content}</v:group>`
      : content);
  };
  const miter = await render('', '<v:stroke joinstyle="miter" miterlimit="8"/>');
  const round = await render('', '<v:stroke joinstyle="round"/>');
  expect(miter.pixels).not.toEqual(round.pixels);
  expect(counter(miter.pixels)(255, 0, 0)).toBeGreaterThan(100);
  expect([...miter.diagnostics, ...round.diagnostics]).toEqual([]);
  for (const grouped of [false, true]) {
    const omitted = await render('', '', grouped);
    const inherited = await render('<v:shapetype id="pen"><v:stroke joinstyle="round"/></v:shapetype>', '<v:stroke color="red"/>', grouped);
    const overridden = await render('<v:shapetype id="pen"><v:stroke joinstyle="round"/></v:shapetype>', '<v:stroke joinstyle="miter"/>', grouped);
    expect(Buffer.compare(Buffer.from(omitted.pixels), Buffer.from(miter.pixels))).toBe(0);
    expect(Buffer.compare(Buffer.from(inherited.pixels), Buffer.from(round.pixels))).toBe(0);
    expect(Buffer.compare(Buffer.from(overridden.pixels), Buffer.from(miter.pixels))).toBe(0);
    expect([...omitted.diagnostics, ...inherited.diagnostics, ...overridden.diagnostics]).toEqual([]);
  }
  if (kind === 'polyline') {
    const largerLimit = await render('', '<v:stroke joinstyle="miter" miterlimit="10"/>');
    expect(miter.pixels).not.toEqual(largerLimit.pixels);
    expect(largerLimit.diagnostics).toEqual([]);
  }
});


it('inherits VML visibility through groups while preserving the root flow extent and visible siblings', async () => {
  const red = '<v:rect style="position:absolute;left:10;top:10;width:60;height:60" fillcolor="red" stroked="f"/>';
  const wrap = (content: string, style = '') => `<v:group style="width:100pt;height:100pt;${style}" coordsize="100,100">${content}</v:group>`;
  const visible = await renderVml(wrap(red));
  expect(counter(visible.pixels)(255, 0, 0)).toBeGreaterThan(3000);
  for (const content of [
    wrap(red.replace('position:absolute;', 'position:absolute;visibility:hidden;')),
    wrap(`<v:group style="width:100;height:100;visibility:hidden" coordsize="100,100">${red}</v:group>`),
    wrap(red, 'visibility:hidden'),
  ]) {
    const hidden = await renderVml(content);
    expect(counter(hidden.pixels)(255, 0, 0)).toBe(0);
    const flowBounds = (result: typeof visible) => result.layout.pages.flatMap(page =>
      page.layers.body.filter(node => node.kind === 'paragraph').map(node => node.flowBounds));
    expect(flowBounds(hidden)).toEqual(flowBounds(visible));
    expect(hidden.diagnostics).toEqual([]);
  }
  const sibling = await renderVml(wrap(`<v:group style="width:100;height:100;visibility:hidden" coordsize="100,100">${red.replace('position:absolute;', 'position:absolute;visibility:visible;')}</v:group>${red.replace('fillcolor="red"', 'fillcolor="lime"')}`));
  expect(counter(sibling.pixels)(255, 0, 0)).toBe(0);
  expect(counter(sibling.pixels)(0, 255, 0)).toBeGreaterThan(3000);
  expect(sibling.diagnostics).toEqual([]);
});

it('lets effective VML path arrowok override both arrowheads, including its false default', async () => {
  const wrap = (definition: string, path: string) => `<v:group style="width:100pt;height:100pt" coordsize="100,100">${definition}<v:shape type="#pen" style="position:absolute;left:20;top:20;width:60;height:60" coordsize="100,100" filled="f" strokecolor="red" strokeweight="6pt"><v:stroke startarrow="block" startarrowwidth="wide" startarrowlength="long" endarrow="block" endarrowwidth="wide" endarrowlength="long"/>${path}</v:shape></v:group>`;
  const path = (flag: string) => `<v:path ${flag} v="m0,0l100,100e"/>`;
  const none = await renderVml(wrap('', path('arrowok="f"')));
  const arrows = await renderVml(wrap('', path('arrowok="t"')));
  // Both ends carry a wide/long (lg/lg) triangle on a 6 pt pen: 30 pt long
  // and 30 pt wide (5 × line width, core arrow-geometry.ts). The line stops
  // 27 pt behind each tip. At 1 px/pt the tail tip is (80, 80), pointing along
  // (1, 1)/√2. Pixel (58, 68) lies 23.3 behind it and 7.1 off the axis: inside
  // the head (half-width 11.7) and outside the 3 pt shaft half-width, with
  // more than a pixel's margin either way. (41, 31) mirrors it at the head.
  const at = (pixels: Uint8ClampedArray, x: number, y: number) => [...pixels.slice((y * 200 + x) * 4, (y * 200 + x) * 4 + 3)];
  for (const [x, y] of [[58, 68], [41, 31]]) {
    expect(at(arrows.pixels, x, y)).toEqual([255, 0, 0]);
    expect(at(none.pixels, x, y)).toEqual([255, 255, 255]);
  }
  for (const [definition, flag, expected] of [
    ['', '', none],
    ['<v:shapetype id="pen"><v:path arrowok="f"/></v:shapetype>', '', none],
    ['<v:shapetype id="pen"><v:path arrowok="t"/></v:shapetype>', 'arrowok="f"', none],
    ['<v:shapetype id="pen"><v:path arrowok="f"/></v:shapetype>', 'arrowok="t"', arrows],
    ['<v:shapetype id="pen"><v:path arrowok="t"/></v:shapetype>', '', arrows],
  ] as const) {
    const actual = await renderVml(wrap(definition, path(flag)));
    expect(actual.pixels).toEqual(expected.pixels);
    expect(actual.diagnostics).toEqual([]);
  }
  expect([...none.diagnostics, ...arrows.diagnostics]).toEqual([]);
});

it('renders omitted VML line endpoints with the normative defaults', async () => {
  const wrap = (endpoints: string) => `<v:group style="width:100pt;height:100pt" coordsize="100,100"><v:line style="position:absolute" ${endpoints} strokecolor="red" strokeweight="6pt"/></v:group>`;
  for (const [omitted, explicit] of [['to="60,60"', 'from="0,0" to="60,60"'], ['', 'from="0,0" to="10,10"']]) {
    const actual = await renderVml(wrap(omitted));
    const control = await renderVml(wrap(explicit));
    expect(counter(actual.pixels)(255, 0, 0)).toBeGreaterThan(10);
    expect(actual.pixels).toEqual(control.pixels);
    expect([...actual.diagnostics, ...control.diagnostics]).toEqual([]);
  }
});


it('applies host fill opacity with per-property shapetype and decoration overrides', async () => {
  const wrap = (definition: string, host: string, fill = '') => `<v:group style="width:80pt;height:80pt" coordsize="80,80">${definition}<v:rect type="#alpha" style="left:10;top:10;width:40;height:40" fillcolor="red" stroked="f" ${host}>${fill}</v:rect></v:group>`;
  const expected = await renderVml(wrap('', '', '<v:fill opacity=".5"/>'));
  for (const [definition, host] of [
    ['', 'opacity="32768f"'],
    ['<v:shapetype id="alpha" opacity=".5"/>', ''],
    ['<v:shapetype id="alpha" opacity="0"/>', 'opacity=".5"'],
  ]) {
    const actual = await renderVml(wrap(definition, host));
    expect(actual.pixels).toEqual(expected.pixels);
    expect(actual.diagnostics).toEqual([]);
  }
  const reset = await renderVml(wrap('<v:shapetype id="alpha" opacity=".5"/>', 'opacity="1"', '<v:fill opacity="0"/>'));
  expect(counter(reset.pixels)(255, 255, 255)).toBe(40000);
  expect(counter(expected.pixels)(255, 128, 128)).toBeGreaterThan(1000);
  expect([...expected.diagnostics, ...reset.diagnostics]).toEqual([]);
});

it('retains path-shaped radial fills and effective text rectangles', async () => {
  const triangle = (enabled: boolean) =>
    `<v:shape style="width:100pt;height:100pt" coordsize="1000,1000" path="m0,0l1000,1000,0,1000xe" stroked="f"><v:path gradientshapeok="${enabled}"/><v:fill type="gradientradial" color="red" color2="blue" focusposition="25%,75%"/></v:shape>`;
  const ordinary = await renderVml(triangle(false));
  const shaped = await renderVml(triangle(true));
  expect(shaped.pixels).not.toEqual(ordinary.pixels);
  expect(shaped.diagnostics).toEqual([]);
  const translucent = await renderVml(
    '<v:rect style="width:100pt;height:100pt" stroked="f"><v:path gradientshapeok="t"/><v:fill type="gradientradial" color="red" color2="red" opacity=".5" opacity2=".5"/></v:rect>',
  );
  expect([...translucent.pixels.slice((25 * 200 + 25) * 4, (25 * 200 + 25) * 4 + 4)]).toEqual([
    255, 127, 127, 255,
  ]);
  const box = (inherited: boolean, half: boolean) =>
    `<v:group style="width:120pt;height:80pt" coordsize="120,80"><v:shapetype id="box" coordsize="1000,1000">${inherited ? `<v:path textboxrect="0,0,${half ? 500 : 1000},1000"/>` : ''}</v:shapetype><v:shape type="#box" style="width:100;height:60" filled="f" stroked="f">${inherited ? '' : `<v:path textboxrect="0,0,${half ? 500 : 1000},1000"/>`}<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="40"/></w:rPr><w:t>MMMMMMMMMMMM</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></v:group>`;
  const full = await renderVml(box(false, false));
  const half = await renderVml(box(false, true));
  const inherited = await renderVml(box(true, true));
  expect(half.pixels).not.toEqual(full.pixels);
  let fullOverflow = 0,
    halfOverflow = 0;
  for (let y = 0; y < 60; y++)
    for (let x = 50; x < 100; x++) {
      const i = (y * 200 + x) * 4;
      if (full.pixels[i] === 255 && full.pixels[i + 1] === 0 && full.pixels[i + 2] === 0)
        fullOverflow++;
      if (half.pixels[i] === 255 && half.pixels[i + 1] === 0 && half.pixels[i + 2] === 0)
        halfOverflow++;
    }
  expect(fullOverflow).toBeGreaterThan(0);
  expect(halfOverflow).toBe(0);
  expect(inherited.pixels).toEqual(half.pixels);
  expect(half.diagnostics).toEqual([]);
});

it('paints inherited textured and inset pens, including the path veto', async () => {
  const draw = (texture: boolean, inset: boolean, inherited = false, veto = false) =>
    renderVml(
      `<v:group style="width:100pt;height:100pt" coordsize="100,100"><v:shapetype id="pen">${inherited ? `<v:stroke color="red" weight="10pt" insetpen="${inset}" ${texture ? 'filltype="tile" src="media/two.png" imagesize="20pt,10pt"' : ''}/>` : ''}</v:shapetype><v:rect type="#pen" style="left:20;top:20;width:60;height:60" filled="f">${inherited ? '' : `<v:stroke color="red" weight="10pt" insetpen="${inset}" ${texture ? 'filltype="tile" src="media/two.png" imagesize="20pt,10pt"' : ''}/>`}${veto ? '<v:path insetpenok="f"/>' : ''}</v:rect></v:group>`,
    );
  const centered = await draw(false, false);
  const inset = await draw(false, true);
  expect(inset.pixels).not.toEqual(centered.pixels);
  expect((await draw(false, true, true)).pixels).toEqual(inset.pixels);
  expect((await draw(false, true, false, true)).pixels).toEqual(centered.pixels);
  const textured = await draw(true, true);
  expect(counter(textured.pixels)(0, 0, 255)).toBeGreaterThan(50);
  expect((await draw(true, true, true)).pixels).toEqual(textured.pixels);
  expect([...inset.diagnostics, ...textured.diagnostics]).toEqual([]);
});

it('paints fitted straight text paths and effective inherited pattern brushes', async () => {
  const draw = (mode: string, inherited = false, pattern = false) =>
    renderVml(
      `<v:group style="width:100pt;height:70pt" coordsize="100,70"><v:shapetype id="text" coordsize="1000,1000" path="m0,900l1000,900e">${inherited ? `<v:path textpathok="t"/><v:textpath on="t" ${mode} string="MMMM" style="font:20pt Arial"/>` : ''}</v:shapetype><v:shape type="#text" style="width:100;height:60" fillcolor="red" stroked="f">${pattern ? '<v:fill type="pattern" r:id="pic" color="red" color2="blue" size="10pt,10pt"/>' : ''}${inherited ? '' : `<v:path textpathok="t"/><v:textpath on="t" ${mode} string="MMMM" style="font:20pt Arial"/>`}</v:shape></v:group>`,
      true,
    );
  for (const mode of ['fitpath="t"', 'xscale="t"', 'fitshape="t"']) {
    const direct = await draw(mode);
    const inherited = await draw(mode, true);
    expect(counter(direct.pixels)(255, 0, 0)).toBeGreaterThan(50);
    expect(inherited.pixels).toEqual(direct.pixels);
    expect(direct.diagnostics).toEqual([]);
  }
  const direct = await draw('fitpath="t"', false, true);
  expect(
    counter(direct.pixels)(0, 0, 255),
    JSON.stringify({
      d: direct.diagnostics,
      red: counter(direct.pixels)(255, 0, 0),
      white: counter(direct.pixels)(255, 255, 255),
    }),
  ).toBeGreaterThan(10);
  expect((await draw('fitpath="t"', true, true)).pixels).toEqual(direct.pixels);
  expect(direct.diagnostics).toEqual([]);
});

it('keeps rotated image text, miter tips and pen alpha in their acquired ink frames', async () => {
  const text = (pattern: boolean) =>
    `<v:shape style="left:50pt;top:50pt;width:100pt;height:20pt;rotation:90" fillcolor="red" stroked="f"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="MMMMMM" style="font-size:20pt;font-family:Arial"/>${pattern ? '<v:fill type="pattern" r:id="pic" color="red" color2="red" size="10pt,10pt"/>' : ''}</v:shape>`;
  const solid = await renderVml(text(false), true);
  const pattern = await renderVml(text(true), true);
  const inkCount = (pixels: Uint8ClampedArray) => {
    let count = 0;
    for (let i = 0; i < pixels.length; i += 4)
      if (pixels[i] !== 255 || pixels[i + 1] !== 255 || pixels[i + 2] !== 255) count++;
    return count;
  };
  // Rasterizing a brush may change antialiasing, but must preserve the full
  // tall rotated text rather than clipping it to the short authored box.
  expect(inkCount(pattern.pixels)).toBeGreaterThan(inkCount(solid.pixels) * 0.9);
  const pen = (texture: boolean, alpha: number) =>
    `<v:shape style="left:50pt;top:50pt;width:100pt;height:100pt" coordsize="1000,1000" path="m0,0l1000,1000,700,1000xe" filled="f"><v:stroke color="red" weight="10pt" joinstyle="miter" miterlimit="10" opacity="${alpha}" ${texture ? 'filltype="tile" src="media/two.png" imagesize="20pt,10pt"' : ''}/></v:shape>`;
  const plain = await renderVml(pen(false, 1));
  const textured = await renderVml(pen(true, 1));
  const topInk = (pixels: Uint8ClampedArray) => {
    for (let y = 0; y < 200; y++)
      for (let x = 0; x < 200; x++) {
        const i = (y * 200 + x) * 4;
        if (pixels[i] !== 255 || pixels[i + 1] !== 255 || pixels[i + 2] !== 255) return y;
      }
    return 200;
  };
  expect(topInk(textured.pixels)).toBeLessThanOrEqual(topInk(plain.pixels) + 1);
  const translucent = await renderVml(pen(true, 0.5));
  expect(
    counter(translucent.pixels)(255, 127, 127) + counter(translucent.pixels)(127, 127, 255),
  ).toBeGreaterThan(100);
  expect([...pattern.diagnostics, ...textured.diagnostics, ...translucent.diagnostics]).toEqual([]);
});

it('lets instance path and pen properties override inherited drawing modes', async () => {
  const brush = '<v:fill type="gradientradial" color="red" color2="blue" focusposition="25%,75%"/>';
  const triangle = (inherited: boolean, override: string) =>
    `<v:shapetype id="path" coordsize="1000,1000" path="m0,0l1000,1000,0,1000xe"><v:path gradientshapeok="${inherited}"/></v:shapetype><v:shape type="#path" style="width:100pt;height:100pt" stroked="f">${override}${brush}</v:shape>`;
  const ordinary = await renderVml(triangle(false, ''));
  const concentric = await renderVml(triangle(true, ''));
  expect((await renderVml(triangle(true, '<v:path gradientshapeok="f"/>'))).pixels).toEqual(
    ordinary.pixels,
  );
  expect((await renderVml(triangle(false, '<v:path gradientshapeok="t"/>'))).pixels).toEqual(
    concentric.pixels,
  );
  const pen = (inherited: boolean) =>
    `<v:shapetype id="pen">${inherited ? '<v:stroke color="red" weight="10pt" insetpen="t" filltype="tile" src="media/two.png"/>' : ''}</v:shapetype><v:rect type="#pen" style="left:20pt;top:20pt;width:60pt;height:60pt" filled="f" insetpen="f"><v:stroke color="red" weight="10pt" filltype="solid"/></v:rect>`;
  const reset = await renderVml(pen(true));
  expect(reset.pixels).toEqual((await renderVml(pen(false))).pixels);
  expect([...reset.diagnostics, ...concentric.diagnostics]).toEqual([]);
});

it('preserves complete shaped text on an affine baseline, including combining clusters', async () => {
  const draw = (path: string) => renderVml(`<v:${path ? 'shape' : 'rect'} style="width:100pt;height:60pt" coordsize="1000,1000" ${path} fillcolor="red" stroked="f"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="Áffi" style="font-size:20pt;font-family:Arial"/></v:${path ? 'shape' : 'rect'}>`);
  const box = await draw('');
  const baseline = await draw('path="m0,1000l1000,1000e"');
  expect(baseline.pixels).toEqual(box.pixels);
  expect(baseline.diagnostics).toEqual([]);
});

it('maps VML pattern pens by exact black versus coloured source pixels', async () => {
  const pen = '<v:stroke filltype="pattern" src="media/two.png" color="lime" color2="blue" weight="10pt"/>';
  const direct = `<v:rect style="left:20pt;top:20pt;width:60pt;height:60pt" filled="f">${pen}</v:rect>`;
  const mono = await renderVml(direct, true);
  expect(counter(mono.pixels)(0,0,255)).toBeGreaterThan(50);
  expect(counter(mono.pixels)(0,255,0)).toBeGreaterThan(50);
  const coloured = await renderVml(direct);
  expect(counter(coloured.pixels)(0,0,255)).toBe(0);
  expect(counter(coloured.pixels)(0,255,0)).toBeGreaterThan(200);
  const inherited = await renderVml(`<v:shapetype id="pattern">${pen}</v:shapetype><v:rect type="#pattern" style="left:20pt;top:20pt;width:60pt;height:60pt" filled="f"/>`, true);
  expect(inherited.pixels).toEqual(mono.pixels);
  expect([...mono.diagnostics,...coloured.diagnostics,...inherited.diagnostics]).toEqual([]);
});

it('paints fitted text outlines with independent fill, pen alpha and image brushes', async () => {
  const draw = (filled: boolean, stroke = '<v:stroke color="blue" weight="3pt"/>') => renderVml(`<v:group style="width:160pt;height:100pt" coordsize="160,100"><v:shape style="left:20;top:20;width:120;height:50" coordsize="100,100" path="m0,100l100,100e" filled="${filled ? 't' : 'f'}" fillcolor="red"><v:path textpathok="t"/><v:textpath on="t" fitpath="t" string="Fjp" style="font:24pt Arial"/>${stroke}</v:shape></v:group>`);
  const both = await draw(true), outline = await draw(false);
  expect(counter(both.pixels)(255, 0, 0)).toBeGreaterThan(100);
  expect(counter(both.pixels)(0, 0, 255)).toBeGreaterThan(100);
  expect(counter(outline.pixels)(255, 0, 0)).toBe(0);
  expect(counter(outline.pixels)(0, 0, 255)).toBeGreaterThan(100);
  const translucent = await draw(false, '<v:stroke color="blue" weight="3pt" opacity=".5"/>');
  expect(counter(translucent.pixels)(127, 127, 255)).toBeGreaterThan(50);
  const textured = await draw(false, '<v:stroke color="lime" color2="blue" weight="3pt" filltype="pattern" r:id="pic"/>');
  expect(counter(textured.pixels)(0, 255, 0)).toBeGreaterThan(50);
  expect([...both.diagnostics, ...outline.diagnostics, ...translucent.diagnostics, ...textured.diagnostics]).toEqual([]);
});

it('reflects asymmetric text and its image mask through the acquired host rotation and nested frame', async () => {
  const draw = (flip: string, rotation: number, image = false, nested = false, shadow = false) => renderVml(`<v:group style="width:180pt;height:180pt" coordsize="180,180">${nested ? `<v:group style="left:20;top:20;width:140;height:140;flip:${flip};rotation:${rotation}" coordsize="140,140">` : ''}<v:shape style="left:${nested ? 10 : 30};top:${nested ? 40 : 60};width:120;height:60;flip:${nested ? '' : flip};rotation:${nested ? 0 : rotation}" stroked="f" fillcolor="red">${image ? '<v:fill type="pattern" r:id="pic" color="red" color2="red" size="10pt,10pt"/>' : ''}${shadow ? '<v:shadow on="t" color="lime" offset="0,70pt"/>' : ''}<v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="Fjp" style="font:24pt Arial"/></v:shape>${nested ? '</v:group>' : ''}</v:group>`, true);
  const mass = (pixels: Uint8ClampedArray, channel = 1) => {
    let x = 0, y = 0, n = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      const ink = (255 - pixels[i + channel]!) / 255;
      x += ((i / 4) % 200 + .5) * ink; y += (Math.floor(i / 4 / 200) + .5) * ink; n += ink;
    }
    return { x: x / n, y: y / n, n };
  };
  const findRect = (value: unknown): {xPt: number; yPt: number; widthPt: number; heightPt: number} | undefined => {
    if (!value || typeof value !== 'object') return undefined;
    if ('kind' in value && value.kind === 'watermark-text') return (value as unknown as {rect: {xPt: number; yPt: number; widthPt: number; heightPt: number}}).rect;
    for (const child of Object.values(value)) { const r = findRect(child); if (r) return r; }
    return undefined;
  };
  const baseResult = await draw('', 0);
  const frame = findRect(baseResult.layout)!;
  const cx = frame.xPt + frame.widthPt / 2, cy = frame.yPt + frame.heightPt / 2;
  const base = mass(baseResult.pixels);
  const imageBase = mass((await draw('', 0, true)).pixels);
  const shadowBase = mass((await draw('', 0, false, false, true)).pixels, 0);
  for (const axis of ['x', 'y']) {
    const direct = await draw(axis, 0), image = await draw(axis, 0, true);
    const shadow = mass((await draw(axis, 0, false, false, true)).pixels, 0);
    expect(axis === 'x' ? shadow.x + shadowBase.x : shadow.y + shadowBase.y)
      .toBeCloseTo(axis === 'x' ? 2 * cx : 2 * cy + 140, 0);
    const a = mass(direct.pixels), b = mass(image.pixels);

    expect(a.n).toBeCloseTo(base.n, -1);
    expect(axis === 'x' ? a.x + base.x : a.y + base.y).toBeCloseTo(axis === 'x' ? 2 * cx : 2 * cy, 0);
    expect(axis === 'x' ? b.x + imageBase.x : b.y + imageBase.y).toBeCloseTo(axis === 'x' ? 2 * cx : 2 * cy, 0);
    const rotResult = await draw(axis, 90, true);
    const rotated = mass(rotResult.pixels);
    expect(rotated.x).toBeCloseTo(cx + cy - b.y, 0); expect(rotated.y).toBeCloseTo(cy - cx + b.x, 0);
    const nested = await draw(axis, 0, false, true);
    expect(mass(nested.pixels).x).toBeCloseTo(a.x, 0);
    expect(mass(nested.pixels).y).toBeCloseTo(a.y, 0);
    expect([...direct.diagnostics, ...image.diagnostics, ...nested.diagnostics]).toEqual([]);
  }
});

it('acquires skew geometry and independent Office edge pens including inherited defaults and resets', async () => {
  const shape = (skew: string) => `<v:group style="width:140pt;height:120pt" coordsize="140,120"><v:rect style="left:20;top:20;width:60;height:40" fillcolor="red" stroked="f"><o:skew on="t" matrix="1,0.5,0,1,0,0" offset="0,0" origin="-.5,-.5" ${skew}/></v:rect></v:group>`;
  const skewed = await renderVml(shape('')), reset = await renderVml(shape('on="f"').replace('on="t"', ''));
  // Top-left pivot: x=50 is below the original top edge after y += .5*x.
  const at = (p: Uint8ClampedArray, x: number, y: number) => [...p.slice((y * 200 + x) * 4, (y * 200 + x) * 4 + 3)];
  expect(at(skewed.pixels, 60, 25)).toEqual([255, 255, 255]);
  expect(at(skewed.pixels, 60, 65)).toEqual([255, 0, 0]);
  expect(at(reset.pixels, 60, 25)).toEqual([255, 0, 0]);
  const edge = await renderVml(`<v:group style="width:160pt;height:100pt" coordsize="160,100"><v:shapetype id="edges" o:spt="202" coordsize="21600,21600" path="m0,0l21600,0,21600,21600,0,21600xe" filled="f" stroked="f"><v:stroke color="red" weight="15pt" insetpen="t" dashstyle="dash"><o:top on="t" color="blue" weight="8pt"/><o:bottom on="t"/><o:left on="t" color="lime" weight="4pt"/></v:stroke></v:shapetype><v:shape type="#edges" style="left:20;top:20;width:100;height:50"><v:stroke><o:left on="f"/></v:stroke></v:shape></v:group>`);
  expect(counter(edge.pixels)(0, 0, 255)).toBeGreaterThan(500);
  const grayRows = Array.from({length: 50}, (_, y) => y + 40).filter(y => {
    const [r, g, b] = at(edge.pixels, 60, y); return r === g && g === b && r! < 250;
  });
  expect(grayRows.length).toBeGreaterThan(0); expect(grayRows.length).toBeLessThanOrEqual(2);
  expect(counter(edge.pixels)(255, 0, 0)).toBe(0);
  expect(counter(edge.pixels)(0, 255, 0)).toBe(0);
  expect([...skewed.diagnostics, ...reset.diagnostics, ...edge.diagnostics]).toEqual([]);
});

it('keeps trim outline overhang and physical width through nonuniform and curved glyph fitting', async () => {
  for (const path of ['', 'path="m0,20c30,0,70,0,100,20em0,100c30,80,70,80,100,100e"']) {
    const {pixels, diagnostics} = await renderVml(`<v:group style="width:180pt;height:140pt" coordsize="180,140"><v:shape style="left:30;top:30;width:120;height:45" coordsize="100,100" ${path} fillcolor="red" strokecolor="blue" strokeweight="3pt"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" trim="t" string="Fjp" style="font:24pt Arial"/></v:shape></v:group>`);
    let redInk = 0;
    for (let i = 0; i < pixels.length; i += 4)
      if (pixels[i] === 255 && pixels[i+1] === pixels[i+2]) redInk += (255-pixels[i+1]!)/255;
    expect(redInk).toBeGreaterThan(100);
    expect(counter(pixels)(0, 0, 255)).toBeGreaterThan(300);
    if (!path) {
      let left = 200, right = 0, redLeft = 200, redRight = 0;
      for (let i = 0; i < pixels.length; i += 4) {
        const x = (i/4)%200;
        if (pixels[i] === 0 && pixels[i+1] === 0 && pixels[i+2] === 255) {
          left = Math.min(left, x); right = Math.max(right, x);
        }
        if (pixels[i] === 255 && pixels[i+1] === 0 && pixels[i+2] === 0) {
          redLeft = Math.min(redLeft, x); redRight = Math.max(redRight, x);
        }
      }
      expect(left).toBeLessThan(redLeft); expect(right).toBeGreaterThan(redRight);
    }
    expect(diagnostics).toEqual([]);
  }
});

it('preserves image host transforms with identity skew and renders affine text skew', async () => {
  for (const flip of ['x', 'y']) {
    const image = (skew: boolean) => `<v:group style="width:170pt;height:170pt" coordsize="170,170"><v:shape style="left:40;top:50;width:70;height:50;rotation:90;flip:${flip}" stroked="f"><v:imagedata r:id="pic"/>${skew ? '<o:skew on="t" matrix="1,0,0,1,0,0" offset="0,0"/>' : ''}</v:shape></v:group>`;
    const control = await renderVml(image(false)), candidate = await renderVml(image(true));
    expect(candidate.pixels).toEqual(control.pixels);
    expect(candidate.diagnostics).toEqual([]);
  }
  const text = (skew: boolean) => `<v:group style="width:180pt;height:150pt" coordsize="180,150"><v:shape style="left:20;top:20;width:100;height:40" fillcolor="red" stroked="f"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="Fjp" style="font:20pt Arial"/>${skew ? '<o:skew on="t" matrix="1,0.5,0,1,0,0" offset="0,0" origin="-.5,-.5"/>' : ''}</v:shape></v:group>`;
  const plain = await renderVml(text(false)), skewed = await renderVml(text(true));
  const mass = (pixels: Uint8ClampedArray) => {
    let x = 0, y = 0, n = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      const a = (255-pixels[i+1]!)/255;
      x += ((i/4)%200+.5)*a; y += (Math.floor(i/4/200)+.5)*a; n += a;
    }
    return {x:x/n, y:y/n};
  };
  const a = mass(plain.pixels), b = mass(skewed.pixels);
  expect(b.x).toBeCloseTo(a.x, 0);
  expect(b.y-a.y).toBeCloseTo(.5*(a.x-20), 0);
  expect(skewed.diagnostics).toEqual([]);
});

it('transforms acquired arrow contours with their shafts without retracting twice', async () => {
  const draw = (matrix?: string) => renderVml(`<v:group style="width:150pt;height:100pt" coordsize="150,100"><v:line from="20,40" to="120,40" filled="f"><v:stroke color="blue" weight="6pt" endarrow="block" endarrowwidth="wide" endarrowlength="long"/>${matrix ? `<o:skew on="t" offset="0,0" matrix="${matrix}"/>` : ''}</v:line></v:group>`);
  const plain = await draw(), identity = await draw('1,0,0,1,0,0');
  expect(identity.pixels).toEqual(plain.pixels);
  expect(identity.diagnostics).toEqual([]);
  const skewed = await draw('1,.5,0,1,0,0');
  expect(skewed.diagnostics).toEqual([]);
  expect(skewed.pixels).not.toEqual(plain.pixels);
  expect(counter(skewed.pixels)(0, 0, 255)).toBeGreaterThan(200);
  // The wide/long head on a 6 pt pen is 30 pt long and 30 pt wide (5 × line
  // width, core arrow-geometry.ts). Skew maps (x, y) to (x, y + x/2 − 34) at
  // 1 px/pt, so the head becomes (120,66) (90,36) (90,66). A head regenerated
  // on the skewed shaft would be (120,66) (99.9,39.2) (86.5,66). Pixel (92, 41)
  // is ~2.1 px inside the first (≥1.4 px across the pixel) and ~7 px outside the second. The shaft
  // (y 48.6–55.4 there) misses it.
  expect([...skewed.pixels.slice((41*200+92)*4,(41*200+92)*4+3)]).toEqual([0,0,255]);
});

it('keeps page-aligned image-text brushes independent of reflected glyph masks', async () => {
  const {pixels, diagnostics} = await renderVml(`<v:group style="width:170pt;height:120pt" coordsize="170,120"><v:shape style="left:20;top:20;width:120;height:60;flip:x" stroked="f"><v:fill type="frame" r:id="pic" alignshape="f"/><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="MMMM" style="font:24pt Arial"/></v:shape></v:group>`);
  let redLeft = 0, blueRight = 0, wrong = 0;
  for (let i = 0; i < pixels.length; i += 4) {
    const x = (i/4)%200;
    if (pixels[i] === 255 && pixels[i+1] === 0 && pixels[i+2] === 0) { if (x < 80) redLeft++; else wrong++; }
    if (pixels[i] === 0 && pixels[i+1] === 0 && pixels[i+2] === 255) { if (x >= 80) blueRight++; else wrong++; }
  }
  expect(redLeft).toBeGreaterThan(100); expect(blueRight).toBeGreaterThan(100); expect(wrong).toBe(0);
  expect(diagnostics).toEqual([]);
});


it('outlines horizontal and vertical path text in independent glyph frames with physical and image pens', async () => {
  for (const vertical of [false, true]) {
    const draw = (texture = false, transformed = false) => renderVml(`<v:group style="width:180pt;height:180pt" coordsize="180,180"><v:line from="${vertical ? '70,20' : '20,70'}" to="${vertical ? '70,120' : '120,70'}" style="${transformed ? 'rotation:90;flip:x' : ''}" filled="f"><v:path textpathok="t"/><v:textpath on="t" fitpath="t" string="Fjp" style="font:24pt Arial"/><v:stroke weight="3pt" ${texture ? 'filltype="pattern" r:id="pic" color="lime" color2="lime"' : 'color="blue"'}/></v:line></v:group>`);
    const plain = await draw(), rotated = await draw(false, true), image = await draw(true, true);
    expect(counter(plain.pixels)(0, 0, 255)).toBeGreaterThan(300);
    expect(counter(rotated.pixels)(0, 0, 255)).toBeGreaterThan(300);
    expect(counter(image.pixels)(0, 255, 0)).toBeGreaterThan(100);
    expect([...plain.diagnostics, ...rotated.diagnostics, ...image.diagnostics]).toEqual([]);
  }
});

it('retains affine perspective glyph shadows for filled and outlined text', async () => {
  for (const outline of [false, true]) {
    const draw = (type: string, matrix = '1,0,0,1,0,0') => renderVml(`<v:group style="width:180pt;height:180pt" coordsize="180,180"><v:shape style="left:20;top:20;width:100;height:40;flip:x" fillcolor="red" stroked="${outline ? 't' : 'f'}" strokecolor="blue" strokeweight="3pt"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="Fjp" style="font:24pt Arial"/><v:shadow on="t" type="${type}" color="lime" offset="0,70pt" matrix="${matrix}"/></v:shape></v:group>`);
    const single = await draw('single'), perspective = await draw('perspective');
    expect(counter(perspective.pixels)(0, 255, 0)).toBeGreaterThan(200);
    // Native glyph fill and bounded acquired contours agree on area, while
    // contour edge antialiasing can differ by a pixel.
    const a = counter(single.pixels)(0, 255, 0), b = counter(perspective.pixels)(0, 255, 0);
    expect(Math.abs(a-b)).toBeLessThan(.15*a);
    const affine = await draw('perspective', '1,.3,0,1,0,0');
    expect(affine.pixels).not.toEqual(perspective.pixels);
    expect(counter(affine.pixels)(0, 255, 0)).toBeGreaterThan(200);
    expect([...perspective.diagnostics, ...affine.diagnostics]).toEqual([]);
  }
});

it('composes root and nested group skew across vector leaves, text and decorated lines', async () => {
  const skew = '<o:skew on="t" matrix="1,.5,0,1,0,0" offset="0,0" origin="-.5,-.5"/>';
  const draw = (at: 'none' | 'root' | 'nested' | 'both') => renderVml(`<v:group style="width:180pt;height:180pt" coordsize="180,180">${at === 'root' || at === 'both' ? skew : ''}<v:group style="left:20;top:20;width:140;height:140" coordsize="140,140">${at === 'nested' || at === 'both' ? skew : ''}<v:rect style="left:10;top:10;width:40;height:20" fillcolor="red" stroked="f"/><v:line from="10,60" to="100,60" filled="f"><v:stroke color="blue" weight="3pt" endarrow="block"/></v:line><v:shape style="left:10;top:90;width:100;height:30" fillcolor="lime" stroked="f"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" string="Fjp" style="font:20pt Arial"/></v:shape></v:group></v:group>`);
  const plain = await draw('none'), root = await draw('root'), nested = await draw('nested'), both = await draw('both');
  const at = (p: Uint8ClampedArray, x: number, y: number) => [...p.slice((y*200+x)*4, (y*200+x)*4+3)];
  expect(at(plain.pixels, 60, 35)).toEqual([255,0,0]);
  expect(at(root.pixels, 60, 35)).toEqual([255,255,255]);
  expect(at(root.pixels, 60, 65)).toEqual([255,0,0]);
  expect(at(nested.pixels, 60, 55)).toEqual([255,0,0]);
  expect(at(both.pixels, 60, 85)).toEqual([255,0,0]);
  for (const result of [root, nested, both]) {
    expect(counter(result.pixels)(0,0,255)).toBeGreaterThan(100);
    expect(counter(result.pixels)(0,255,0)).toBeGreaterThan(100);
    expect(result.diagnostics).toEqual([]);
  }
});


it('applies group skew after root rotation while picture data keeps its host frame', async () => {
  const skew = '<o:skew on="t" matrix="1,.5,0,1,0,0" offset="0,0" origin="-.5,-.5"/>';
  const rotated = await renderVml(`<v:group style="width:160pt;height:160pt;rotation:90;flip:x" coordsize="160,160">${skew}<v:rect style="left:30;top:30;width:40;height:20" fillcolor="red" stroked="f"/></v:group>`);
  // Host rotation/reflection maps the rect to x=110..130,y=90..130;
  // only then does group y += x/2. The interior starts at y=145 at x=110.
  const at = (p: Uint8ClampedArray, x: number, y: number) => [...p.slice((y*200+x)*4,(y*200+x)*4+3)];
  expect(at(rotated.pixels,120,160)).toEqual([255,0,0]);
  expect(at(rotated.pixels,120,110)).toEqual([255,255,255]);
  const picture = await renderVml(`<v:group style="width:160pt;height:160pt" coordsize="160,160">${skew}<v:shape style="left:20;top:20;width:100;height:60" filled="f" stroked="f"><v:imagedata r:id="pic"/></v:shape></v:group>`, 'vertical');
  // The clip shears, but the horizontal red/blue source boundary stays y=50.
  expect(at(picture.pixels,35,45)).toEqual([255,0,0]);
  expect(at(picture.pixels,35,60)).toEqual([0,0,255]);
  expect(at(picture.pixels,100,45)).toEqual([255,255,255]);
  expect([...rotated.diagnostics,...picture.diagnostics]).toEqual([]);
});


it('preserves skewed textured arrow contours and their independent single shadow', async () => {
  const {pixels, diagnostics} = await renderVml(`<v:group style="width:180pt;height:180pt" coordsize="180,180"><v:line from="20,40" to="120,40" filled="f"><v:stroke color="blue" weight="3pt" endarrow="block" endarrowwidth="wide" endarrowlength="long" filltype="tile" r:id="pic" color2="blue"/><o:skew on="t" matrix="1,.5,0,1,0,0" offset="0,0"/><v:shadow on="t" color="lime" offset="0,70pt" obscured="t"/></v:line></v:group>`);
  expect(diagnostics).toEqual([]);
  let brushInk = 0;
  for (let i=0; i<pixels.length; i+=4) brushInk += (255-pixels[i+1]!)/255;
  expect(brushInk).toBeGreaterThan(250);
  expect(counter(pixels)(0,0,255)).toBeGreaterThan(0);
  expect(counter(pixels)(255,0,0)).toBeGreaterThan(0);
  expect(counter(pixels)(0,255,0)).toBeGreaterThan(100);
});
