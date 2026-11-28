import { readFile } from 'node:fs/promises';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { resolvedBidiLevels } from '../character-spacing.js';
import { storeZip } from '../conformance/generate.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import type { DocumentLayout } from '../layout/types.js';
import { orientation } from '../run-clip.js';
import { paintLayoutPage } from '../paint/canvas-page.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { textRunsForPage } from '../text-run-projection.js';
import type { DocxDocumentModel, DocxTextRunInfo } from '../types.js';

/**
 * A paint oracle for overlay geometry: a Canvas 2D stand-in that keeps the
 * full transform and the rectangular clip, and reports where every character
 * of every `fillText` is drawn on the page. It knows nothing of the layout:
 * a glyph's box follows from the call alone (position, alignment, baseline,
 * direction, letter spacing, transform) and the stand-in's own metrics.
 *
 * Metrics: a character advances half an em, a CJK one a full em; the font
 * box is 0.8em above and 0.2em below the alphabetic baseline, and half an em
 * to each side of the `middle` baseline, in measurement as in paint. A test
 * may name pairs that are kerned when shaped in one string.
 */

export interface PageBox { left: number; top: number; right: number; bottom: number }
export interface PaintedGlyph {
  /** The character drawn (one code point). */
  character: string;
  /** Its box on the page: its advance along the text, the font box across. */
  box: PageBox;
  /** The clip it was drawn under; null when unclipped. */
  clip: PageBox | null;
}

type Matrix = [number, number, number, number, number, number];

const defaultAdvanceEm = (character: string): number => ((character.codePointAt(0) ?? 0) >= 0x2e80 ? 1 : 0.5);
let advanceEm = defaultAdvanceEm;
/** Run `body` with every character advancing `ems(character)` of the em in
 * the oracle's metrics (measurement and paint alike). */
export function withAdvances<T>(ems: (character: string) => number, body: () => T): T {
  const previous = advanceEm;
  advanceEm = ems;
  try {
    return body();
  } finally {
    advanceEm = previous;
  }
}
/** Pairs drawn closer when shaped together, by the ems their first
 * character's advance loses; none unless a test sets some. */
let kernedPairs: Readonly<Record<string, number>> = {};
/** Advance of each character of `text` shaped as one string, in ems. */
function shapedAdvancesEm(text: string): number[] {
  const characters = [...text];
  return characters.map((character, index) => advanceEm(character)
    - (kernedPairs[character + (characters[index + 1] ?? '')] ?? 0));
}
const fontPx = (font: string): number => Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '10');

function measure(font: string, text: string, baseline: string): TextMetrics {
  const size = fontPx(font);
  const width = shapedAdvancesEm(text).reduce((sum, advance) => sum + advance, 0) * size;
  const ascent = baseline === 'middle' ? size * 0.5 : size * 0.8;
  return {
    width,
    fontBoundingBoxAscent: ascent, fontBoundingBoxDescent: size - ascent,
    actualBoundingBoxAscent: ascent, actualBoundingBoxDescent: size - ascent,
    actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
  } as TextMetrics;
}

/** The measuring context layout is given: the oracle's metrics. */
export function oracleMeasureContext(): CanvasRenderingContext2D {
  const state = {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'normal', textAlign: 'left', textBaseline: 'alphabetic',
    direction: 'ltr',
    measureText: (text: string): TextMetrics => measure(state.font, text, state.textBaseline),
    save() {}, restore() {},
  };
  return state as unknown as CanvasRenderingContext2D;
}

function bounding(points: readonly (readonly [number, number])[]): PageBox {
  return {
    left: Math.min(...points.map(([x]) => x)), right: Math.max(...points.map(([x]) => x)),
    top: Math.min(...points.map(([, y]) => y)), bottom: Math.max(...points.map(([, y]) => y)),
  };
}

export function intersect(left: PageBox, right: PageBox | null): PageBox | null {
  if (!right) return left;
  const box = {
    left: Math.max(left.left, right.left), top: Math.max(left.top, right.top),
    right: Math.min(left.right, right.right), bottom: Math.min(left.bottom, right.bottom),
  };
  return box.right - box.left > 1e-9 && box.bottom - box.top > 1e-9 ? box : null;
}

/** A canvas whose 2D context records the glyphs drawn on it. */
export function glyphRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  glyphs: PaintedGlyph[];
  /** The box and clip of every resource (picture) painted on the canvas,
   * when its page is painted with `resources`. */
  pictures: { box: PageBox; clip: PageBox | null }[];
  resources: { paint(key: string, kind: string, bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>): void };
} {
  const glyphs: PaintedGlyph[] = [];
  const pictures: { box: PageBox; clip: PageBox | null }[] = [];
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  let clip: PageBox | null = null;
  let path: PageBox[] | null = [];
  const stack: { matrix: Matrix; clip: PageBox | null }[] = [];
  const state = {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto', textAlign: 'left', textBaseline: 'alphabetic',
    direction: 'ltr', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, globalAlpha: 1, lineCap: 'butt',
    lineJoin: 'miter', globalCompositeOperation: 'source-over', filter: 'none', imageSmoothingEnabled: true,
  };
  const map = (x: number, y: number): [number, number] => [
    matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5],
  ];
  const multiply = (m: Matrix): void => {
    const [a, b, c, d, e, f] = matrix;
    matrix = [a * m[0] + c * m[1], b * m[0] + d * m[1], a * m[2] + c * m[3], b * m[2] + d * m[3],
      a * m[4] + c * m[5] + e, b * m[4] + d * m[5] + f];
  };
  const methods = {
    measureText: (text: string) => measure(state.font, text, state.textBaseline),
    save() { stack.push({ matrix: [...matrix] as Matrix, clip }); },
    restore() {
      const saved = stack.pop();
      if (saved) ({ matrix, clip } = saved);
    },
    translate(x: number, y: number) { multiply([1, 0, 0, 1, x, y]); },
    scale(x: number, y: number) { multiply([x, 0, 0, y, 0, 0]); },
    rotate(angle: number) { multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]); },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) { multiply([a, b, c, d, e, f]); },
    setTransform(a: number | { a: number; b: number; c: number; d: number; e: number; f: number } = 1,
      b = 0, c = 0, d = 1, e = 0, f = 0) {
      matrix = typeof a === 'object' ? [a.a, a.b, a.c, a.d, a.e, a.f] : [a, b, c, d, e, f];
    },
    resetTransform() { matrix = [1, 0, 0, 1, 0, 0]; },
    getTransform() {
      const [a, b, c, d, e, f] = matrix;
      return { a, b, c, d, e, f };
    },
    beginPath() { path = []; },
    rect(x: number, y: number, width: number, height: number) {
      path?.push(bounding([map(x, y), map(x + width, y), map(x + width, y + height), map(x, y + height)]));
    },
    // A path that is not rectangles is not a clip this oracle can state.
    moveTo() { path = null; },
    clip() {
      if (!path || path.length !== 1) return;
      const region = path[0]!;
      clip = clip ? intersect(clip, region) ?? { left: 0, top: 0, right: 0, bottom: 0 } : region;
    },
    fillText(text: string, x: number, y: number) {
      const size = fontPx(state.font);
      const spacing = Number.parseFloat(state.letterSpacing) || 0;
      const characters = [...text];
      const advances = shapedAdvancesEm(text).map((advance) => advance * size + spacing);
      const total = advances.reduce((sum, advance) => sum + advance, 0);
      // Visual order: UAX#9 levels in the paint direction, then L2.
      const levels = resolvedBidiLevels(text, state.direction as 'ltr' | 'rtl');
      let offset = 0;
      const level = characters.map((character) => {
        const own = levels[offset] ?? 0;
        offset += character.length;
        return own;
      });
      let order = characters.map((_, index) => index);
      const highest = Math.max(0, ...level);
      for (let current = highest; current >= 1; current -= 1) {
        const next: number[] = [];
        let cursor = 0;
        while (cursor < order.length) {
          let end = cursor;
          while (end < order.length && level[order[end]!]! >= current) end += 1;
          if (end === cursor) next.push(order[cursor++]!);
          else {
            for (let member = end - 1; member >= cursor; member -= 1) next.push(order[member]!);
            cursor = end;
          }
        }
        order = next;
      }
      let pen = state.textAlign === 'center' ? x - total / 2 : state.textAlign === 'right' ? x - total : x;
      const top = state.textBaseline === 'middle' ? y - size / 2 : y - size * 0.8;
      for (const index of order) {
        const advance = advances[index]! - spacing;
        // Letter spacing follows a character in its own reading direction.
        const start = level[index]! % 2 === 1 ? pen + spacing : pen;
        glyphs.push({
          character: characters[index]!,
          box: bounding([map(start, top), map(start + advance, top), map(start + advance, top + size), map(start, top + size)]),
          clip,
        });
        pen += advances[index]!;
      }
    },
    createLinearGradient: () => ({ addColorStop() {} }),
    createRadialGradient: () => ({ addColorStop() {} }),
    createPattern: () => null,
  };
  const ctx = new Proxy(state as Record<string, unknown>, {
    get(target, key: string) {
      if (key in methods) return (methods as Record<string, unknown>)[key];
      if (key in target) return target[key];
      // Everything else a page paints (rules, fills, images) draws no text.
      return () => undefined;
    },
    set(target, key: string, value) {
      target[key] = value;
      return true;
    },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  const resources = {
    paint(_key: string, _kind: string, bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>) {
      pictures.push({
        box: bounding([map(bounds.xPt, bounds.yPt), map(bounds.xPt + bounds.widthPt, bounds.yPt),
          map(bounds.xPt + bounds.widthPt, bounds.yPt + bounds.heightPt), map(bounds.xPt, bounds.yPt + bounds.heightPt)]),
        clip,
      });
    },
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, glyphs, pictures, resources };
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DRAWING = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** A one-pixel PNG, and the package entries that make it `rIdP` of the main
 * document part. */
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='),
  (character) => character.charCodeAt(0));
export const PICTURE_PART = {
  parts: { 'word/media/picture.png': PNG },
  relationships: '<Relationship Id="rIdP" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
    + 'Target="media/picture.png"/>',
  contentTypes: '<Default Extension="png" ContentType="image/png"/>',
} as const;

/** A run holding an inline DrawingML picture of `widthPt` by `heightPt`,
 * turned clockwise by `rotationDeg`. */
export function inlinePictureRun(widthPt: number, heightPt: number, properties = '', rotationDeg = 0): string {
  const extent = `cx="${widthPt * 12700}" cy="${heightPt * 12700}"`;
  return `<w:r><w:rPr>${properties}</w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${extent}/>`
    + '<wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip r:embed="rIdP"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + `<pic:spPr><a:xfrm${rotationDeg ? ` rot="${rotationDeg * 60000}"` : ''}><a:off x="0" y="0"/><a:ext ${extent}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

let ready: Promise<unknown> | undefined;
export function initParser(): Promise<unknown> {
  return ready ??= readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url))
    .then((wasmBytes) => init({ module_or_path: wasmBytes }));
}

/** Lay out a one-section document of `body` (compatibility mode 15; text
 * area 468pt wide from (72, 72), or a tbRl section) with the oracle's
 * metrics, and paint its first page on the oracle. */
export async function paintedDocument(body: string, options: Readonly<{
  vertical?: boolean;
  section?: string;
  /** Kerned pairs, by the ems the first character's advance loses. */
  kerning?: Readonly<Record<string, number>>;
  /** Further package parts by name, with the relationships of the main
   * document part to them and their content-type declarations. */
  parts?: Readonly<Record<string, string | Uint8Array>>;
  relationships?: string;
  contentTypes?: string;
}> = {}): Promise<{
  layout: DocumentLayout;
  model: DocxDocumentModel;
  runs: DocxTextRunInfo[];
  glyphs: PaintedGlyph[];
  pictures: { box: PageBox; clip: PageBox | null }[];
}> {
  const section = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    + `w:header="0" w:footer="0" w:gutter="0"/>${options.vertical ? '<w:textDirection w:val="tbRl"/>' : ''}${options.section ?? ''}`;
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + `${options.contentTypes ?? ''}</Types>`)],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdT" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      + `${options.relationships ?? ''}</Relationships>`)],
    ['word/settings.xml', encode(`<w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" `
      + 'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')],
    ['word/document.xml', encode(`<w:document ${W} ${DRAWING}><w:body>${body}<w:sectPr>${section}</w:sectPr></w:body></w:document>`)],
    ...Object.entries(options.parts ?? {}).map(([name, content]) => (
      [name, typeof content === 'string' ? encode(content) : content] as [string, Uint8Array])),
  ]));
  await initParser();
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  kernedPairs = options.kerning ?? {};
  try {
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: oracleMeasureContext() }), { currentDateMs: 0 });
    const { canvas, glyphs, pictures, resources } = glyphRecordingCanvas();
    await paintLayoutPage(layout, 0, canvas, { dpr: 1, scale: 1 }, resources as never);
    return { layout, model, runs: textRunsForPage(layout, 0, { scale: 1 }), glyphs, pictures };
  } finally {
    kernedPairs = {};
  }
}

/** The page box of a rect placed like DocxTextRunInfo.highlightBounds whose
 * orientation is a quarter turn or a reflection. */
export function pageBoxOf(box: Readonly<{ x: number; y: number; width: number; height: number; transform?: string }>): PageBox {
  const [a, b, c, d] = orientation(box.transform);
  return bounding([[0, 0], [box.width, 0], [box.width, box.height], [0, box.height]].map(([u, v]) => (
    [box.x + a * u! + c * v!, box.y + b * u! + d * v!] as const)));
}

/** A stand-in DOM element that keeps what the overlay layers write. */
export interface FakeElement {
  style: { cssText: string; transform: string };
  dataset: Record<string, string>;
  children: FakeElement[];
  ownerDocument: { createElement(): FakeElement };
  parentElement: null;
  innerHTML: string;
  appendChild(child: FakeElement): void;
}

export function fakeElement(): FakeElement {
  const ownerDocument = { createElement: fakeElement };
  return {
    style: { cssText: '', transform: '' }, dataset: {}, children: [], ownerDocument, parentElement: null, innerHTML: '',
    appendChild(child) { this.children.push(child); },
  };
}

/** Run `build` with `document.createElement` making stand-in elements. */
export function withFakeDocument<T>(build: () => T): T {
  const scope = globalThis as { document?: unknown };
  const previous = scope.document;
  scope.document = { createElement: fakeElement };
  try {
    return build();
  } finally {
    scope.document = previous;
  }
}

/** The page boxes of the leaf elements of an overlay `layer` built for a
 * page `width` by `height`, each cut by the clip containers it lies in:
 * percentages resolve against the parent, orientations compose about the
 * top-left corner (quarter turns and reflections). Boxes cut away whole are
 * omitted. */
export function layerBoxes(layer: FakeElement, width: number, height: number): PageBox[] {
  type Frame = { matrix: Matrix; width: number; height: number; clip: PageBox | null };
  const boxes: PageBox[] = [];
  const visit = (element: FakeElement, parent: Frame): void => {
    const css = element.style.cssText;
    const percent = (name: string, basis: number): number => (
      Number(new RegExp(`(?:^|;)${name}:([-\\d.e]+)%`, 'u').exec(css)?.[1] ?? 0) * basis / 100);
    const left = percent('left', parent.width);
    const top = percent('top', parent.height);
    const own = { width: percent('width', parent.width), height: percent('height', parent.height) };
    const transform = element.style.transform || /transform:([^;]+);/u.exec(css)?.[1] || undefined;
    const [a, b, c, d] = orientation(transform);
    const [pa, pb, pc, pd, pe, pf] = parent.matrix;
    const matrix: Matrix = [pa * a + pc * b, pb * a + pd * b, pa * c + pc * d, pb * c + pd * d,
      pa * left + pc * top + pe, pb * left + pd * top + pf];
    const corner = (u: number, v: number): [number, number] => [
      matrix[0] * u + matrix[2] * v + matrix[4], matrix[1] * u + matrix[3] * v + matrix[5]];
    const box = bounding([corner(0, 0), corner(own.width, 0), corner(own.width, own.height), corner(0, own.height)]);
    if (/overflow:/u.test(css)) {
      const clip = parent.clip ? intersect(box, parent.clip) ?? { left: 0, top: 0, right: 0, bottom: 0 } : box;
      for (const child of element.children) visit(child, { matrix, ...own, clip });
      return;
    }
    const visible = intersect(box, parent.clip);
    if (visible) boxes.push(visible);
  };
  for (const child of layer.children) visit(child, { matrix: [1, 0, 0, 1, 0, 0], width, height, clip: null });
  return boxes;
}
