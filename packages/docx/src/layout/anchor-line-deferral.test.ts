/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import type { DocxDocumentModel } from '../types.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { layoutDocumentProgressively, type ProgressiveLayoutPreview } from './progressive.js';
import {
  anchorLineDeferralApplies,
  createAnchorLineDeferralProof,
  hostFloatBottomMarginPlan,
  serializeAnchorInput,
  type PageAnchorInputEvent,
} from './anchor-line-deferral.js';
import type { DocumentLayout, LayoutPage } from './types.js';

// Every anchor-solver pass, counted at the solver's convergence boundary.
const solver = vi.hoisted(() => ({ passes: 0 }));
vi.mock('./convergence.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./convergence.js')>();
  return {
    ...actual,
    convergeExactStateSteps: ((options) => actual.convergeExactStateSteps({
      ...options,
      step: function* countedStep(previous, pass) {
        solver.passes += 1;
        return yield* options.step(previous, pass);
      },
    })) as typeof actual.convergeExactStateSteps,
  };
});

// Word's resolution of a page-owned anchor whose own exclusion pushes its
// anchor line off the page (issue #1615; rule in anchor-line-deferral.ts).
// Each case mirrors a Word for Mac control and asserts Word's placement.
//
// Pagination is metric independent: every paragraph line is an exact 24 pt
// line (§17.3.1.33) measured by an injected linear-metric context, so a US
// Letter page with 1 in margins holds 27 lines. Anchored pictures are
// 200 x 300 pt unless stated, with zero wrap distances.

const EMU = 12700;
const PNG = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC',
  'base64',
));
const encoder = new TextEncoder();

type Anchor = Readonly<{
  wrap?: 'topAndBottom' | 'square' | 'tight';
  relativeFrom?: 'margin' | 'page' | 'paragraph';
  xPt?: number;
  yPt?: number;
  widthPt?: number;
  heightPt?: number;
  allowOverlap?: boolean;
  distBPt?: number;
  /** §20.4.2.6 effectExtent top and bottom (pt). */
  effectTPt?: number;
  effectBPt?: number;
  /** Line after which this anchor's run sits (default `anchorAfterLine`). */
  line?: number;
}>;

type Paragraph = Readonly<{
  anchors?: readonly Anchor[];
  /** Additional lines forced with w:br; the anchor run follows `anchorAfterLine`. */
  lines?: number;
  anchorAfterLine?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  /** An automatic single line (§17.3.1.33) instead of the exact 24pt line. */
  auto?: boolean;
}>;

function anchorXml(id: number, anchor: Anchor): string {
  const width = Math.round((anchor.widthPt ?? 200) * EMU);
  const height = Math.round((anchor.heightPt ?? 300) * EMU);
  const vertical = anchor.relativeFrom ?? 'margin';
  const horizontal = vertical === 'page' ? 'page' : 'margin';
  const xPt = anchor.xPt ?? (horizontal === 'page' ? 72 : 0);
  const wrap = anchor.wrap === 'square'
    ? '<wp:wrapSquare wrapText="bothSides"/>'
    : anchor.wrap === 'tight'
      ? '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>'
        + '<wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>'
        + '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
      : '<wp:wrapTopAndBottom/>';
  return '<w:r><w:drawing>'
    + `<wp:anchor distT="0" distB="${Math.round((anchor.distBPt ?? 0) * EMU)}" distL="0" distR="0" simplePos="0" `
    + `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" `
    + `allowOverlap="${anchor.allowOverlap === false ? 0 : 1}">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="${horizontal}"><wp:posOffset>${Math.round(xPt * EMU)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="${vertical}"><wp:posOffset>${Math.round((anchor.yPt ?? 0) * EMU)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="${Math.round((anchor.effectTPt ?? 0) * EMU)}" r="0" b="${Math.round((anchor.effectBPt ?? 0) * EMU)}"/>${wrap}`
    + `<wp:docPr id="${id}" name="Anchor ${id}"/>`
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="a${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>`
    + '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

function docx(
  paragraphCount: number,
  special: Readonly<Record<number, Paragraph>>,
  columns = 1,
  compatibilityMode?: number,
  /** A §17.6.5 lines grid with this pitch (pt). */
  gridPitchPt?: number,
): Uint8Array {
  let nextId = 1;
  const body = Array.from({ length: paragraphCount }, (_, index) => {
    const spec = special[index] ?? {};
    const label = `P${String(index).padStart(3, '0')}`;
    const texts = Array.from({ length: spec.lines ?? 1 }, (_, line) => (
      `<w:r><w:t>${line === 0 ? label : `${label}.${line + 1}`}</w:t></w:r>`
    ));
    const after = spec.anchorAfterLine ?? 0;
    const anchors = (spec.anchors ?? []).map((anchor) => [anchor.line ?? after, anchorXml(nextId++, anchor)] as const);
    const runs = texts.flatMap((text, line) => [
      ...(line > 0 ? ['<w:r><w:br/></w:r>'] : []),
      text,
      ...anchors.filter(([anchorLine]) => anchorLine === line).map(([, xml]) => xml),
    ]).join('');
    return '<w:p><w:pPr>' + (spec.keepNext ? '<w:keepNext/>' : '')
      + (spec.keepLines ? '<w:keepLines/>' : '')
      + `<w:widowControl w:val="${spec.widowControl ? 1 : 0}"/>`
      // On a grid, automatic single lines snap to one cell each.
      + (gridPitchPt === undefined && spec.auto !== true
        ? '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>'
        : '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>')
      + `${runs}</w:p>`;
  }).join('');
  const document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'
    + body
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + (columns > 1 ? `<w:cols w:num="${columns}" w:space="480"/>` : '<w:cols w:space="720"/>')
    + (gridPitchPt === undefined ? '' : `<w:docGrid w:type="lines" w:linePitch="${Math.round(gridPitchPt * 20)}"/>`)
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (compatibilityMode === undefined ? ''
        : '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>')
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>'
      + (compatibilityMode === undefined ? ''
        : '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>')
      + '</Relationships>')],
    ['word/document.xml', encoder.encode(document)],
    ...(compatibilityMode === undefined ? [] : [['word/settings.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>'
      + `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/>`
      + '</w:compat></w:settings>')] as const]),
    ['word/media/a.png', PNG],
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function parse(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    return normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
  } finally {
    archive.free();
  }
}

function layout(bytes: Uint8Array): DocumentLayout {
  const model = parse(bytes);
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
    currentDateMs: 0,
  });
}

/** Page (1-based) and top of each line labelled Pnnn or Pnnn.k. */
function lines(result: DocumentLayout): Map<string, Readonly<{ page: number; yPt: number; xPt: number }>> {
  const found = new Map<string, Readonly<{ page: number; yPt: number; xPt: number }>>();
  for (const page of result.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      for (const line of node.lines) {
        const text = line.placements
          .map((placement) => placement.kind === 'text' ? placement.text : '')
          .join('');
        const label = /P\d{3}(?:\.\d)?/.exec(text)?.[0];
        if (label) {
          found.set(label, Object.freeze({
            page: page.pageIndex + 1, yPt: line.bounds.yPt, xPt: line.bounds.xPt,
          }));
        }
      }
    }
  }
  return found;
}

/** Page (1-based) and top-left of each anchored picture, by anchor paragraph. */
function pictures(result: DocumentLayout) {
  const found = new Map<number, Array<Readonly<{ page: number; xPt: number; yPt: number }>>>();
  for (const page of result.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      for (const drawing of node.drawings) {
        const list = found.get(node.source.path[0]!) ?? [];
        list.push(Object.freeze({
          page: page.pageIndex + 1, xPt: drawing.flowBounds.xPt, yPt: drawing.flowBounds.yPt,
        }));
        found.set(node.source.path[0]!, list);
      }
    }
  }
  return found;
}

function pageRange(result: DocumentLayout, pageNumber: number): readonly number[] {
  const indexes = result.pages[pageNumber - 1]!.layers.body
    .filter((node) => node.kind === 'paragraph')
    .map((node) => node.source.path[0]!);
  return [indexes[0]!, indexes.at(-1)!];
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('page-owned anchor whose own exclusion pushes its anchor line off the page', () => {
  it('moves the anchor line and picture to the next page and keeps the page unwrapped', () => {
    // Word control: the anchor paragraph is the last line of page 1. Word keeps
    // all 26 earlier lines on page 1 without wrapping them below the picture
    // (only 14 would fit below it) and starts page 2 with picture and line.
    const result = layout(docx(27, { 26: { anchors: [{}] } }));
    expect(result.pages).toHaveLength(2);
    expect(pageRange(result, 1)).toEqual([0, 25]);
    expect(lines(result).get('P000')).toMatchObject({ page: 1, yPt: 72 });
    expect(pictures(result).get(26)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
    expect(lines(result).get('P026')).toMatchObject({ page: 2, yPt: 372 });
  });

  it('keeps the picture on the page when its line still fits below it', () => {
    // Word control: the 14th line is the last that fits below the picture;
    // the earlier lines wrap below it and the picture stays on page 1.
    const result = layout(docx(15, { 13: { anchors: [{}] } }));
    expect(pictures(result).get(13)).toEqual([{ page: 1, xPt: 72, yPt: 72 }]);
    expect(lines(result).get('P000')).toMatchObject({ page: 1, yPt: 372 });
    expect(lines(result).get('P013')).toMatchObject({ page: 1, yPt: 684 });
    expect(lines(result).get('P014')).toMatchObject({ page: 2, yPt: 72 });
  });

  it('splits a paragraph above an anchor run on a later line', () => {
    // Word control: the anchor run follows the third line of P013. The first
    // two lines stay on page 1 without the picture; the third line starts
    // page 2 below it.
    const result = layout(docx(16, { 13: { anchors: [{}], lines: 3, anchorAfterLine: 2 } }));
    const placed = lines(result);
    expect(placed.get('P013')).toMatchObject({ page: 1, yPt: 384 });
    expect(placed.get('P013.2')).toMatchObject({ page: 1, yPt: 408 });
    expect(placed.get('P013.3')).toMatchObject({ page: 2, yPt: 372 });
    expect(placed.get('P014')).toMatchObject({ page: 2, yPt: 396 });
    expect(pictures(result).get(13)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
  });

  it('moves a keepNext predecessor with the deferred anchor paragraph', () => {
    const result = layout(docx(16, { 13: { keepNext: true }, 14: { anchors: [{}] } }));
    expect(pageRange(result, 1)).toEqual([0, 12]);
    expect(lines(result).get('P013')).toMatchObject({ page: 2, yPt: 372 });
    expect(lines(result).get('P014')).toMatchObject({ page: 2, yPt: 396 });
    expect(pictures(result).get(14)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
  });

  it('tests a later anchor with only the anchors accepted before it', () => {
    // Word control: P003's picture fits; adding P010's partially overlapping
    // picture would push P010. Page 1 keeps P003's wrap only and ends at P009.
    const result = layout(docx(30, {
      3: { anchors: [{}] },
      10: { anchors: [{ xPt: 100, yPt: 100 }] },
    }));
    expect(pageRange(result, 1)).toEqual([0, 9]);
    expect(lines(result).get('P000')).toMatchObject({ page: 1, yPt: 372 });
    expect(pictures(result).get(3)).toEqual([{ page: 1, xPt: 72, yPt: 72 }]);
    expect(pictures(result).get(10)).toEqual([{ page: 2, xPt: 172, yPt: 172 }]);
    expect(lines(result).get('P010')).toMatchObject({ page: 2, yPt: 72 });
  });

  it('keeps a picture on its page when its line moves to the next column', () => {
    // Word control: two 222 pt columns. Column 1 stays wrapped below the
    // picture and the anchor paragraph starts column 2 of the same page.
    const result = layout(docx(80, { 14: { anchors: [{}] } }, 2));
    expect(pictures(result).get(14)).toEqual([{ page: 1, xPt: 72, yPt: 72 }]);
    expect(lines(result).get('P000')).toMatchObject({ page: 1, yPt: 372, xPt: 72 });
    expect(lines(result).get('P014')).toMatchObject({ page: 1, yPt: 72, xPt: 318 });
  });

  it('resolves the issue #1615 document with Word pagination', () => {
    const result = layout(docx(400, Object.fromEntries(
      [30, 31, 90, 200, 201, 350].map((index) => [index, { anchors: [{}] }]),
    )));
    // Word: 18 pages; P090 and P200 start pages 5 and 10 below their pictures.
    expect(result.pages).toHaveLength(18);
    expect(pageRange(result, 4)).toEqual([68, 89]);
    expect(pageRange(result, 9)).toEqual([185, 199]);
    expect(lines(result).get('P090')).toMatchObject({ page: 5, yPt: 372 });
    expect(lines(result).get('P200')).toMatchObject({ page: 10, yPt: 372 });
    expect(lines(result).get('P350')).toMatchObject({ page: 16, yPt: 396 });
  });

  it('publishes only pages the converged layout keeps', async () => {
    const model = parse(docx(400, Object.fromEntries(
      [30, 31, 90, 200, 201, 350].map((index) => [index, { anchors: [{}] }]),
    )));
    const source = layoutSourceStore(model);
    const previews: ProgressiveLayoutPreview[] = [];
    const final = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: measureContext() }),
      normalizeLayoutOptions(undefined, 0),
      { onPreview: (preview) => { previews.push(preview); } },
    );
    expect(previews.length).toBeGreaterThan(1);
    const pageFingerprint = (page: LayoutPage) => layoutFingerprint({ pages: [page], diagnostics: [] });
    for (const preview of previews) {
      expect(preview.layout.pages.length).toBeLessThanOrEqual(final.pages.length);
      preview.layout.pages.forEach((page, index) => {
        expect(pageFingerprint(page as LayoutPage)).toBe(pageFingerprint(final.pages[index] as LayoutPage));
      });
    }
  }, 300_000);
});

describe('anchor-line cycle recovery (library policy)', () => {
  // LIBRARY_ANCHOR_LINE_CYCLE_RECOVERY: where the contextual anchor-line
  // tests above alternate, the solver must still converge. These cases pin
  // the library's deterministic result, not a Word measurement.
  it('settles a later anchor whose test also pushes the lines accepted before it', () => {
    // P003's picture fits on page 1. Registering P010's 600 pt picture as well
    // leaves room for two lines only, so that test pushes P003's line off the
    // page too and the contextual tests alternate. Recovery keeps P003's wrap
    // on page 1, which ends at P009; P010 starts page 2 below its picture.
    const result = layout(docx(30, {
      3: { anchors: [{}] },
      10: { anchors: [{ heightPt: 600 }] },
    }));
    expect(pageRange(result, 1)).toEqual([0, 9]);
    expect(lines(result).get('P000')).toMatchObject({ page: 1, yPt: 372 });
    expect(pictures(result).get(3)).toEqual([{ page: 1, xPt: 72, yPt: 72 }]);
    expect(pictures(result).get(10)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
    expect(lines(result).get('P010')).toMatchObject({ page: 2, yPt: 672 });
  });

  it('settles later anchors in source order when one test pushes several lines', () => {
    // As above, plus P011's 100 pt picture, which P010's picture pushes off
    // page 1 too. Recovery keeps P003's wrap on page 1 and moves P010 and
    // P011 to page 2.
    const result = layout(docx(30, {
      3: { anchors: [{}] },
      10: { anchors: [{ heightPt: 600 }] },
      11: { anchors: [{ heightPt: 100 }] },
    }));
    expect(pageRange(result, 1)).toEqual([0, 9]);
    expect(pictures(result).get(3)).toEqual([{ page: 1, xPt: 72, yPt: 72 }]);
    expect(pictures(result).get(10)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
    expect(pictures(result).get(11)).toEqual([{ page: 2, xPt: 72, yPt: 72 }]);
    expect(lines(result).get('P010')).toMatchObject({ page: 2, yPt: 672 });
    expect(lines(result).get('P011')).toMatchObject({ page: 2, yPt: 696 });
  });

  it('settles a deferral that keepNext extends to an earlier anchor paragraph', () => {
    // Page 2 opens with P027. P028's square picture fits beside its line;
    // adding P029's full-width picture pushes P027-P029 below it. Deferring
    // P029 also moves P028 (keepNext, §17.3.1.15), whose picture follows it,
    // so the contextual tests alternate; recovery moves the kept pair and
    // both pictures to page 3.
    const result = layout(docx(30, {
      28: { keepNext: true, anchors: [{ wrap: 'square', widthPt: 200, heightPt: 600 }] },
      29: { anchors: [{ widthPt: 468, heightPt: 600 }] },
    }));
    const placed = lines(result);
    expect(pageRange(result, 2)).toEqual([27, 27]);
    expect(placed.get('P028')).toMatchObject({ page: 3, yPt: 672 });
    expect(placed.get('P029')).toMatchObject({ page: 3, yPt: 696 });
    expect(pictures(result).get(28)).toEqual([{ page: 3, xPt: 72, yPt: 72 }]);
    expect(pictures(result).get(29)).toEqual([{ page: 3, xPt: 72, yPt: 72 }]);
  });

  it('converges with every line and picture placed once for generated anchor sets', () => {
    // Deterministic generated documents: 2-5 anchors of varied heights and
    // wraps at varied paragraphs, many of whose later tests push earlier
    // lines too, with keepNext/keepLines chains, widow control and
    // multi-line anchor paragraphs around them.
    let seed = 0x2306;
    const next = (bound: number) => {
      seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
      return (seed >>> 8) % bound;
    };
    const heights = [100, 300, 450, 600, 640];
    const paragraphs = 45;
    for (let sample = 0; sample < 32; sample += 1) {
      const special: Record<number, Paragraph> = {};
      const count = 2 + next(4);
      for (let index = 0; index < count; index += 1) {
        const paragraph = 20 + next(paragraphs - 20);
        special[paragraph] = {
          ...special[paragraph],
          anchors: [...(special[paragraph]?.anchors ?? []), {
            heightPt: heights[next(heights.length)]!,
            wrap: next(2) === 0 ? 'square' : 'topAndBottom',
            widthPt: next(2) === 0 ? 200 : 468,
          }],
        };
        // A keepNext chain of up to three paragraphs ending at the anchor.
        for (let chain = next(4); chain > 0 && paragraph - chain >= 0; chain -= 1) {
          special[paragraph - chain] = { ...special[paragraph - chain], keepNext: true };
        }
        if (next(3) === 0) {
          special[paragraph] = {
            ...special[paragraph], lines: 3, anchorAfterLine: next(3),
            keepLines: next(2) === 0, widowControl: next(2) === 0,
          };
        }
      }
      const label = JSON.stringify(special);
      let result: DocumentLayout;
      try {
        result = layout(docx(paragraphs, special));
      } catch (error) {
        throw new Error(`${label}: ${String(error)}`);
      }
      const placed = new Map<string, number>();
      for (const page of result.pages) {
        for (const node of page.layers.body) {
          if (node.kind !== 'paragraph') continue;
          for (const line of node.lines) {
            const text = line.placements.map((p) => p.kind === 'text' ? p.text : '').join('');
            const name = /P\d{3}(?:\.\d)?/.exec(text)?.[0];
            if (name) placed.set(name, (placed.get(name) ?? 0) + 1);
          }
        }
      }
      const expected = Array.from({ length: paragraphs }, (_, index) => {
        const name = `P${String(index).padStart(3, '0')}`;
        return Array.from({ length: special[index]?.lines ?? 1 }, (_, line) => (
          line === 0 ? name : `${name}.${line + 1}`
        ));
      }).flat();
      expect([...placed.keys()].sort(), label).toEqual(expected.sort());
      expect([...placed.values()].every((times) => times === 1), label).toBe(true);
      const drawn = pictures(result);
      for (const [paragraph, spec] of Object.entries(special)) {
        expect(drawn.get(Number(paragraph))?.length ?? 0, label).toBe(spec.anchors?.length ?? 0);
      }
      expect([...drawn.values()].flat().length, label).toBe(count);
    }
  }, 120_000);

  it('bounds the solver work of many independent anchor-line cycles', () => {
    // Eight copies of the P003/P010 cycle above, one per 30 paragraphs. Each
    // copy needs its own recovery; the whole run must stay within the
    // solver's ordinary operational pass budget for its anchors (16 + 4 per
    // page-owned anchor), not a budget per settlement.
    const blocks = 8;
    const special: Record<number, Paragraph> = {};
    for (let block = 0; block < blocks; block += 1) {
      special[block * 30 + 3] = { anchors: [{}] };
      special[block * 30 + 10] = { anchors: [{ heightPt: 600 }] };
    }
    solver.passes = 0;
    const result = layout(docx(blocks * 30, special));
    expect(solver.passes).toBeLessThanOrEqual(16 + 4 * 2 * blocks);
    const drawn = pictures(result);
    const placed = lines(result);
    for (let block = 0; block < blocks; block += 1) {
      const first = block * 30;
      expect(drawn.get(first + 3)).toHaveLength(1);
      expect(drawn.get(first + 10)).toHaveLength(1);
      // P010's picture and line share the page after P003's.
      expect(drawn.get(first + 10)![0]!.page).toBe(drawn.get(first + 3)![0]!.page + 1);
      expect(placed.get(`P${String(first + 10).padStart(3, '0')}`)?.page).toBe(drawn.get(first + 10)![0]!.page);
    }
  }, 120_000);

});

describe('anchor-line deferral proof', () => {
  // A proof is the proving pass's exact page context: a pass may apply it
  // only when its reads up to the anchor line are that context without the
  // deferred anchors, with no read missing and none added.
  const source = { story: 'body', storyInstance: 'body', path: [5] } as const;
  const identity = (anchors: readonly { occurrenceId: string }[]) => (
    anchors.map((anchor) => anchor.occurrenceId).join(',')
  );
  const prescan = (pageIndex: number, keys: readonly string[]): PageAnchorInputEvent => ({
    kind: 'prescan',
    pageIndex,
    flowDomainId: `page:${pageIndex}`,
    anchors: keys.map((occurrenceId) => ({ kind: 'drawing', occurrenceId, paragraphSource: source })),
  });
  const table = (pageIndex: number): PageAnchorInputEvent => ({
    kind: 'page-owned-table', pageIndex, key: 'table:body:3', floor: undefined,
  });
  const reads = (events: readonly PageAnchorInputEvent[]) => events.map(
    (event) => serializeAnchorInput(event, identity),
  );
  const prove = (events: readonly PageAnchorInputEvent[]) => createAnchorLineDeferralProof(
    ['a'], 1, events, reads(events), 1, identity,
  );

  it('applies to the same page context without the deferred anchor', () => {
    const proof = prove([prescan(0, []), prescan(1, ['a']), prescan(2, ['a'])]);
    expect(anchorLineDeferralApplies(proof, ['a'], 1, {
      reads: reads([prescan(0, []), prescan(1, [])]),
    })).toBe(true);
  });

  it('rejects a context with an additional read before the anchor line', () => {
    // Without the anchor, a page-owned table is reached on the page first.
    const proof = prove([prescan(0, []), prescan(1, ['a']), prescan(2, ['a'])]);
    expect(anchorLineDeferralApplies(proof, ['a'], 1, {
      reads: reads([prescan(0, []), prescan(1, []), table(1)]),
    })).toBe(false);
  });

  it('rejects a context missing a read of the proving page', () => {
    const proof = prove([prescan(0, []), prescan(1, ['a']), table(1), prescan(2, ['a'])]);
    expect(anchorLineDeferralApplies(proof, ['a'], 1, {
      reads: reads([prescan(0, []), prescan(1, [])]),
    })).toBe(false);
    expect(anchorLineDeferralApplies(proof, ['a'], 1, {
      reads: reads([prescan(0, []), prescan(1, []), table(1)]),
    })).toBe(true);
  });
});

describe('WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL', () => {
  // 27 automatic single lines, one 24pt lines-grid cell each, fill the 648pt
  // US Letter body (72-720pt); without the grid they are exact 24pt lines. A
  // 30pt paragraph-relative float anchored in P025 (672-696pt) at +40pt would
  // end at 742pt, below the body; risen to end at 720pt it would start above
  // the line's bottom.
  const paginate = (special: Readonly<Record<number, Paragraph>>, options: Readonly<{
    mode?: number; grid?: boolean; columns?: number; count?: number;
  }> = {}) => lines(layout(docx(options.count ?? 27, special, options.columns ?? 1,
    options.mode ?? 15, options.grid === false ? undefined : 24)));
  const crossing = (mode: number, yPt: number, wrap: 'square' | 'topAndBottom' | 'tight' = 'square',
    grid = true) => paginate({
    25: { anchors: [{ wrap, relativeFrom: 'paragraph', yPt, widthPt: 100, heightPt: 30 }] },
  }, { mode, grid });
  it.each(['square', 'topAndBottom', 'tight'] as const)('moves the anchor line of a crossing %s float to the next page in mode 15', (wrap) => {
    const placed = crossing(15, 40, wrap);
    expect(placed.get('P025')).toMatchObject({ page: 2, yPt: 72 });
    expect(placed.get('P024')?.page).toBe(1);
  });
  it('keeps the anchor line when the float ends inside the body', () => {
    expect(crossing(15, 10).get('P025')?.page).toBe(1);
  });
  it('lets the float overhang the margin in mode 14', () => {
    expect(crossing(14, 40).get('P025')?.page).toBe(1);
  });
  it('moves the anchor line of exact lines without a lines grid', () => {
    expect(crossing(15, 40, 'square', false).get('P025')).toMatchObject({ page: 2, yPt: 72 });
    expect(crossing(15, 10, 'square', false).get('P025')?.page).toBe(1);
    expect(crossing(14, 40, 'square', false).get('P025')?.page).toBe(1);
  });
  it('compares the box widened by its bottom effect extent with the body bottom', () => {
    // The drawing ends at 712pt; effectExtent b=10pt takes it to 722pt.
    const placed = paginate({
      25: { anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 20, widthPt: 100, heightPt: 20, effectBPt: 10 }] },
    });
    expect(placed.get('P025')).toMatchObject({ page: 2, yPt: 72 });
  });
  it('raises a crossing float that stays below its anchor line to end at the body bottom', () => {
    // P025 is an automatic single line at 672pt without a grid; its float
    // would span 712-732pt and rises to 700-720pt.
    const at = (mode: number, heightPt: number) => layout(docx(27, {
      25: { auto: true, anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 40, widthPt: 100, heightPt }] },
    }, 1, mode));
    const risen = at(15, 20);
    expect(lines(risen).get('P025')).toMatchObject({ page: 1, yPt: 672 });
    expect(pictures(risen).get(25)).toEqual([{ page: 1, xPt: 72, yPt: 700 }]);
    expect(pictures(at(14, 20)).get(25)).toEqual([{ page: 1, xPt: 72, yPt: 712 }]);
    // Risen to 674pt the float would start above the line's bottom: the line moves.
    const moved = at(15, 46);
    expect(lines(moved).get('P025')).toMatchObject({ page: 2, yPt: 72 });
    expect(pictures(moved).get(25)?.map((picture) => picture.page)).toEqual([2]);
  });
  it('raises an effect-extent crossing so that the effect extent ends at the body bottom', () => {
    const result = layout(docx(27, {
      25: { auto: true, anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 30, widthPt: 100, heightPt: 10, effectBPt: 10 }] },
    }, 1, 15));
    expect(pictures(result).get(25)).toEqual([{ page: 1, xPt: 72, yPt: 700 }]);
  });
  it('compares the drawing box, not its bottom wrap distance, with the body bottom', () => {
    // The drawing ends at 712pt; distB=10pt takes its exclusion to 722pt.
    const placed = paginate({
      25: { anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 20, widthPt: 100, heightPt: 20, distBPt: 10 }] },
    });
    expect(placed.get('P025')?.page).toBe(1);
  });
  it('moves a keepNext predecessor with the deferred anchor paragraph', () => {
    const placed = paginate({
      24: { keepNext: true },
      25: { anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 40, widthPt: 100, heightPt: 30 }] },
    });
    expect(placed.get('P024')).toMatchObject({ page: 2, yPt: 72 });
    expect(placed.get('P025')).toMatchObject({ page: 2, yPt: 96 });
  });
  it('still defers a later crossing line of a paragraph that starts the page', () => {
    // Both floats end at 753pt. The first line starts the page and stays;
    // the second line's float moves it on.
    const placed = paginate({
      0: { lines: 3, anchors: [
        { wrap: 'square', relativeFrom: 'paragraph', yPt: 1, widthPt: 100, heightPt: 680, line: 0 },
        { wrap: 'square', relativeFrom: 'paragraph', yPt: 1, widthPt: 100, heightPt: 680, line: 1 },
      ] },
    }, { count: 2 });
    expect(placed.get('P000')).toMatchObject({ page: 1, xPt: 72 });
    expect(placed.get('P000.2')).toMatchObject({ page: 2, yPt: 72 });
    // The moved line keeps its float as its page's first line; the region ends after it.
    expect(placed.get('P000.3')).toMatchObject({ page: 3, yPt: 72 });
  });
  it('lays the lines left behind out without the deferred float', () => {
    // The float anchored on P025's second line lies beside its first line.
    const result = layout(docx(27, {
      25: { lines: 2, anchorAfterLine: 1,
        anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 0, widthPt: 100, heightPt: 60 }] },
    }, 1, 15, 24));
    const placed = lines(result);
    expect(placed.get('P025')).toMatchObject({ page: 1, xPt: 72 });
    expect(placed.get('P025.2')).toMatchObject({ page: 2, yPt: 72 });
    expect(pictures(result).get(25)?.map((picture) => picture.page)).toEqual([2]);
  });
  it('moves the anchor line to the next column of a multi-column section', () => {
    const at = (mode: number) => paginate({
      25: { anchors: [{ wrap: 'square', relativeFrom: 'paragraph', yPt: 40, widthPt: 100, heightPt: 30 }] },
    }, { columns: 2, count: 54, mode });
    const moved = at(15).get('P025');
    expect(moved).toMatchObject({ page: 1, yPt: 72 });
    expect(moved!.xPt).toBeGreaterThan(300);
    expect(at(14).get('P025')).toMatchObject({ page: 1, yPt: 672, xPt: 72 });
  });
  it('ends the region after a page-top line that keeps its crossing float', () => {
    // The 680pt float on P000's first line ends at 753pt; the line starts the
    // page, stays without wrapping around it, and the next line moves on.
    const placed = paginate({
      0: { lines: 2, anchors: [
        { wrap: 'square', relativeFrom: 'paragraph', yPt: 1, widthPt: 100, heightPt: 680, line: 0 },
      ] },
    }, { count: 2 });
    expect(placed.get('P000')).toMatchObject({ page: 1, xPt: 72 });
    expect(placed.get('P000.2')).toMatchObject({ page: 2, yPt: 72 });
    expect(placed.get('P001')?.page).toBe(2);
  });
});

describe('hostFloatBottomMarginPlan', () => {
  const float = (key: string, topPt: number, bottomPt: number, lineBottomPt: number) =>
    Object.freeze({ key, topPt, bottomPt, lineBottomPt });
  it('lifts a crossing float that stays below its anchor line and moves the first line that cannot', () => {
    const plan = hostFloatBottomMarginPlan([
      [float('a', 700, 730, 680)],
      [],
      [float('b', 712, 760, 690), float('c', 712, 724, 690)],
      [float('d', 712, 760, 700)],
    ], 720);
    expect(plan.lifts).toEqual([['a', 10]]);
    expect(plan.line).toEqual({ index: 2, keys: ['b', 'c'], fixedKeys: ['b'], lifts: [['c', 4]] });
  });
  it('lets a risen float start exactly at its anchor line bottom', () => {
    expect(hostFloatBottomMarginPlan([[float('a', 700, 730, 690)]], 720).line).toBeUndefined();
    expect(hostFloatBottomMarginPlan([[float('a', 700, 730, 690.5)]], 720).line?.index).toBe(0);
  });
  it('tests the paragraph at an offset and over a line prefix', () => {
    const byLine = [[], [float('a', 600, 620, 570)]];
    expect(hostFloatBottomMarginPlan(byLine, 720, 120).lifts).toEqual([['a', 20]]);
    expect(hostFloatBottomMarginPlan(byLine, 720, 140).line?.index).toBe(1);
    expect(hostFloatBottomMarginPlan(byLine, 720, 140, 1)).toEqual({ lifts: [] });
  });
});
