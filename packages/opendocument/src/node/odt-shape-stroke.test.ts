// ODF stroke caps (§20.171 svg:stroke-linecap) and joins (§20.170
// draw:stroke-linejoin) reach the DOCX model in the canvas vocabulary the
// shared shape renderer accepts. LibreOffice writes `svg:stroke-linecap` on
// ordinary shapes and lines, so a DrawingML token (`flat`) failed every page that
// carried one with "Invalid retained stroke geometry".
import { describe, expect, it } from 'vitest';
import { buildOdtFixture } from '../odt-fixture.js';
import { testOdtSource } from '../test-sources.js';
import { materializeDocxDocument } from './node-facade.js';
import { installImageBitmapShim, installOffscreenCanvasShim, type NodeCanvasFactory } from '../../../node/src/render.ts';
import { loadDocxRendererForTests, loadSkiaForTests } from '../../../node/src/test-imports.ts';
import type { DocxDocumentModel, ShapeRun } from '../../../docx/src/types.ts';

const skia = await loadSkiaForTests();
const renderer = await loadDocxRendererForTests();

const DRAW = 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"';
const graphic = (name: string, cap: string, join: string) =>
  `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${DRAW}`
  + ` draw:stroke="solid" svg:stroke-width="0cm" svg:stroke-color="#000000" draw:stroke-linejoin="${join}" svg:stroke-linecap="${cap}"`
  + ' draw:fill="solid" draw:fill-color="#ffffff" style:wrap="run-through" style:run-through="background"/></style:style>';

const fixture = () => buildOdtFixture({
  automaticStyles: graphic('gr1', 'butt', 'round') + graphic('gr2', 'round', 'bevel') + graphic('gr3', 'square', 'miter'),
  body: `<text:p><draw:line ${DRAW} draw:style-name="gr1" text:anchor-type="paragraph" svg:x1="1cm" svg:y1="0cm" svg:x2="1cm" svg:y2="5cm"/>`
    + `<draw:rect ${DRAW} draw:style-name="gr2" text:anchor-type="paragraph" svg:x="6cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"/>`
    + `<draw:rect ${DRAW} draw:style-name="gr3" text:anchor-type="paragraph" svg:x="9cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"/>`
    + 'body</text:p>',
});

const shapes = (model: DocxDocumentModel): ShapeRun[] => {
  const found: ShapeRun[] = [];
  JSON.stringify(model, (_key, value: unknown) => {
    if (value && typeof value === 'object' && (value as { type?: unknown }).type === 'shape') found.push(value as ShapeRun);
    return value;
  });
  return found;
};

describe('ODT shape strokes', () => {
  it('carry canvas cap and join names', async () => {
    const model = await materializeDocxDocument(fixture(), { odt: testOdtSource() }) as unknown as DocxDocumentModel;
    const strokes = shapes(model).map((shape) => [shape.strokeCap, shape.strokeJoin]);
    expect(strokes).toEqual([['butt', 'round'], ['round', 'bevel'], ['square', 'miter']]);
  });

  it.skipIf(!skia || !renderer)('render the page instead of rejecting its stroke geometry', async () => {
    const { Canvas } = skia!;
    const factory: NodeCanvasFactory = {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (() => { throw new Error('loadImage not needed'); }) as unknown as NodeCanvasFactory['loadImage'],
    };
    const model = await materializeDocxDocument(fixture(), { odt: testOdtSource() }) as unknown as DocxDocumentModel;
    const canvas = new Canvas(600, 850);
    const restoreOffscreen = installOffscreenCanvasShim(factory);
    const restoreImage = installImageBitmapShim(factory);
    try {
      await expect(renderer!.renderDocumentToCanvas(model, canvas as unknown as HTMLCanvasElement, 0, { dpr: 1, width: 600, currentDate: 0 }))
        .resolves.toBeUndefined();
    } finally {
      restoreImage();
      restoreOffscreen();
    }
  });
});
