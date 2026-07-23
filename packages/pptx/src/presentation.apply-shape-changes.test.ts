import { describe, expect, it } from 'vitest';
import { PptxPresentation } from './presentation';
import type { PptxShapeChange } from './shape-changes';
import type {
  Paragraph,
  Presentation,
  ShapeElement,
  TextBody,
  TextRunData,
} from './types';

function textRun(text: string): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 18,
    color: null,
    fontFamily: null,
  };
}

function paragraph(text: string): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [textRun(text)],
  };
}

function textBody(text: string): TextBody {
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph(text)],
    defaultFontSize: 18,
    defaultBold: null,
    defaultItalic: null,
    lIns: 91440,
    rIns: 91440,
    tIns: 45720,
    bIns: 45720,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  };
}

function shape(id: string, text: string): ShapeElement {
  return {
    type: 'shape',
    id,
    name: `Shape ${id}`,
    x: 0,
    y: 0,
    width: 2_000_000,
    height: 1_000_000,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody: textBody(text),
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
  };
}

function presentation(): Presentation {
  return {
    slideWidth: 9_144_000,
    slideHeight: 6_858_000,
    slides: [
      { index: 0, slideNumber: 1, background: null, elements: [shape('7', 'Slide one')] },
      { index: 1, slideNumber: 2, background: null, elements: [shape('7', 'Slide two')] },
    ],
    defaultTextColor: '383838',
    majorFont: null,
    minorFont: null,
  };
}

function makePresentation(mode: 'main' | 'worker' = 'main') {
  const model = presentation();
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  instance._mode = mode;
  instance._presentation = mode === 'main' ? model : null;
  instance._meta = mode === 'worker'
    ? {
        slideCount: model.slides.length,
        slideWidth: model.slideWidth,
        slideHeight: model.slideHeight,
      }
    : null;
  return { pres: instance as unknown as PptxPresentation, model };
}

function firstTextRun(shapeElement: ShapeElement): TextRunData {
  const run = shapeElement.textBody?.paragraphs[0]?.runs[0];
  if (!run || run.type !== 'text') throw new Error('Expected a text run');
  return run;
}

describe('PptxPresentation.applyShapeChanges', () => {
  it('applies typed semantic changes and returns a directly applicable inverse batch', () => {
    const { pres, model } = makePresentation();
    const original = structuredClone(model.slides[0].elements[0]) as ShapeElement;
    const changes = [
      {
        type: 'shape.update',
        patch: {
          x: 914400,
          fill: { fillType: 'solid', color: '4472C4' },
          hyperlink: 'https://example.com',
        },
        unset: ['name'],
      },
      {
        type: 'textBody.update',
        patch: { verticalAnchor: 'ctr' },
      },
      {
        type: 'paragraph.update',
        paragraphIndex: 0,
        patch: { alignment: 'ctr' },
      },
      {
        type: 'textRun.update',
        paragraphIndex: 0,
        runIndex: 0,
        patch: {
          text: 'Updated',
          fontFamily: 'Aptos',
          fontSize: 24,
          color: 'FF0000',
          bold: true,
        },
      },
      {
        type: 'run.insert',
        paragraphIndex: 0,
        runIndex: 1,
        run: textRun('Second run'),
      },
    ] satisfies PptxShapeChange[];

    const result = pres.applyShapeChanges({ slideIndex: 0, shapeId: '7', changes });
    const updated = model.slides[0].elements[0] as ShapeElement;

    expect(updated).toMatchObject({
      x: 914400,
      fill: { fillType: 'solid', color: '4472C4' },
      hyperlink: 'https://example.com',
      textBody: {
        verticalAnchor: 'ctr',
        paragraphs: [{ alignment: 'ctr' }],
      },
    });
    expect(updated.name).toBeUndefined();
    expect(updated.textBody?.paragraphs[0]?.runs).toHaveLength(2);
    expect(firstTextRun(updated)).toMatchObject({
      text: 'Updated',
      fontFamily: 'Aptos',
      fontSize: 24,
      color: 'FF0000',
      bold: true,
    });
    expect(result.applied).toEqual(changes);
    expect(result.inverse.map((change) => change.type)).toEqual([
      'run.remove',
      'textRun.update',
      'paragraph.update',
      'textBody.update',
      'shape.update',
    ]);

    pres.applyShapeChanges({
      slideIndex: result.slideIndex,
      shapeId: result.shapeId,
      changes: result.inverse,
    });
    expect(model.slides[0].elements[0]).toEqual(original);
  });

  it('supports paragraph and run insertion, replacement, and removal', () => {
    const { pres, model } = makePresentation();
    const original = structuredClone(model.slides[0].elements[0]);

    const result = pres.applyShapeChanges({
      slideIndex: 0,
      shapeId: '7',
      changes: [
        { type: 'paragraph.insert', paragraphIndex: 1, paragraph: paragraph('New paragraph') },
        { type: 'run.replace', paragraphIndex: 0, runIndex: 0, run: textRun('Replacement') },
        { type: 'run.remove', paragraphIndex: 1, runIndex: 0 },
      ],
    });

    const updated = model.slides[0].elements[0] as ShapeElement;
    expect(firstTextRun(updated).text).toBe('Replacement');
    expect(updated.textBody?.paragraphs).toHaveLength(2);
    expect(updated.textBody?.paragraphs[1]?.runs).toEqual([]);

    pres.applyShapeChanges({
      slideIndex: 0,
      shapeId: '7',
      changes: result.inverse,
    });
    expect(model.slides[0].elements[0]).toEqual(original);
  });

  it('supports replacing and restoring a complete text body', () => {
    const { pres, model } = makePresentation();

    const result = pres.applyShapeChanges({
      slideIndex: 0,
      shapeId: '7',
      changes: [{ type: 'textBody.replace', textBody: null }],
    });
    expect((model.slides[0].elements[0] as ShapeElement).textBody).toBeNull();

    pres.applyShapeChanges({
      slideIndex: 0,
      shapeId: '7',
      changes: result.inverse,
    });
    expect(firstTextRun(model.slides[0].elements[0] as ShapeElement).text).toBe('Slide one');
  });

  it('scopes a shape id to the requested slide', () => {
    const { pres, model } = makePresentation();

    pres.applyShapeChanges({
      slideIndex: 1,
      shapeId: '7',
      changes: [{
        type: 'textRun.update',
        paragraphIndex: 0,
        runIndex: 0,
        patch: { text: 'Only slide two' },
      }],
    });

    expect(firstTextRun(model.slides[0].elements[0] as ShapeElement).text).toBe('Slide one');
    expect(firstTextRun(model.slides[1].elements[0] as ShapeElement).text).toBe('Only slide two');
  });

  it('rejects missing targets and invalid indices without a partial commit', () => {
    const { pres, model } = makePresentation();
    const before = model.slides[0].elements[0];

    expect(() =>
      pres.applyShapeChanges({
        slideIndex: 0,
        shapeId: '7',
        changes: [
          { type: 'shape.update', patch: { x: 123 } },
          {
            type: 'textRun.update',
            paragraphIndex: 99,
            runIndex: 0,
            patch: { text: 'Missing' },
          },
        ],
      }),
    ).toThrow(/index 1/i);
    expect(model.slides[0].elements[0]).toBe(before);

    (model.slides[0].elements[0] as ShapeElement).id = undefined;
    expect(() =>
      pres.applyShapeChanges({ slideIndex: 0, shapeId: '7', changes: [] }),
    ).toThrow(/shape.*7.*not found/i);
    expect(() =>
      pres.applyShapeChanges({ slideIndex: 3, shapeId: '7', changes: [] }),
    ).toThrow(/slide.*3.*out of range/i);
  });

  it('rejects immutable properties and non-JSON values atomically', () => {
    const { pres, model } = makePresentation();
    const before = structuredClone(model.slides[0].elements[0]);
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;

    expect(() =>
      pres.applyShapeChanges({
        slideIndex: 0,
        shapeId: '7',
        changes: [{
          type: 'shape.update',
          patch: { id: '8' },
        } as unknown as PptxShapeChange],
      }),
    ).toThrow(/immutable/i);
    expect(() =>
      pres.applyShapeChanges({
        slideIndex: 0,
        shapeId: '7',
        changes: [{
          type: 'shape.update',
          patch: { fill: cyclic },
        } as unknown as PptxShapeChange],
      }),
    ).toThrow(/cycle/i);
    expect(model.slides[0].elements[0]).toEqual(before);
  });

  it('detaches request and result values from the committed model', () => {
    const { pres, model } = makePresentation();
    const fill = { fillType: 'solid' as const, color: '4472C4' };
    const request = {
      slideIndex: 0,
      shapeId: '7',
      changes: [{ type: 'shape.update', patch: { fill } }],
    } satisfies Parameters<PptxPresentation['applyShapeChanges']>[0];

    const result = pres.applyShapeChanges(request);
    fill.color = 'FFFFFF';
    const applied = result.applied[0]!;
    if (applied.type !== 'shape.update') throw new Error('Expected a shape update');
    applied.patch.fill = null;
    result.shape.fill = null;

    expect((model.slides[0].elements[0] as ShapeElement).fill).toEqual({
      fillType: 'solid',
      color: '4472C4',
    });
  });

  it('treats an empty batch as a no-op', () => {
    const { pres, model } = makePresentation();
    const before = model.slides[0].elements[0];

    const result = pres.applyShapeChanges({ slideIndex: 0, shapeId: '7', changes: [] });

    expect(model.slides[0].elements[0]).toBe(before);
    expect(result.applied).toEqual([]);
    expect(result.inverse).toEqual([]);
    expect(result.shape).toEqual(before);
    expect(result.shape).not.toBe(before);
  });

  it('requires main mode because the editable model lives in the render worker otherwise', () => {
    const { pres } = makePresentation('worker');
    expect(() =>
      pres.applyShapeChanges({ slideIndex: 0, shapeId: '7', changes: [] }),
    ).toThrow(/mode.*main/i);
  });
});
