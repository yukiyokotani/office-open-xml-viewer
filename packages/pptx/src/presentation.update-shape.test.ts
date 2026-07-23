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

describe('PptxPresentation.updateShape', () => {
  it('atomically updates arbitrary top-level and nested shape properties', () => {
    const { pres, model } = makePresentation();
    const original = model.slides[1].elements[0] as ShapeElement;

    const result = pres.updateShape(1, '7', (draft) => {
      draft.x = 914400;
      draft.rotation = 15;
      draft.fill = { fillType: 'solid', color: '4472C4' };
      draft.textBody!.verticalAnchor = 'ctr';
      const run = firstTextRun(draft);
      run.text = 'Updated';
      run.fontFamily = 'Aptos';
      run.fontSize = 24;
      run.color = 'FF0000';
      run.bold = true;
    });

    const updated = model.slides[1].elements[0] as ShapeElement;
    expect(result).toMatchObject({ slideIndex: 1, shapeId: '7' });
    expect(result.shape).not.toBe(updated);
    expect(updated).not.toBe(original);
    expect(updated).toMatchObject({
      id: '7',
      x: 914400,
      rotation: 15,
      fill: { fillType: 'solid', color: '4472C4' },
      textBody: { verticalAnchor: 'ctr' },
    });
    expect(firstTextRun(updated)).toMatchObject({
      text: 'Updated',
      fontFamily: 'Aptos',
      fontSize: 24,
      color: 'FF0000',
      bold: true,
    });
    expect(firstTextRun(original).text).toBe('Slide two');
    result.shape.x = 456;
    expect((model.slides[1].elements[0] as ShapeElement).x).toBe(914400);
  });

  it('scopes an id to the requested slide', () => {
    const { pres, model } = makePresentation();

    pres.updateShape(1, '7', (draft) => {
      firstTextRun(draft).text = 'Only slide two';
    });

    expect(firstTextRun(model.slides[0].elements[0] as ShapeElement).text).toBe('Slide one');
    expect(firstTextRun(model.slides[1].elements[0] as ShapeElement).text).toBe('Only slide two');
  });

  it('rejects identity changes without committing a partial update', () => {
    const { pres, model } = makePresentation();
    const before = model.slides[0].elements[0];

    expect(() =>
      pres.updateShape(0, '7', (draft) => {
        draft.x = 123;
        draft.id = '8';
      }),
    ).toThrow(/identity/i);

    expect(model.slides[0].elements[0]).toBe(before);
    expect((model.slides[0].elements[0] as ShapeElement).x).toBe(0);
  });

  it('rejects missing and parser-synthesized targets without a stable id', () => {
    const { pres, model } = makePresentation();
    (model.slides[0].elements[0] as ShapeElement).id = undefined;

    expect(() => pres.updateShape(0, '7', () => {})).toThrow(/shape.*7.*not found/i);
    expect(() => pres.updateShape(3, '7', () => {})).toThrow(/slide.*3.*out of range/i);
  });

  it('requires main mode because the editable model lives in the render worker otherwise', () => {
    const { pres } = makePresentation('worker');
    expect(() => pres.updateShape(0, '7', () => {})).toThrow(/mode.*main/i);
  });
});

describe('PptxPresentation.applyShapeChanges', () => {
  it('applies a serializable batch and returns a directly applicable inverse batch', () => {
    const { pres, model } = makePresentation();
    const original = structuredClone(model.slides[0].elements[0]) as ShapeElement;
    const insertedRun = textRun('Second run');
    const changes = [
      { op: 'replace', path: ['x'], value: 914400 },
      {
        op: 'replace',
        path: ['textBody', 'paragraphs', 0, 'runs', 0, 'text'],
        value: 'Updated',
      },
      {
        op: 'add',
        path: ['textBody', 'paragraphs', 0, 'runs', 1],
        value: insertedRun,
      },
      { op: 'add', path: ['hyperlink'], value: 'https://example.com' },
      { op: 'remove', path: ['name'] },
    ] satisfies PptxShapeChange[];

    const result = pres.applyShapeChanges(0, '7', changes);
    const updated = model.slides[0].elements[0] as ShapeElement;

    expect(updated.x).toBe(914400);
    expect(updated.name).toBeUndefined();
    expect(updated.hyperlink).toBe('https://example.com');
    expect(updated.textBody?.paragraphs[0]?.runs).toHaveLength(2);
    expect(firstTextRun(updated).text).toBe('Updated');
    expect(result).toMatchObject({ slideIndex: 0, shapeId: '7', applied: changes });
    expect(result.inverse.map((change) => change.op)).toEqual([
      'add',
      'remove',
      'remove',
      'replace',
      'replace',
    ]);

    pres.applyShapeChanges(0, '7', result.inverse);
    expect(model.slides[0].elements[0]).toEqual(original);
  });

  it('uses JSON Patch add semantics for arrays and replacement semantics for object properties', () => {
    const { pres, model } = makePresentation();
    const replacement = textRun('Replacement');

    const result = pres.applyShapeChanges(0, '7', [
      {
        op: 'add',
        path: ['textBody', 'paragraphs', 0, 'runs', 0],
        value: replacement,
      },
      { op: 'add', path: ['name'], value: 'Renamed' },
    ]);

    const updated = model.slides[0].elements[0] as ShapeElement;
    expect(updated.textBody?.paragraphs[0]?.runs).toHaveLength(2);
    expect(firstTextRun(updated).text).toBe('Replacement');
    expect(updated.name).toBe('Renamed');

    pres.applyShapeChanges(0, '7', result.inverse);
    expect(firstTextRun(model.slides[0].elements[0] as ShapeElement).text).toBe('Slide one');
    expect((model.slides[0].elements[0] as ShapeElement).name).toBe('Shape 7');
  });

  it('rejects an invalid change without committing earlier changes in the batch', () => {
    const { pres, model } = makePresentation();
    const before = model.slides[0].elements[0];

    expect(() =>
      pres.applyShapeChanges(0, '7', [
        { op: 'replace', path: ['x'], value: 123 },
        { op: 'replace', path: ['missing', 'value'], value: true },
      ]),
    ).toThrow(/index 1/i);

    expect(model.slides[0].elements[0]).toBe(before);
    expect((model.slides[0].elements[0] as ShapeElement).x).toBe(0);
  });

  it('rejects identity, unsafe paths, and non-JSON values atomically', () => {
    const { pres, model } = makePresentation();
    const before = structuredClone(model.slides[0].elements[0]);
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;

    expect(() =>
      pres.applyShapeChanges(0, '7', [{ op: 'replace', path: ['id'], value: '8' }]),
    ).toThrow(/identity/i);
    expect(() =>
      pres.applyShapeChanges(0, '7', [
        { op: 'add', path: ['__proto__', 'polluted'], value: true },
      ]),
    ).toThrow(/unsafe/i);
    expect(() =>
      pres.applyShapeChanges(0, '7', [{ op: 'replace', path: ['fill'], value: cyclic }]),
    ).toThrow(/cycle/i);

    expect(model.slides[0].elements[0]).toEqual(before);
  });

  it('detaches input changes and returned snapshots from the committed model', () => {
    const { pres, model } = makePresentation();
    const path: Array<string | number> = ['fill'];
    const value = { fillType: 'solid', color: '4472C4' };
    const result = pres.applyShapeChanges(0, '7', [{ op: 'replace', path, value }]);

    path[0] = 'stroke';
    value.color = 'FFFFFF';
    const applied = result.applied[0]!;
    if (applied.op === 'remove') throw new Error('Expected a value-bearing change');
    applied.value = null;
    result.shape.fill = null;

    expect((model.slides[0].elements[0] as ShapeElement).fill).toEqual({
      fillType: 'solid',
      color: '4472C4',
    });
  });

  it('treats an empty batch as a no-op', () => {
    const { pres, model } = makePresentation();
    const before = model.slides[0].elements[0];

    const result = pres.applyShapeChanges(0, '7', []);

    expect(model.slides[0].elements[0]).toBe(before);
    expect(result.applied).toEqual([]);
    expect(result.inverse).toEqual([]);
    expect(result.shape).toEqual(before);
    expect(result.shape).not.toBe(before);
  });

  it('requires main mode because the editable model lives in the render worker otherwise', () => {
    const { pres } = makePresentation('worker');
    expect(() => pres.applyShapeChanges(0, '7', [])).toThrow(/mode.*main/i);
  });
});
