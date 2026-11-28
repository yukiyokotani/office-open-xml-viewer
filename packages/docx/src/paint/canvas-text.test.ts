import { describe, expect, it } from 'vitest';
import { stableFingerprint } from '../layout/fingerprint.js';
import type { ParagraphLayout, TableLayout, TextBoxLayout } from '../layout/types.js';
import { paintParagraphLayout, paintPlacedParagraphLayout, paintPlacedTextBoxLayout } from './canvas-text.js';
import { inverseMapAffinePoint, inverseMapAffineVector, mapAffinePoint } from './affine.js';
import type { CanvasPaintResourcePainter } from './types.js';

const noPaintResources: CanvasPaintResourcePainter = {
  paint(resourceKey, kind): never {
    throw new Error(`Unexpected ${kind} paint resource: ${resourceKey}`);
  },
};

const fontRoute = {
  familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-font-route',
} as const;

function node(): ParagraphLayout {
  return {
    kind: 'paragraph', id: 'p', source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 10, widthPt: 50, heightPt: 14 },
    inkBounds: { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10 },
    advancePt: 14, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 4 }, bounds: { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10 },
      baselinePt: 18, advancePt: 14,
      placements: [{
        kind: 'text', text: 'test', range: { start: 0, end: 4 },
        origin: { xPt: 10, yPt: 18 }, bounds: { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10 },
        advancePt: 20, clusters: [{
          range: { start: 0, end: 4 }, offset: { xPt: 0, yPt: 0 }, advancePt: 20,
        }], paintOps: [{
          text: 'test', range: { start: 0, end: 4 }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb',
        }], color: { kind: 'explicit', color: '#112233' }, fontRoute, fontSizePt: 10,
        fontWeight: 400, fontStyle: 'normal', direction: 'ltr', decorations: [],
      }],
    }],
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

describe('paintParagraphLayout', () => {
  it.each([
    { dpr: 1, expectedWidthPt: 1 },
    { dpr: 2, expectedWidthPt: 0.5 },
  ])(
    'rasterizes a retained paragraph-border hairline to one device pixel at DPR $dpr',
    ({ dpr, expectedWidthPt }) => {
      const strokes: Array<Readonly<{
        widthPt: number;
        from: readonly [number, number];
        to: readonly [number, number];
      }>> = [];
      let path: Array<readonly [number, number]> = [];
      const ctx = {
        globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
        font: '', textAlign: 'left' as CanvasTextAlign,
        textBaseline: 'alphabetic' as CanvasTextBaseline,
        direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
        fontKerning: 'auto' as CanvasFontKerning,
        save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, transform() {},
        setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() { path = []; },
        rect() {}, clip() {},
        moveTo(x: number, y: number) { path.push([x, y]); },
        lineTo(x: number, y: number) { path.push([x, y]); },
        stroke() {
          if (path.length === 2) {
            strokes.push({ widthPt: this.lineWidth, from: path[0]!, to: path[1]! });
          }
        },
        fill() {}, drawImage() {}, fillText() {},
      };
      const source = node();
      const bordered = {
        ...source,
        borders: [{
          edge: 'top' as const,
          from: { xPt: 10, yPt: 10.2 },
          to: { xPt: 30, yPt: 10.2 },
          color: '#111111',
          widthPt: 0.5,
          authoredStyle: 'single',
          style: 'solid' as const,
        }],
      } as ParagraphLayout;
      const before = stableFingerprint('paragraph-border-hairline', bordered);

      paintParagraphLayout(bordered, {
        ctx,
        scale: 1,
        dpr,
        resources: noPaintResources,
      });

      expect(strokes).toHaveLength(1);
      expect(strokes[0]!.widthPt).toBe(expectedWidthPt);
      expect(strokes[0]!.from[0]).toBe(10);
      expect(strokes[0]!.to[0]).toBe(30);
      expect(strokes[0]!.from[1] * dpr - 0.5).toBeCloseTo(
        Math.round(strokes[0]!.from[1] * dpr - 0.5),
        10,
      );
      expect(strokes[0]!.to[1]).toBe(strokes[0]!.from[1]);
      expect(bordered.borders[0]!.widthPt).toBe(0.5);
      expect(stableFingerprint('paragraph-border-hairline', bordered)).toBe(before);
    },
  );

  it('does not widen retained text decorations with the paragraph-border hairline policy', () => {
    const widths: number[] = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, transform() {},
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {}, rect() {}, clip() {},
      moveTo() {}, lineTo() {}, stroke() { widths.push(this.lineWidth); },
      fill() {}, drawImage() {}, fillText() {},
    };
    const source = node();
    const line = source.lines[0]!;
    const placement = line.placements[0]!;
    if (placement.kind !== 'text') throw new Error('Expected text placement');
    const bordered = {
      ...source,
      lines: [{
        ...line,
        placements: [{
          ...placement,
          decorations: [{
            kind: 'underline' as const,
            from: { xPt: 10, yPt: 19 },
            to: { xPt: 30, yPt: 19 },
            color: '#222222',
            widthPt: 0.5,
            style: 'solid' as const,
          }],
        }],
      }],
      borders: [{
        edge: 'bottom' as const,
        from: { xPt: 10, yPt: 20 },
        to: { xPt: 30, yPt: 20 },
        color: '#111111',
        widthPt: 0.5,
        authoredStyle: 'single',
        style: 'solid' as const,
      }],
    } as ParagraphLayout;

    paintParagraphLayout(bordered, {
      ctx,
      scale: 1,
      dpr: 1,
      resources: noPaintResources,
    });

    expect(widths).toEqual([0.5, 1]);
  });

  it('resets an unthemed nested text box to the document default text color', () => {
    const fills: Array<readonly [string, string]> = [];
    let fillStyle = '';
    const ctx = {
      globalAlpha: 1,
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { fillStyle = String(value); },
      strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, transform() {},
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {}, rect() {}, clip() {},
      moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText(text: string) { fills.push([text, fillStyle]); },
    } as unknown as CanvasRenderingContext2D;
    const defaultColorParagraph = (text: string, textBoxes: readonly TextBoxLayout[] = []): ParagraphLayout => {
      const paragraph = node();
      const line = paragraph.lines[0]!;
      const placement = line.placements[0]!;
      if (placement.kind !== 'text') throw new Error('Expected text placement');
      return {
        ...paragraph,
        id: `p-${text}`,
        lines: [{
          ...line,
          placements: [{
            ...placement,
            text,
            paintOps: [{ ...placement.paintOps[0]!, text }],
            color: { kind: 'default' },
          }],
        }],
        textBoxes,
      };
    };
    const textBox = (
      id: string,
      paragraph: ParagraphLayout,
      defaultTextColor?: string,
    ): TextBoxLayout => ({
      kind: 'textbox',
      id,
      source: { story: 'textbox', storyInstance: id, path: [] },
      flowDomainId: `textbox:${id}`,
      ordinaryFlow: false,
      flowBounds: paragraph.flowBounds,
      inkBounds: paragraph.inkBounds,
      advancePt: 0,
      story: {
        story: 'textbox',
        flowBounds: paragraph.flowBounds,
        inkBounds: paragraph.inkBounds,
        blocks: [paragraph],
        advancePt: paragraph.advancePt,
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      ...(defaultTextColor ? { defaultTextColor } : {}),
    });
    const inner = textBox('inner', defaultColorParagraph('INNR'));
    const outer = textBox('outer', defaultColorParagraph('OUTR', [inner]), '#aa0000');

    paintPlacedTextBoxLayout(outer, {
      ctx,
      scale: 1,
      dpr: 1,
      resources: noPaintResources,
      defaultTextColor: '#123456',
    });

    expect(fills).toEqual([
      ['OUTR', '#aa0000'],
      ['INNR', '#123456'],
    ]);
  });

  it('paints retained table blocks inside a text-box story', () => {
    const texts: string[] = [];
    const ctx = {
      globalAlpha: 1,
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, transform() {},
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {}, rect() {}, clip() {},
      moveTo() {}, lineTo() {}, stroke() {}, fill() {},
      fillText(text: string) { texts.push(text); },
      drawImage() {},
    } as unknown as CanvasRenderingContext2D;
    const paragraph = {
      ...node(),
      source: { story: 'textbox' as const, storyInstance: 'shape:table', path: [0, 0, 0, 0] },
      flowDomainId: 'textbox:table',
    };
    const bounds = { xPt: 0, yPt: 0, widthPt: 50, heightPt: 20 };
    const table: TableLayout = {
      kind: 'table',
      id: 'textbox-table',
      source: { story: 'textbox', storyInstance: 'shape:table', path: [0] },
      flowDomainId: 'textbox:table',
      ordinaryFlow: true,
      flowBounds: bounds,
      inkBounds: bounds,
      advancePt: 20,
      columnWidthsPt: [50],
      borders: [],
      rows: [{
        kind: 'table-row',
        id: 'textbox-table-row',
        source: { story: 'textbox', storyInstance: 'shape:table', path: [0, 0] },
        flowDomainId: 'textbox:table',
        ordinaryFlow: true,
        flowBounds: bounds,
        inkBounds: bounds,
        advancePt: 20,
        heightPt: 20,
        contentHeightPt: 14,
        cells: [{
          kind: 'table-cell',
          id: 'textbox-table-cell',
          source: { story: 'textbox', storyInstance: 'shape:table', path: [0, 0, 0] },
          flowDomainId: 'textbox:table',
          ordinaryFlow: true,
          flowBounds: bounds,
          inkBounds: bounds,
          advancePt: 20,
          contentBounds: bounds,
          verticalMerge: 'none',
          vAlign: 'top',
          blocks: [{ layout: paragraph, offsetPt: 0, advancePt: paragraph.advancePt }],
        }],
      }],
    };
    const textBox: TextBoxLayout = {
      kind: 'textbox',
      id: 'textbox-with-table',
      source: { story: 'textbox', storyInstance: 'shape:table', path: [] },
      flowDomainId: 'textbox:table',
      ordinaryFlow: false,
      flowBounds: bounds,
      inkBounds: bounds,
      advancePt: 0,
      story: {
        story: 'textbox',
        flowBounds: bounds,
        inkBounds: bounds,
        blocks: [table],
        advancePt: 20,
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    };

    paintPlacedTextBoxLayout(textBox, {
      ctx, scale: 1, dpr: 1, resources: noPaintResources,
    });

    expect(texts).toContain('test');
  });

  it('clips direct and placed paragraph paint to the retained clipBounds', () => {
    const calls: unknown[] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() { calls.push('save'); }, restore() { calls.push('restore'); },
      translate() {}, scale() {}, rotate() {}, setLineDash() {}, fillRect() {}, strokeRect() {},
      beginPath() { calls.push('beginPath'); },
      rect(x: number, y: number, w: number, h: number) { calls.push(['rect', x, y, w, h]); },
      clip() { calls.push('clip'); },
      moveTo() {}, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const clipped = {
      ...node(),
      clipBounds: { xPt: 2, yPt: 3, widthPt: 40, heightPt: 12 },
    } as ParagraphLayout;

    paintPlacedParagraphLayout(clipped, { xPt: 20, yPt: 30 }, {
      ctx, scale: 2, dpr: 1, resources: noPaintResources,
    });

    expect(calls).toContainEqual(['rect', 2, 3, 40, 12]);
    expect(calls).toContain('clip');

    calls.length = 0;
    paintParagraphLayout(clipped, { ctx, scale: 1, dpr: 1, resources: noPaintResources });
    expect(calls).toContainEqual(['rect', 2, 3, 40, 12]);
    expect(calls).toContain('clip');
  });

  it('owns placement scale', () => {
    const transforms: unknown[] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() { transforms.push('save'); }, restore() { transforms.push('restore'); },
      translate(x: number, y: number) { transforms.push(['translate', x, y]); },
      scale(x: number, y: number) { transforms.push(['scale', x, y]); },
      rotate() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;

    paintPlacedParagraphLayout(node(), { xPt: 30, yPt: 40 }, {
      ctx, scale: 2, dpr: 1, resources: noPaintResources,
    });

    expect(transforms).toEqual([
      'save', ['translate', 40, 60], ['scale', 2, 2], 'restore',
    ]);
  });

  it.each([
    { scale: 1, dpr: 1, placement: { xPt: 10.25, yPt: 10.25 }, outer: { xPt: 0, yPt: 0 } },
    { scale: 1.5, dpr: 2, placement: { xPt: 10.2, yPt: 10.3 }, outer: { xPt: .15, yPt: .1 } },
  ])('snaps borders using the final translated coordinate at scale=$scale dpr=$dpr', ({ scale, dpr, placement, outer }) => {
    const translations: Array<readonly [number, number]> = [];
    const moves: Array<readonly [number, number]> = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {},
      translate(x: number, y: number) { translations.push([x, y]); },
      scale() {}, rotate() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo(x: number, y: number) { moves.push([x, y]); }, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const bordered = {
      ...node(),
      borders: [
        { edge: 'top', from: { xPt: 10, yPt: 10 }, to: { xPt: 30, yPt: 10 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
        { edge: 'left', from: { xPt: 10, yPt: 10 }, to: { xPt: 10, yPt: 20 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
      ],
    } as ParagraphLayout;

    paintPlacedParagraphLayout(bordered, placement, {
      ctx, scale, dpr, resources: noPaintResources,
      pointToCss: { a: scale, b: 0, c: 0, d: scale, e: outer.xPt * scale, f: outer.yPt * scale },
    });

    const [translated] = translations;
    const horizontalFinalCss = outer.yPt * scale + translated![1] + moves[0]![1] * scale;
    const verticalFinalCss = outer.xPt * scale + translated![0] + moves[1]![0] * scale;
    const target = Math.round(scale * dpr) % 2 === 1 ? .5 : 0;
    expect(horizontalFinalCss * dpr - target).toBeCloseTo(
      Math.round(horizontalFinalCss * dpr - target),
      10,
    );
    expect(verticalFinalCss * dpr - target).toBeCloseTo(
      Math.round(verticalFinalCss * dpr - target),
      10,
    );
  });

  it.each([
    { scale: 1, dpr: 1 },
    { scale: 1.5, dpr: 2 },
  ])('snaps vertical text-box borders in final device axes at scale=$scale dpr=$dpr', ({ scale, dpr }) => {
    type Matrix = { a: number; b: number; c: number; d: number; e: number; f: number };
    let matrix: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    const stack: Matrix[] = [];
    const moves: Array<readonly [number, number]> = [];
    const transform = (x: number, y: number): readonly [number, number] => [
      matrix.a * x + matrix.c * y + matrix.e,
      matrix.b * x + matrix.d * y + matrix.f,
    ];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() { stack.push({ ...matrix }); },
      restore() { matrix = stack.pop()!; },
      translate(x: number, y: number) {
        matrix = {
          ...matrix,
          e: matrix.e + matrix.a * x + matrix.c * y,
          f: matrix.f + matrix.b * x + matrix.d * y,
        };
      },
      transform(a: number, b: number, c: number, d: number, e: number, f: number) {
        matrix = { a: matrix.a * a + matrix.c * b, b: matrix.b * a + matrix.d * b,
          c: matrix.a * c + matrix.c * d, d: matrix.b * c + matrix.d * d,
          e: matrix.a * e + matrix.c * f + matrix.e, f: matrix.b * e + matrix.d * f + matrix.f };
      },
      scale(x: number, y: number) {
        matrix = { ...matrix, a: matrix.a * x, b: matrix.b * x, c: matrix.c * y, d: matrix.d * y };
      },
      rotate(angle: number) {
        const cos = Math.cos(angle), sin = Math.sin(angle);
        matrix = {
          ...matrix,
          a: matrix.a * cos + matrix.c * sin,
          b: matrix.b * cos + matrix.d * sin,
          c: -matrix.a * sin + matrix.c * cos,
          d: -matrix.b * sin + matrix.d * cos,
        };
      },
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo(x: number, y: number) { moves.push(transform(x, y)); }, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const paragraph = {
      ...node(),
      lines: [],
      borders: [
        {
          edge: 'top', from: { xPt: 2, yPt: 3.2 }, to: { xPt: 22, yPt: 3.2 },
          color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid',
        },
        {
          edge: 'left', from: { xPt: 4.1, yPt: 2 }, to: { xPt: 4.1, yPt: 12 },
          color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid',
        },
      ],
    } as ParagraphLayout;
    const textBox = {
      kind: 'textbox', id: 'textbox', source: paragraph.source,
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 10.2, yPt: 20.3, widthPt: 40.2, heightPt: 20.4 },
      inkBounds: { xPt: 10.2, yPt: 20.3, widthPt: 40.2, heightPt: 20.4 },
      advancePt: 0,
      story: {
        story: 'textbox',
        flowBounds: paragraph.flowBounds,
        inkBounds: paragraph.inkBounds,
        blocks: [paragraph],
        advancePt: paragraph.advancePt,
        diagnostics: [],
      },
      transform: { a: 0, b: 1, c: -1, d: 0, e: 30.3, f: 30.5 },
      writingMode: 'vertical-rl', verticalMode: 'vert',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    } as TextBoxLayout;

    paintPlacedTextBoxLayout(textBox, { ctx, scale, dpr, resources: noPaintResources });

    const target = Math.round(scale * dpr) % 2 === 1 ? .5 : 0;
    // A local horizontal border becomes device-vertical; a local vertical border
    // becomes device-horizontal after the production text-box quarter-turn.
    expect(moves[0]![0] * dpr - target).toBeCloseTo(Math.round(moves[0]![0] * dpr - target), 10);
    expect(moves[1]![1] * dpr - target).toBeCloseTo(Math.round(moves[1]![1] * dpr - target), 10);
  });

  it('does not apply axis snapping when an affine rotates an authored axis off device axes', () => {
    const moves: Array<readonly [number, number]> = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, setLineDash() {},
      fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo(x: number, y: number) { moves.push([x, y]); }, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const bordered = {
      ...node(), lines: [], borders: [{
        edge: 'top', from: { xPt: 2.25, yPt: 3.2 }, to: { xPt: 22.25, yPt: 3.2 },
        color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid',
      }],
    } as ParagraphLayout;
    const angle = Math.PI / 4;

    paintParagraphLayout(bordered, {
      ctx, scale: 1, dpr: 2, resources: noPaintResources,
      pointToCss: {
        a: Math.cos(angle), b: Math.sin(angle),
        c: -Math.sin(angle), d: Math.cos(angle), e: .3, f: .4,
      },
    });

    expect(moves[0]).toEqual([2.25, 3.2]);
  });

  it('does not snap a 20pt authored axis after any nonzero device-axis rotation', () => {
    const moves: Array<readonly [number, number]> = [];
    const lines: Array<readonly [number, number]> = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {}, setLineDash() {},
      fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo(x: number, y: number) { moves.push([x, y]); },
      lineTo(x: number, y: number) { lines.push([x, y]); },
      stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const from = { xPt: 2.25, yPt: 3.2 };
    const to = { xPt: 22.25, yPt: 3.2 };
    const bordered = {
      ...node(), lines: [], borders: [{
        edge: 'top', from, to, color: '#111111', widthPt: 1,
        authoredStyle: 'single', style: 'solid',
      }],
    } as ParagraphLayout;
    const angle = 1e-11;

    paintParagraphLayout(bordered, {
      ctx, scale: 1, dpr: 2, resources: noPaintResources,
      pointToCss: {
        a: Math.cos(angle), b: Math.sin(angle),
        c: -Math.sin(angle), d: Math.cos(angle), e: 0, f: 0,
      },
    });

    expect(moves[0]).toEqual([from.xPt, from.yPt]);
    expect(lines[0]).toEqual([to.xPt, to.yPt]);
  });

  it('round-trips finite invertible affine data with a determinant below 1e-12', () => {
    const matrix = { a: 1e-7, b: 0, c: 0, d: 1e-6, e: .3, f: .4 };
    const point = { xPt: 2.25, yPt: 3.2 };
    const vector = { xPt: -.75, yPt: 1.1 };

    const mapped = mapAffinePoint(matrix, point);
    const restoredPoint = inverseMapAffinePoint(matrix, mapped);
    const restoredVector = inverseMapAffineVector(matrix, {
      xPt: matrix.a * vector.xPt + matrix.c * vector.yPt,
      yPt: matrix.b * vector.xPt + matrix.d * vector.yPt,
    });

    expect(matrix.a * matrix.d - matrix.b * matrix.c).toBeGreaterThan(0);
    expect(matrix.a * matrix.d - matrix.b * matrix.c).toBeLessThan(1e-12);
    expect(restoredPoint?.xPt).toBeCloseTo(point.xPt, 8);
    expect(restoredPoint?.yPt).toBeCloseTo(point.yPt, 8);
    expect(restoredVector?.xPt).toBeCloseTo(vector.xPt, 8);
    expect(restoredVector?.yPt).toBeCloseTo(vector.yPt, 8);
  });

  it('paints retained DOCX dash patterns and double rails without source parsing', () => {
    const operations: unknown[] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, rotate() {},
      setLineDash(pattern: number[]) { operations.push(['dash', pattern]); },
      fillRect(...args: number[]) { operations.push(['fillRect', ...args]); },
      strokeRect() {}, beginPath() {}, moveTo() {}, lineTo() {},
      stroke() { operations.push(['stroke']); }, fillText() {},
    } as unknown as CanvasRenderingContext2D;
    const retained = {
      ...node(),
      borders: [{
        edge: 'top', from: { xPt: 10, yPt: 10 }, to: { xPt: 30, yPt: 10 },
        color: '#111111', widthPt: 2, authoredStyle: 'dotDash', style: 'dashed', dashPatternPt: [2, 4, 6, 4],
      }],
      lines: [{
        ...node().lines[0]!, placements: [{
          ...(node().lines[0]!.placements[0] as object),
          runBorderFragments: [{
            edge: 'bottom', from: { xPt: 10, yPt: 20 }, to: { xPt: 30, yPt: 20 },
            color: '#222222', widthPt: 3, authoredStyle: 'double', style: 'double', dashPatternPt: [],
          }],
        }],
      }],
    } as unknown as ParagraphLayout;

    paintParagraphLayout(retained, { ctx, scale: 1, dpr: 2, resources: noPaintResources });

    expect(operations).toContainEqual(['dash', [2, 4, 6, 4]]);
    expect(operations.filter((operation) => Array.isArray(operation) && operation[0] === 'fillRect')).toHaveLength(2);
  });

  it.each([1, 2])('uses stored glyph geometry at scale %s without measuring or mutating layout', (scale) => {
    const calls: Array<readonly [string, number, number]> = [];
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
      save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, setLineDash() {},
      fillRect() {}, strokeRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
      fillText(text: string, x: number, y: number) { calls.push([text, x, y]); },
      measureText() { throw new Error('paint must not measure text'); },
    } as unknown as CanvasRenderingContext2D;
    const layout = node();
    const before = stableFingerprint('paragraph', layout);

    paintParagraphLayout(layout, { ctx, scale, dpr: 1, resources: noPaintResources });

    expect(calls).toEqual([['test', 10, 18]]);
    expect(stableFingerprint('paragraph', layout)).toBe(before);
  });

  it('paints only retained leader, ruby, emphasis, decoration, border, and highlight operations', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const operations: unknown[] = [];
    const retained: ParagraphLayout = {
      ...layout,
      lines: [{
        ...layout.lines[0]!,
        placements: [
          {
            kind: 'tab', range: { start: 0, end: 1 },
            bounds: { xPt: 2, yPt: 4, widthPt: 12, heightPt: 10 }, advancePt: 12,
            leader: 'dot',
            leaderGlyphs: [{
              text: '.', origin: { xPt: 4, yPt: 12 }, fontRoute,
              fontSizePt: 8, fontWeight: 400, fontStyle: 'normal', color: { kind: 'explicit', color: '#101010' },
            }],
          },
          {
            ...placement,
            highlightFragments: [{ rect: { xPt: 14, yPt: 10, widthPt: 24, heightPt: 10 }, color: '#ffff00' }],
            ruby: {
              text: 'ふり', advancePt: 8,
              authored: { align: 'center', raisePt: 5, language: 'ja-JP' },
              paintOps: [{
                text: 'ふり', origin: { xPt: 20, yPt: 13 }, fontRoute,
                fontSizePt: 5, fontWeight: 400, fontStyle: 'normal', color: { kind: 'explicit', color: '#112233' },
              }],
            },
            emphasis: {
              authored: 'circle',
              glyphs: [{
                text: '○', origin: { xPt: 20, yPt: 7 }, fontRoute,
                fontSizePt: 5, fontWeight: 400, fontStyle: 'normal',
                color: { kind: 'explicit', color: '#112233' },
                inkBounds: { xMinPt: 0, xMaxPt: 5, ascentPt: 4, descentPt: 1 },
              }],
            },
            decorations: [{
              kind: 'underline', authoredStyle: 'double',
              from: { xPt: 14, yPt: 20 }, to: { xPt: 38, yPt: 20 },
              color: '#445566', widthPt: .75, style: 'double',
              path: [{ xPt: 14, yPt: 20 }, { xPt: 38, yPt: 20 }],
            }],
            runBorderFragments: [{
              edge: 'top', from: { xPt: 12, yPt: 8 }, to: { xPt: 40, yPt: 8 },
              color: '#778899', widthPt: 1, authoredStyle: 'single', style: 'solid',
            }],
          },
        ],
      }],
    };
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, setLineDash(value: number[]) { operations.push(['dash', value]); },
      fillRect(this: { fillStyle: string }, ...args: number[]) { operations.push(['fillRect', this.fillStyle, ...args]); },
      beginPath() { operations.push(['begin']); }, moveTo(...args: number[]) { operations.push(['moveTo', ...args]); },
      lineTo(...args: number[]) { operations.push(['lineTo', ...args]); },
      stroke(this: { strokeStyle: string }) { operations.push(['stroke', this.strokeStyle]); },
      fillText(this: { fillStyle: string }, text: string, x: number, y: number) { operations.push(['text', text, x, y, this.fillStyle]); },
      measureText() { throw new Error('paint must not measure text'); },
    } as unknown as CanvasRenderingContext2D;

    paintParagraphLayout(retained, { ctx, scale: 1, dpr: 1, resources: noPaintResources });

    expect(operations).toContainEqual(['text', '.', 4, 12, '#101010']);
    expect(operations).toContainEqual(['fillRect', '#ffff00', 14, 10, 24, 10]);
    expect(operations).toContainEqual(['text', 'ふり', 20, 13, '#112233']);
    expect(operations).toContainEqual(['text', '○', 20, 7, '#112233']);
    expect(operations).toContainEqual(['stroke', '#778899']);
  });

  it('dispatches stable resource keys through the explicit atomic painter registry', () => {
    const layout = node();
    const resource = {
      kind: 'resource' as const, resourceKey: 'chart:body:0', resourceKind: 'chart' as const,
      range: { start: 0, end: 1 },
      bounds: { xPt: 12, yPt: 20, widthPt: 30, heightPt: 18 }, advancePt: 30,
    };
    const withResource: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [resource] }],
    };
    const calls: unknown[][] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left', textBaseline: 'alphabetic',
      save() {}, restore() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {}, moveTo() {},
      lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;

    paintParagraphLayout(withResource, {
      ctx, scale: 2, dpr: 1,
      resources: { paint: (...args) => { calls.push(args as unknown[]); } },
    });

    expect(calls[0]?.slice(0, 3)).toEqual([
      'chart:body:0', 'chart', { xPt: 12, yPt: 20, widthPt: 30, heightPt: 18 },
    ]);
  });

  it('clips a turned inline picture to its line band only while painting it (word-inline-rotated-picture-line-clip)', () => {
    const layout = node();
    const picture = {
      kind: 'resource' as const, resourceKey: 'image:body:0', resourceKind: 'image' as const,
      range: { start: 0, end: 1 }, bounds: { xPt: 12, yPt: 20, widthPt: 30, heightPt: 40 }, advancePt: 30,
      lineClip: { topPt: -2, bottomPt: 43 },
    };
    const operations: unknown[][] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left', textBaseline: 'alphabetic',
      save() { operations.push(['save']); }, restore() { operations.push(['restore']); },
      rect(...args: number[]) { operations.push(['rect', ...args]); }, clip() { operations.push(['clip']); },
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;

    paintParagraphLayout({ ...layout, lines: [{ ...layout.lines[0]!, placements: [picture] }] }, {
      ctx, scale: 1, dpr: 1, resources: { paint: (key) => { operations.push(['paint', key]); } },
    });

    // The band spans the resource's turn about its centre (reach 50pt) and
    // the line box from 18pt to 63pt.
    expect(operations).toEqual([
      ['save'], ['rect', 27 - 50, 18, 100, 45], ['clip'], ['paint', 'image:body:0'], ['restore'],
    ]);
  });

  it('paints retained per-code-point origins without reconstructing justification', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const justified: ParagraphLayout = {
      ...layout,
      lines: [{
        ...layout.lines[0]!,
        placements: [{
          ...placement,
          text: '観察', range: { start: 0, end: 2 }, advancePt: 40,
          clusters: [
            { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
            { range: { start: 1, end: 2 }, offset: { xPt: 30, yPt: 0 }, advancePt: 10 },
          ],
          paintOps: [{
            text: '観察', range: { start: 0, end: 2 }, offset: { xPt: 0, yPt: 0 },
            letterSpacingPt: 20, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb',
          }],
        }],
      }],
    };
    const calls: Array<readonly [string, number, number]> = [];
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px',
      save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, setLineDash() {},
      fillRect() {}, strokeRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
      fillText(value: string, x: number, y: number) { calls.push([value, x, y]); },
      measureText() { throw new Error('paint must not measure text'); },
    } as unknown as CanvasRenderingContext2D;

    paintParagraphLayout(justified, { ctx, scale: 3, dpr: 2, resources: noPaintResources });

    expect(calls).toEqual([['観察', 10, 18]]);
  });

  it('paints an acquisition-built kashida string without treating inserted tatweels as source text', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const calls: string[] = [];
    const retained: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [{
        ...placement,
        text: 'سلام', range: { start: 0, end: 4 },
        clusters: [{ range: { start: 0, end: 4 }, offset: { xPt: 0, yPt: 0 }, advancePt: 20 }],
        paintOps: [{
          ...placement.paintOps[0]!, text: 'سـلام', range: { start: 0, end: 4 },
          direction: 'rtl', sourceMapping: 'kashida',
        }],
        direction: 'rtl',
      }] }],
    };
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, setLineDash() {},
      fillRect() {}, strokeRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
      fillText(value: string) { calls.push(value); },
      measureText() { throw new Error('paint must not measure kashida'); },
    } as unknown as CanvasRenderingContext2D;

    paintParagraphLayout(retained, { ctx, scale: 1, dpr: 1, resources: noPaintResources });

    expect(calls).toEqual(['سـلام']);
  });

  it('rejects incomplete retained slice geometry instead of falling back to whole-run paint', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const invalid: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [{
        ...placement,
        paintOps: [{
          text: 't', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb',
        }],
      }] }],
    };
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
      save() {}, restore() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;

    expect(() => paintParagraphLayout(invalid, {
      ctx, scale: 1, dpr: 1, resources: noPaintResources,
    }))
      .toThrow('Retained glyph slices are incomplete');
  });

  it('accepts finite negative cluster advances from authored overlapping character spacing', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const overlapping: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [{
        ...placement,
        clusters: [{
          range: { ...placement.range }, offset: { xPt: 0, yPt: 0 }, advancePt: -3,
        }],
      }] }],
    };
    const painted: string[] = [];
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
      letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {}, fillText(text: string) { painted.push(text); },
    } as unknown as CanvasRenderingContext2D;

    expect(() => paintParagraphLayout(overlapping, {
      ctx, scale: 1, dpr: 1, resources: noPaintResources,
    })).not.toThrow();
    expect(painted).toContain('test');
  });

  it('rejects a non-positive retained glyph-local scaleY', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const invalid: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [{
        ...placement,
        paintOps: placement.paintOps.map((operation) => ({ ...operation, scaleY: 0 })),
      }] }],
    };
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
      save() {}, restore() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {}, fillText() {},
    } as unknown as CanvasRenderingContext2D;

    expect(() => paintParagraphLayout(invalid, {
      ctx, scale: 1, dpr: 1, resources: noPaintResources,
    })).toThrow('Retained glyph slices are incomplete (geometry)');
  });

  it('paints retained vertical glyph routing, pitch, and offsets without remeasurement', () => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const vertical: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, range: { start: 0, end: 5 }, placements: [{
        ...placement,
        text: 'A（ー２９',
        range: { start: 0, end: 5 },
        clusters: [{
          range: { start: 0, end: 5 }, offset: { xPt: 0, yPt: 0 }, advancePt: 32,
        }],
        paintOps: [
          {
            text: 'A', range: { start: 0, end: 1 }, offset: { xPt: 1, yPt: 2 },
            glyphOffsetPt: { xPt: 3, yPt: 4 }, glyphOrientation: 'sideways',
            letterSpacingPt: -6, scaleX: .8, direction: 'ltr', kerning: 'auto',
            writingMode: 'vertical-rl',
          },
          {
            text: '︵', range: { start: 1, end: 2 }, offset: { xPt: 8, yPt: 0 },
            glyphOffsetPt: { xPt: 2, yPt: 3 }, glyphOrientation: 'upright',
            letterSpacingPt: 0, scaleX: .8, direction: 'ltr', kerning: 'auto',
            writingMode: 'vertical-rl',
          },
          {
            text: 'ー', range: { start: 2, end: 3 }, offset: { xPt: 16, yPt: 0 },
            glyphOffsetPt: { xPt: 1, yPt: 2 }, glyphOrientation: 'rotate',
            letterSpacingPt: 0, scaleX: .8, direction: 'ltr', kerning: 'auto',
            writingMode: 'vertical-rl',
          },
          {
            text: '２９', range: { start: 3, end: 5 }, offset: { xPt: 24, yPt: 0 },
            glyphOrientation: 'upright',
            letterSpacingPt: -6, scaleX: .8, scaleY: .75, direction: 'ltr', kerning: 'auto',
            writingMode: 'horizontal-tb',
          },
        ],
      }] }],
    };
    const calls: unknown[] = [];
    const ctx = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() { calls.push('save'); }, restore() { calls.push('restore'); },
      translate(x: number, y: number) { calls.push(['translate', x, y]); },
      scale(x: number, y: number) { calls.push(['scale', x, y]); },
      rotate(angle: number) { calls.push(['rotate', angle]); },
      setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {},
      fillText(this: { letterSpacing: string }, text: string, x: number, y: number) {
        calls.push(['fillText', text, x, y, this.letterSpacing]);
      },
      measureText() { throw new Error('retained vertical paint must not measure'); },
    } as unknown as CanvasRenderingContext2D;

    expect(structuredClone(vertical).lines[0]!.placements[0]!.kind).toBe('text');
    expect((structuredClone(vertical).lines[0]!.placements[0] as typeof placement)
      .paintOps[3]?.scaleY).toBe(.75);

    paintParagraphLayout(vertical, { ctx, scale: 1, dpr: 1, resources: noPaintResources });

    expect(calls).toEqual([
      'save', ['translate', 14, 24], ['scale', .8, 1], ['fillText', 'A', 0, 0, '-7.5px'], 'restore',
      'save', ['translate', 18, 18], ['rotate', -Math.PI / 2], ['scale', 1, .8],
      ['fillText', '︵', 2, 3, '0px'], 'restore',
      'save', ['translate', 26, 18], ['scale', .8, 1],
      ['fillText', 'ー', 1, 2, '0px'], 'restore',
      'save', ['translate', 34, 18], ['rotate', -Math.PI / 2], ['scale', .8, .75],
      ['fillText', '２９', 0, 0, '-6px'], 'restore',
    ]);
  });

  it.each([
    ['missing cluster coverage', {
      clusters: [{ range: { start: 0, end: 2 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 }],
    }],
    ['discontinuous ranges', {
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
        { range: { start: 2, end: 4 }, offset: { xPt: 10, yPt: 0 }, advancePt: 10 },
      ],
    }],
    ['non-finite advance', {
      clusters: [
        { range: { start: 0, end: 2 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
        { range: { start: 2, end: 4 }, offset: { xPt: 10, yPt: 0 }, advancePt: Number.NaN },
      ],
    }],
    ['mismatched UTF-16 range', {
      range: { start: 0, end: 3 },
    }],
  ] as const)('rejects %s before text paint', (_name, mutation) => {
    const layout = node();
    const placement = layout.lines[0]!.placements[0]!;
    if (placement.kind !== 'text') throw new Error('fixture must contain text');
    const invalid: ParagraphLayout = {
      ...layout,
      lines: [{ ...layout.lines[0]!, placements: [{ ...placement, ...mutation }] }],
    };
    const ctx = {
      fillStyle: '', font: '', textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
      save() {}, restore() {}, setLineDash() {}, fillRect() {}, strokeRect() {}, beginPath() {},
      moveTo() {}, lineTo() {}, stroke() {}, fillText() { throw new Error('must not paint invalid text'); },
    } as unknown as CanvasRenderingContext2D;

    expect(() => paintParagraphLayout(invalid, {
      ctx, scale: 1, dpr: 1, resources: noPaintResources,
    }))
      .toThrow(/Retained glyph slices are incomplete|UTF-16 text range is inconsistent/);
  });
});
