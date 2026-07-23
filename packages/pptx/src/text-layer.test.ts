import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildPptxTextLayer } from './text-layer.js';
import type { PptxTextRunInfo } from './renderer';

// node env: no document. Recording DOM stub (see docx text-layer.test.ts). pptx
// groups runs into one positioned <div> per shape frame (keyed by shape geom +
// total rotation) and applies a CSS rotate() when the shape is rotated; each
// run's <span> is absolutely positioned INSIDE its shape div.

interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
}
function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    children: [],
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const i = decl.indexOf(':');
            if (i > 0) target[decl.slice(0, i).trim()] = decl.slice(i + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
    }),
    appendChild(c: FakeEl) {
      this.children.push(c);
    },
  };
  return el;
}
afterEach(() => vi.unstubAllGlobals());

function run(p: Partial<PptxTextRunInfo>): PptxTextRunInfo {
  return {
    text: 'X',
    inShapeX: 0,
    inShapeY: 0,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    shapeX: 0,
    shapeY: 0,
    shapeW: 100,
    shapeH: 50,
    rotation: 0,
    ...p,
  };
}

describe('buildPptxTextLayer (extracted from PptxViewer._buildTextLayer)', () => {
  it('sizes the layer and groups two runs of the same shape under one div', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'STALE';
    const runs = [
      run({ text: 'A', inShapeX: 2, inShapeY: 4, shapeX: 10, shapeY: 20, shapeW: 200, shapeH: 80 }),
      run({ text: 'B', inShapeX: 2, inShapeY: 24, shapeX: 10, shapeY: 20, shapeW: 200, shapeH: 80 }),
    ];
    buildPptxTextLayer(layer as unknown as HTMLDivElement, runs, 960, 540);

    expect(layer.innerHTML).toBe('');
    // The overlay container must NOT be pinned to a literal px size: it stays at
    // its creation `width:100%;height:100%` so it tracks the canvas's ACTUAL
    // rendered box (external CSS may scale the canvas down responsively). The
    // build must leave the container size untouched.
    expect(layer.style.width ?? '').toBe('');
    expect(layer.style.height ?? '').toBe('');
    // Same shape frame + rotation ⇒ ONE group div, holding both spans.
    expect(layer.children.length).toBe(1);
    const shapeDiv = layer.children[0];
    expect(shapeDiv.tag).toBe('div');
    expect(shapeDiv.style.position).toBe('absolute');
    // The shape frame is placed as a PERCENTAGE of cssWidth/cssHeight so it
    // scales with the container: left 10/960, top 20/540, w 200/960, h 80/540.
    expect(shapeDiv.style.left).toBe(`${(10 / 960) * 100}%`);
    expect(shapeDiv.style.top).toBe(`${(20 / 540) * 100}%`);
    expect(shapeDiv.style.width).toBe(`${(200 / 960) * 100}%`);
    expect(shapeDiv.style.height).toBe(`${(80 / 540) * 100}%`);
    expect(shapeDiv.style.overflow).toBe('hidden');
    expect(shapeDiv.children.map((c) => c.textContent)).toEqual(['A', 'B']);
    // Span uses the shipped `font` shorthand + line-height + letter-spacing reset.
    // Its position inside the shape div is a PERCENTAGE of the shape's own
    // box (shapeW/shapeH), so it scales with the rotated group.
    const spanA = shapeDiv.children[0];
    expect(spanA.style.position).toBe('absolute');
    expect(spanA.style.left).toBe(`${(2 / 200) * 100}%`);
    expect(spanA.style.top).toBe(`${(4 / 80) * 100}%`);
    expect(spanA.style.font).toBe('12px serif');
    expect(spanA.style['line-height']).toBe('12px');
    expect(spanA.style['letter-spacing']).toBe('0');
    expect(spanA.style.color).toBe('transparent');
  });

  it('splits runs of different shapes into separate group divs', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildPptxTextLayer(
      layer as unknown as HTMLDivElement,
      [run({ text: 'A', shapeX: 0, shapeY: 0 }), run({ text: 'B', shapeX: 300, shapeY: 100 })],
      960,
      540,
    );
    expect(layer.children.length).toBe(2);
  });

  it('applies a CSS rotate() to a rotated shape (rotation + textBodyRotation)', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // rotation 30 + vertical text body 90 ⇒ totalRot 120.
    buildPptxTextLayer(layer as unknown as HTMLDivElement, [run({ text: 'R', rotation: 30, textBodyRotation: 90 })], 960, 540);
    const shapeDiv = layer.children[0];
    // The verbatim viewer code sets these via DOM style *properties*
    // (`div.style.transformOrigin` / `.transform`), not via the `cssText`
    // string, so the recording stub records them under their camelCase keys.
    expect(shapeDiv.style.transformOrigin).toBe('center center');
    expect(shapeDiv.style.transform).toBe('rotate(120deg)');
  });

  it('does not set a transform for an unrotated shape', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildPptxTextLayer(layer as unknown as HTMLDivElement, [run({ text: 'U', rotation: 0 })], 960, 540);
    const shapeDiv = layer.children[0];
    expect(shapeDiv.style.transform ?? '').toBe('');
  });
});
