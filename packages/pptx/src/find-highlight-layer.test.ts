import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  buildPptxHighlightLayer,
  DEFAULT_FIND_HIGHLIGHT,
  DEFAULT_FIND_ACTIVE_HIGHLIGHT,
  type PptxHighlightMatch,
} from './find-highlight-layer.js';
import type { PptxTextRunInfo } from './renderer';

// node env: recording DOM stub (same as text-layer.test.ts). Highlights are
// grouped into one rotated <div> per shape frame; each box sits inside its
// shape div at inShapeX + slice-x.
interface FakeEl {
  tag: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
}
function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
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

const W = 7;
const measureForFont = () => (s: string) => s.length * W;

describe('buildPptxHighlightLayer', () => {
  it('places a box inside the shape div at inShapeX + slice extent', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'the quick', inShapeX: 5, inShapeY: 8, h: 16 })];
    const matches: PptxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 4, end: 9 }], active: false }, // "quick"
    ];
    buildPptxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 960, 540, measureForFont);
    // The overlay container must NOT be pinned to a literal px size — it keeps
    // its `width:100%;height:100%` so it tracks the canvas's actual rendered box.
    expect(layer.style.width ?? '').toBe('');
    expect(layer.style.height ?? '').toBe('');
    // One shape div, containing one box.
    expect(layer.children).toHaveLength(1);
    const shapeDiv = layer.children[0];
    expect(shapeDiv.children).toHaveLength(1);
    const box = shapeDiv.children[0];
    // Box is placed as a PERCENTAGE of the shape frame (shapeW=100, shapeH=50)
    // so it scales with the (percentage-sized, rotatable) group.
    // left = (inShapeX 5 + "the " 28) / 100; top = inShapeY 8 / 50;
    // width = 35 / 100; height = 16 / 50.
    expect(box.style.left).toBe(`${(33 / 100) * 100}%`);
    expect(box.style.top).toBe(`${(8 / 50) * 100}%`);
    expect(box.style.width).toBe(`${(35 / 100) * 100}%`);
    expect(box.style.height).toBe(`${(16 / 50) * 100}%`);
    expect(box.style.background).toBe(DEFAULT_FIND_HIGHLIGHT);
  });

  it('applies the shape rotation to the group div', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc', rotation: 30, textBodyRotation: 90 })];
    const matches: PptxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: true },
    ];
    buildPptxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    const shapeDiv = layer.children[0];
    expect(shapeDiv.style.transform).toBe('rotate(120deg)'); // 30 + 90
    expect(shapeDiv.children[0].style.background).toBe(DEFAULT_FIND_ACTIVE_HIGHLIGHT);
  });

  it('groups boxes from the same shape under one div', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [
      run({ text: 'a', shapeX: 10, shapeY: 20 }),
      run({ text: 'b', shapeX: 10, shapeY: 20 }),
    ];
    const matches: PptxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 1 }], active: false },
      { slices: [{ runIndex: 1, start: 0, end: 1 }], active: false },
    ];
    buildPptxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    expect(layer.children).toHaveLength(1); // one shape group
    expect(layer.children[0].children).toHaveLength(2); // two boxes
  });

  it('clears the layer on rebuild', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'STALE';
    buildPptxHighlightLayer(layer as unknown as HTMLDivElement, [], [], 100, 100, measureForFont);
    expect(layer.innerHTML).toBe('');
    expect(layer.children).toHaveLength(0);
  });
});
