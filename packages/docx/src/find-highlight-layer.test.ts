import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  buildDocxHighlightLayer,
  DEFAULT_FIND_HIGHLIGHT,
  DEFAULT_FIND_ACTIVE_HIGHLIGHT,
  type DocxHighlightMatch,
} from './find-highlight-layer.js';
import type { DocxTextRunInfo } from './renderer';

/**
 * IX2 docx highlight overlay. Same node-env fake-DOM stub as the text-layer
 * tests (no jsdom in this workspace). buildDocxHighlightLayer draws one absolute
 * box per matched run-slice, positioned from the run's x/y/h plus the slice's
 * measured horizontal extent, with the active match in the emphasis colour.
 */
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
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
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

function run(partial: Partial<DocxTextRunInfo>): DocxTextRunInfo {
  return { text: 'X', x: 0, y: 0, w: 10, h: 12, fontSize: 12, font: '12px serif', ...partial };
}

// Monospace measurer: each char is 7px wide.
const W = 7;
const measureForFont = () => (s: string) => s.length * W;

describe('buildDocxHighlightLayer', () => {
  it('draws one box per slice, positioned from run x/y + measured extent', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'the quick', x: 100, y: 50, h: 16 })];
    // Match "quick" = slice [4, 9).
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 4, end: 9 }], active: false },
    ];
    buildDocxHighlightLayer(
      layer as unknown as HTMLDivElement,
      runs,
      matches,
      700,
      900,
      measureForFont,
    );
    // The overlay container keeps its `width:100%;height:100%` so it tracks the
    // canvas's actual rendered box; the build must not pin it to literal px.
    expect(layer.style.width ?? '').toBe('');
    expect(layer.style.height ?? '').toBe('');
    expect(layer.children).toHaveLength(1);
    const box = layer.children[0];
    // Box positioned as a PERCENTAGE of cssWidth/cssHeight so it scales with the
    // container: left = (run.x 100 + prefix "the " 28) / 700; top = 50 / 900;
    // width = 35 / 700; height = 16 / 900.
    expect(box.style.left).toBe(`${(128 / 700) * 100}%`);
    expect(box.style.top).toBe(`${(50 / 900) * 100}%`);
    expect(box.style.width).toBe(`${(35 / 700) * 100}%`);
    expect(box.style.height).toBe(`${(16 / 900) * 100}%`);
    expect(box.style.background).toBe(DEFAULT_FIND_HIGHLIGHT);
    expect(box.style['pointer-events']).toBe('none');
  });

  it('uses the active colour for the active match', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: true },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 1, 1, measureForFont);
    expect(layer.children[0].style.background).toBe(DEFAULT_FIND_ACTIVE_HIGHLIGHT);
  });

  it("paints a match in its term's colour, but the active match in the active colour", () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc def' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: false, color: 'rgba(255, 0, 0, 0.4)' },
      { slices: [{ runIndex: 0, start: 4, end: 7 }], active: true, color: 'rgba(255, 0, 0, 0.4)' },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 1, 1, measureForFont);
    expect(layer.children.map((box) => box.style.background)).toEqual(['rgba(255, 0, 0, 0.4)', DEFAULT_FIND_ACTIVE_HIGHLIGHT]);
  });

  it('draws a box per run for a cross-run match', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'Hel', x: 0 }), run({ text: 'lo', x: 21 })];
    const matches: DocxHighlightMatch[] = [
      {
        slices: [
          { runIndex: 0, start: 0, end: 3 },
          { runIndex: 1, start: 0, end: 2 },
        ],
        active: false,
      },
    ];
    // cssWidth = 100 so the percentage reads directly as the px value.
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    expect(layer.children).toHaveLength(2);
    expect(layer.children[0].style.left).toBe('0%'); // run 0 origin (0/100)
    expect(layer.children[1].style.left).toBe(`${(21 / 100) * 100}%`); // run 1 origin
  });

  it('carries the run transform for a vertical (tbRl) page', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc', transform: 'rotate(90deg)' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: false },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 1, 1, measureForFont);
    expect(layer.children[0].style.transform).toBe('rotate(90deg)');
  });

  // ECMA-376 §17.3.2.10 縦中横 (#836): a tate-chu-yoko run is drawn compressed into
  // ONE em cell (`run.w`), but the slice extent measures the run's NATURAL width
  // (~2× for "２９"), so the highlight box overshoots into the following cell along
  // the column. The overlay must clamp the extent to the drawn cell.
  it('clamps an eastAsianVert (縦中横) run highlight to the one-em cell, not natural width', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // "２９": natural width 2·W=14px (monospace stub), drawn cell run.w=7px (one em
    // at fontSize 7). A whole-run match should span the 7px CELL, not 14px.
    const runs = [
      run({ text: '２９', x: 30, y: 40, w: 7, h: 16, fontSize: 7, eastAsianVert: true, transform: 'rotate(90deg)' }),
    ];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 2 }], active: false },
    ];
    // cssWidth = 100 so a percentage reads directly as the px value.
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    expect(layer.children).toHaveLength(1);
    const box = layer.children[0];
    // Clamped: box starts at the run origin and spans exactly the one-em cell.
    expect(box.style.left).toBe(`${(30 / 100) * 100}%`);
    expect(box.style.width).toBe(`${(7 / 100) * 100}%`); // the cell, NOT 14px natural
    // The rotate transform (vertical page) is preserved.
    expect(box.style.transform).toBe('rotate(90deg)');
  });

  it('scales a PARTIAL slice of a 縦中横 run proportionally within the cell', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // "２９" natural 14px, cell 7px ⇒ scale 0.5. A slice of just "９" (the 2nd
    // glyph) is natural [7,14) ⇒ clamped to [3.5, 7): left offset 3.5, width 3.5.
    const runs = [
      run({ text: '２９', x: 0, y: 0, w: 7, h: 16, fontSize: 7, eastAsianVert: true }),
    ];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 1, end: 2 }], active: false },
    ];
    // cssWidth = 100 so a percentage reads directly as the px value.
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    const box = layer.children[0];
    expect(box.style.left).toBe(`${(3.5 / 100) * 100}%`);
    expect(box.style.width).toBe(`${(3.5 / 100) * 100}%`);
  });

  it('leaves a non-eastAsianVert run measured at natural width (unchanged)', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // Same text/w but NOT flagged 縦中横 ⇒ natural 14px extent. w is irrelevant.
    const runs = [run({ text: '２９', x: 0, w: 7, fontSize: 7 })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 2 }], active: false },
    ];
    // cssWidth = 100 so the 14px natural extent reads as 14%.
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measureForFont);
    expect(layer.children[0].style.width).toBe(`${(14 / 100) * 100}%`);
  });

  it('clears the layer and skips zero-width slices', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'stale';
    const runs = [run({ text: 'abc' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 1, end: 1 }], active: false }, // degenerate
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 1, 1, measureForFont);
    expect(layer.innerHTML).toBe('');
    expect(layer.children).toHaveLength(0);
  });
});
