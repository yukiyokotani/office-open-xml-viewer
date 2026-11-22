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

  it.each(['a\u030Ab', 'a\u00ADb'])('does not add mark/SHY pitch to find positions in %s', text => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const natural = () => (value: string) => [...value].filter(c => c === 'a' || c === 'b').length * W;
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement,
      [run({ text, letterSpacingPx: 2 })],
      [{ slices: [{ runIndex: 0, start: 2, end: 3 }], active: false }],
      100, 100, natural);
    expect(layer.children[0].style.left).toBe('9%');
    expect(layer.children[0].style.width).toBe(`${(7 / 100) * 100}%`);
  });

  it('keeps main fractional pitch multiplication for an unchanged slice', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const text = 'AV To WAVE AVATAR';
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement,
      [run({ text, letterSpacingPx: 0.35 })],
      [{ slices: [{ runIndex: 0, start: 3, end: 11 }], active: false }],
      100, 100, measureForFont);
    expect(layer.children[0].style.left).toBe(`${((3 * W + 3 * 0.35) / 100) * 100}%`);
    expect(layer.children[0].style.width).toBe(`${((8 * W + 7 * 0.35) / 100) * 100}%`);
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

describe('highlights in right-to-left runs', () => {
  const hebrew = '\u05d0\u05d1\u05d2';
  const boxes = (runs: DocxTextRunInfo[], matches: DocxHighlightMatch[], measure = measureForFont) => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, 100, 100, measure);
    return layer.children.map((child) => {
      const css = child.style.cssText;
      const rounded = (value: string) => Math.round(Number(value) * 1e9) / 1e9;
      return [rounded(/left:([-\d.e]+)%/.exec(css)![1]!), rounded(/width:([-\d.e]+)%/.exec(css)![1]!)];
    });
  };

  it('leaves out the tracking after the final matched glyph, as in left-to-right runs', () => {
    // 7px glyphs tracked by 2px. Painted left to right: gimel, bet, alef,
    // each with its tracking on its left (after it in reading order).
    const rtl = run({ text: hebrew, direction: 'rtl', letterSpacingPx: 2, w: 27,
      clusterBoxes: [0, 1, 2].map((index) => ({ start: index, end: index + 1, x: (2 - index) * 9, width: 9, rtl: true })) });
    const hit = (start: number, end: number) => boxes([rtl], [{ slices: [{ runIndex: 0, start, end }], active: false }]);
    expect(hit(1, 2)).toEqual([[11, 7]]);
    expect(hit(0, 1)).toEqual([[20, 7]]);
    // Two glyphs keep the tracking between them only.
    expect(hit(0, 2)).toEqual([[11, 16]]);
    // The same slices in a left-to-right run.
    const ltr = run({ text: 'abc', letterSpacingPx: 2, w: 27 });
    expect(boxes([ltr], [{ slices: [{ runIndex: 0, start: 1, end: 2 }], active: false }])).toEqual([[9, 7]]);
  });

  it('subtracts no tracking from a final cluster whose box already leaves it out', () => {
    // A fitted run: its final cluster box is cut to the fitted width.
    const boxesOf = (trackingExcluded: boolean) => boxes([run({ text: 'ab', letterSpacingPx: 2, w: 16,
      clusterBoxes: [{ start: 0, end: 1, x: 0, width: 9, rtl: false },
        { start: 1, end: 2, x: 9, width: 7, rtl: false, ...(trackingExcluded ? { trackingExcluded: true as const } : {}) }] })],
    [{ slices: [{ runIndex: 0, start: 1, end: 2 }], active: false }]);
    expect(boxesOf(true)).toEqual([[9, 7]]);
    expect(boxesOf(false)).toEqual([[9, 5]]);
  });

  it('measures nothing for a run that carries its cluster boxes', () => {
    let measured = 0;
    const measure = () => (value: string) => { measured += value.length; return value.length * W; };
    const length = 300;
    const text = hebrew.repeat(length / 3);
    const clusterBoxes = Array.from({ length }, (_, index) => ({ start: index, end: index + 1,
      x: (length - 1 - index) * W, width: W, rtl: true }));
    const matches = Array.from({ length }, (_, start) => ({ slices: [{ runIndex: 0, start, end: start + 1 }], active: false }));
    const placed = boxes([run({ text, direction: 'rtl', clusterBoxes })], matches, measure);
    expect(placed).toHaveLength(length);
    expect(placed[0]).toEqual([(length - 1) * W, W]);
    expect(measured).toBe(0);
  });
});

describe('tate-chu-yoko highlights', () => {
  it('places the characters side by side across the turned run', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // A tbRl run turned 90deg at (100, 50): '２' and '９' stand side by side
    // across its 8px cell.
    const runs = [run({
      text: '２９', x: 100, y: 50, w: 8, h: 10, fontSize: 8, transform: 'rotate(90deg)', eastAsianVert: true,
      clusterBoxes: [
        { start: 0, end: 1, x: 0, width: 8, rtl: false, y: 8, height: 6 },
        { start: 1, end: 2, x: 0, width: 8, rtl: false, y: 2, height: 6 },
      ],
    })];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs,
      [{ slices: [{ runIndex: 0, start: 0, end: 1 }, { runIndex: 0, start: 1, end: 2 }], active: false }],
      200, 200, measureForFont);
    const [two, nine] = layer.children;
    expect([two!.style.left, two!.style.top]).toEqual(['46%', '25%']);
    expect([nine!.style.left, nine!.style.top]).toEqual(['49%', '25%']);
    expect([two!.style.width, two!.style.height]).toEqual(['4%', '3%']);
  });
});
