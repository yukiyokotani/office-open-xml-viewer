import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildDocxTextLayer } from './text-layer.js';
import type { DocxTextRunInfo } from './renderer';

// The vitest env is `node` (no document). Following renderer.textbox-image.test.ts's
// vi.stubGlobal pattern, we stub a minimal recording DOM: createElement returns a
// fake element that records its style props (via a cssText setter that parses
// `k:v;` pairs), textContent, and appended children. buildDocxTextLayer lays
// absolutely-positioned <span>s directly on the layer, one per DocxTextRunInfo.

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
    // A cssText setter that parses the `k:v;k:v;` string into individual props,
    // so tests can assert either the raw cssText or a parsed key.
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
  return {
    text: 'X',
    x: 0,
    y: 0,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    ...partial,
  };
}

describe('buildDocxTextLayer (extracted from DocxViewer._buildTextLayer)', () => {
  it('does not pin the layer to a literal px size and clears prior content', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'STALE';
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [], 640, 480);
    expect(layer.innerHTML).toBe(''); // cleared
    // The overlay container keeps its creation `width:100%;height:100%` so it
    // tracks the canvas's ACTUAL rendered box (external CSS may scale the canvas
    // responsively). The build must NOT overwrite it with a literal px size.
    expect(layer.style.width ?? '').toBe('');
    expect(layer.style.height ?? '').toBe('');
    expect(layer.children.length).toBe(0);
  });

  it('lays one absolutely-positioned span per run with the shipped inline styles', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [
      run({ text: 'Hello', x: 12, y: 34, h: 16, font: 'bold 14px Arial' }),
      run({ text: 'World', x: 50, y: 34, h: 16, font: '14px Arial' }),
    ];
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 700, 900);

    expect(layer.children.length).toBe(2);
    const [a, b] = layer.children;
    expect(a.tag).toBe('span');
    expect(a.textContent).toBe('Hello');
    // Shipped style contract: absolute position from run.x/run.y expressed as a
    // PERCENTAGE of cssWidth/cssHeight (so the span tracks the canvas's actual
    // rendered size), the CSS `font` shorthand BEFORE line-height, letter-spacing
    // reset, transparent selectable text.
    expect(a.style.position).toBe('absolute');
    expect(a.style.left).toBe(`${(12 / 700) * 100}%`);
    expect(a.style.top).toBe(`${(34 / 900) * 100}%`);
    // Glyph metrics in units of the layer width (cqw), so the span scales with
    // the page: 14px and 16px of a 700px page.
    expect(a.style.font).toBe(`bold ${(14 / 700) * 100}cqw Arial`);
    expect(a.style['line-height']).toBe(`${(16 / 700) * 100}cqw`);
    expect(a.style['letter-spacing']).toBe('0');
    expect(a.style['white-space']).toBe('pre');
    expect(a.style.color).toBe('transparent');
    expect(a.style.cursor).toBe('text');
    expect(a.style['pointer-events']).toBe('all');
    // Declaration ORDER is load-bearing: the `font` shorthand resets line-height
    // to `normal` in real browsers, so `line-height` must come AFTER `font` in the
    // cssText. The parsed-prop asserts above cannot detect a reorder — pin it on
    // the raw string.
    expect(a.style.cssText).toMatch(/font:[^;]+;line-height:/);
    expect(b.textContent).toBe('World');
    expect(b.style.left).toBe(`${(50 / 700) * 100}%`);
  });

  it('bounds a compressed terminal space overlay before the following run', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [
      run({ text: 'A ', x: 0, w: 4, trailingSpaceCompressionPx: 2 }),
      run({ text: 'B', x: 4, w: 2 }),
    ], 100, 100);

    expect(layer.children[0].style.width).toBe('4%');
    expect(layer.children[0].style.overflow).toBe('clip');
    expect(layer.children[1].style.width ?? '').toBe('');
    expect(layer.children[1].style.overflow ?? '').toBe('');
  });

  // ECMA-376 §17.3.2.10 縦中横 (#836): a tate-chu-yoko run is drawn compressed into
  // ONE em cell (`run.w`), but the selection span lays out at the run's NATURAL
  // font width (~2× for "２９"), so the selection box overshoots into the next cell.
  // With a measurer, the span composes a horizontal scaleX(run.w/naturalWidth) so
  // its extent matches the drawn cell.
  it('compresses an eastAsianVert (縦中横) span to the one-em cell via scaleX', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // "２９" natural 14px (monospace 7px/char), drawn cell run.w=7px ⇒ scaleX 0.5.
    const runs = [
      run({ text: '２９', x: 30, y: 40, w: 7, h: 16, fontSize: 7, font: '7px serif', eastAsianVert: true, transform: 'rotate(90deg)' }),
    ];
    const measureForFont = () => (s: string) => [...s].length * 7;
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 1, 1, undefined, measureForFont);
    const span = layer.children[0];
    // The rotate is preceded by scaleX(0.5) so the natural 14px width compresses to
    // the 7px cell BEFORE rotating into the column (transform-origin top-left).
    expect(span.style.transform).toBe('rotate(90deg) scaleX(0.5)');
    expect(span.style.cssText).toContain('transform-origin:top left');
  });

  it('leaves a 縦中横 span un-scaled when no measurer is supplied (graceful degrade)', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [
      run({ text: '２９', w: 7, fontSize: 7, eastAsianVert: true, transform: 'rotate(90deg)' }),
    ];
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 1, 1);
    // No measurer ⇒ the scale cannot be computed; the span keeps the bare rotate
    // (byte-identical to the pre-#836 overlay — no regression when the caller
    // does not thread a measurer).
    expect(layer.children[0].style.transform).toBe('rotate(90deg)');
  });

  it('does not scale a non-eastAsianVert run even with a measurer', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'Hello', x: 12, font: '14px Arial' })];
    const measureForFont = () => (s: string) => s.length * 8;
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 1, 1, undefined, measureForFont);
    // Ordinary run: no transform at all (horizontal page), so no scaleX sneaks in.
    expect(layer.children[0].style.transform ?? '').toBe('');
  });
});

describe('buildDocxTextLayer — clipped runs (word-cell-content-clip)', () => {
  const layerWith = (runs: DocxTextRunInfo[], onLink?: () => void) => {
    vi.stubGlobal('document', {
      createElement: (t: string) => Object.assign(makeEl(t), { addEventListener() {}, title: '', dataset: {} }),
    });
    const layer = Object.assign(makeEl('div'), { dataset: {} });
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 100, 100, onLink);
    return layer.children;
  };
  const cell = { x: 40, y: 0, width: 30, height: 12 };

  it('places a partly clipped hyperlink in a container of its clip, in page percentages', () => {
    // A 60px link at x=10 whose clip starts 30px into it; the container cuts
    // its hit box and scales with the page, whatever the span's own size.
    const [container] = layerWith([run({
      text: 'linklinklink', x: 10, w: 60,
      hyperlink: { kind: 'external', url: 'https://example.com' } as DocxTextRunInfo['hyperlink'],
      clips: [cell],
    })], () => {});
    expect(container!.style.overflow).toBe('clip');
    expect([container!.style.left, container!.style.width]).toEqual(['40%', '30%']);
    const [span] = container!.children;
    expect(span!.style.left).toBe('-100%');
    expect(span!.style.cursor).toBe('pointer');
  });

  it('puts a wholly clipped run in an empty container and leaves an unclipped run on the layer', () => {
    const empty = { x: 20, y: 0, width: 0, height: 12 };
    const [hidden, open] = layerWith([
      run({ x: 0, w: 10, clips: [empty] }),
      run({ x: 0, w: 10 }),
    ]);
    expect(hidden!.style.width).toBe('0%');
    expect(open!.style['pointer-events']).toBe('all');
  });

  it('keeps document order: a container is shared only by consecutive runs with the same clips', () => {
    const [first, between, second] = layerWith([
      run({ text: 'A', x: 45, clips: [cell] }),
      run({ text: 'C', x: 0 }),
      run({ text: 'B', x: 50, clips: [cell] }),
    ]);
    expect(first!.children.map((child) => child.textContent)).toEqual(['A']);
    expect(between!.textContent).toBe('C');
    expect(second!.children.map((child) => child.textContent)).toEqual(['B']);
  });

  it('nests a turned container for a turned clip and orients the run against it', () => {
    // A text box turned 90deg at (60, 10): the run shares its orientation.
    const box = { x: 60, y: 10, width: 40, height: 20, transform: 'rotate(90deg)' };
    const [outer] = layerWith([run({ x: 60, y: 15, w: 30, transform: 'rotate(90deg)', clips: [cell, box] })]);
    const [inner] = outer!.children;
    expect(inner!.style.overflow).toBe('clip');
    expect(inner!.style.transform).toBe('matrix(0, 1, -1, 0, 0, 0)');
    const [span] = inner!.children;
    // 5px along the box's own x axis, and no orientation left to apply.
    expect([span!.style.left, span!.style.top]).toEqual(['12.5%', '0%']);
    expect(span!.style.transform ?? '').toBe('');
  });
});

describe('buildDocxTextLayer — without a measurer', () => {
  const scaled = () => run({
    // `ABCD` painted at twice its natural advances: 12px per cluster.
    text: 'ABCD', x: 10, w: 48,
    clusterBoxes: [0, 1, 2, 3].map((index) => ({ start: index, end: index + 1, x: index * 12, width: 12, rtl: false })),
  });

  it('measures with a canvas of the document, so a scaled run keeps its painted caret positions', () => {
    const canvas = { getContext: () => ({ font: '', measureText: (text: string) => ({ width: text.length * 6 }) }) };
    vi.stubGlobal('document', { createElement: (tag: string) => (tag === 'canvas' ? canvas : makeEl(tag)) });
    const layer = makeEl('div');
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [scaled()], 100, 100);
    expect(layer.children).toHaveLength(1);
    expect(layer.children[0]!.textContent).toBe('ABCD');
    // 48px painted over 24px natural: the span's advances are scaled to it.
    expect(layer.children[0]!.style.transform).toBe('scaleX(2)');
    expect(layer.children[0]!.style.width ?? '').toBe('');
  });

  it('sizes the span to its painted width where the document cannot measure', () => {
    vi.stubGlobal('document', { createElement: (tag: string) => makeEl(tag) });
    const layer = makeEl('div');
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [scaled()], 100, 100);
    expect(layer.children).toHaveLength(1);
    expect(layer.children[0]!.style.width).toBe('48%');
  });
});

describe('buildDocxTextLayer — inline content without text', () => {
  it('gives a tab no span and leaves the runs around it their indices', () => {
    vi.stubGlobal('document', { createElement: (tag: string) => makeEl(tag) });
    const layer = makeEl('div') as FakeEl & { dataset?: Record<string, string> };
    const withData = (tag: string) => Object.assign(makeEl(tag), { dataset: {} as Record<string, string> });
    vi.stubGlobal('document', { createElement: withData });
    buildDocxTextLayer(Object.assign(layer, { dataset: {} }) as unknown as HTMLDivElement, [
      run({ text: 'A', x: 0, w: 6 }),
      run({ text: '', inlineContent: 'tab', x: 6, w: 60 }),
      run({ text: 'B', x: 66, w: 6 }),
    ], 100, 100);
    expect(layer.children.map((span) => span.textContent)).toEqual(['A', 'B']);
    expect(layer.children.map((span) => (span as unknown as { dataset: Record<string, string> }).dataset.runIndex)).toEqual(['0', '2']);
  });
});

describe('buildDocxTextLayer — text its clips hide whole', () => {
  it('is hidden from assistive technology, and visible text is not', () => {
    const attributes = new Map<object, Record<string, string>>();
    const withAttributes = (tag: string) => {
      const element = Object.assign(makeEl(tag), {
        dataset: {} as Record<string, string>,
        setAttribute(name: string, value: string) { attributes.set(element, { ...attributes.get(element), [name]: value }); },
      });
      return element;
    };
    vi.stubGlobal('document', { createElement: withAttributes });
    const layer = withAttributes('div');
    const clips = [{ x: 0, y: 0, width: 100, height: 12 }];
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [
      run({ text: 'Visible', x: 0, y: 0, w: 42, clips }),
      run({ text: 'Hidden', x: 0, y: 40, w: 36, clips }),
    ], 100, 100);
    const spans: FakeEl[] = [];
    const collect = (element: FakeEl): void => {
      if (element.textContent) spans.push(element);
      element.children.forEach(collect);
    };
    collect(layer);
    expect(spans.map((span) => span.textContent)).toEqual(['Visible', 'Hidden']);
    expect(spans.map((span) => attributes.get(span)?.['aria-hidden'])).toEqual([undefined, 'true']);
  });
});
