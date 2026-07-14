import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildDocxTextLayer, getDocxTextRunAtNode } from './text-layer.js';
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
    expect(a.style.font).toBe('bold 14px Arial');
    expect(a.style['line-height']).toBe('16px');
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

  it('maps a selection-layer span back to its rendered source run', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const sourceRefs = [{
      partName: 'word/document.xml',
      path: [{ localName: 't', index: 0 }],
      textStart: 0,
      textEnd: 1,
      sourceStart: 0,
      sourceEnd: 1,
    }];
    const sourceRun = run({ sourceRefs });

    buildDocxTextLayer(layer as unknown as HTMLDivElement, [sourceRun], 100, 100);

    expect(getDocxTextRunAtNode(layer.children[0] as unknown as Node)).toBe(sourceRun);
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
