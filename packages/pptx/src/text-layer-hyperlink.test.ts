import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildPptxTextLayer } from './text-layer.js';
import type { PptxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

// IX1 — clickable hyperlink overlay. No DOM environment is installed (the pptx
// unit tests run in node, see the sibling text-layer.test.ts), so this uses the
// same recording DOM stub, EXTENDED with `title`, `addEventListener`, and a
// `click()` that dispatches to registered `click` listeners. That lets us assert
// the span's click handler fires with the exact HyperlinkTarget without pulling
// in jsdom/happy-dom.

interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  title: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  _listeners: Record<string, ((e: { preventDefault(): void }) => void)[]>;
  appendChild(c: FakeEl): void;
  addEventListener(type: string, fn: (e: { preventDefault(): void }) => void): void;
  click(): void;
}
function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    title: '',
    children: [],
    _listeners: {},
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
    addEventListener(type: string, fn: (e: { preventDefault(): void }) => void) {
      (this._listeners[type] ??= []).push(fn);
    },
    click() {
      for (const fn of this._listeners.click ?? []) fn({ preventDefault() {} });
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

const EXTERNAL: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };

describe('buildPptxTextLayer — clickable hyperlink spans (IX1)', () => {
  it('makes a hyperlink run clickable (pointer + title + click handler) and leaves a plain run plain', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const runs = [
      run({ text: 'link', inShapeX: 2, inShapeY: 4, hyperlink: EXTERNAL }),
      run({ text: 'plain', inShapeX: 40, inShapeY: 4 }),
    ];
    buildPptxTextLayer(layer as unknown as HTMLDivElement, runs, 960, 540, onHyperlinkClick);

    // Both runs share one shape frame ⇒ one group div holding two spans.
    const shapeDiv = layer.children[0];
    const [linkSpan, plainSpan] = shapeDiv.children;

    // Link span: pointer cursor, tooltip = the URL, transparent glyph colour.
    expect(linkSpan.textContent).toBe('link');
    expect(linkSpan.style.cursor).toBe('pointer');
    expect(linkSpan.title).toBe('https://example.com/');
    expect(linkSpan.style.color).toBe('transparent');

    // Plain span: default text cursor, no tooltip.
    expect(plainSpan.textContent).toBe('plain');
    expect(plainSpan.style.cursor).toBe('text');
    expect(plainSpan.title).toBe('');

    // Clicking the link span invokes the handler with the exact target; the
    // plain run registered no listener, so clicking it does nothing.
    linkSpan.click();
    expect(onHyperlinkClick).toHaveBeenCalledTimes(1);
    expect(onHyperlinkClick).toHaveBeenCalledWith(EXTERNAL);

    plainSpan.click();
    expect(onHyperlinkClick).toHaveBeenCalledTimes(1);
  });

  it('leaves hyperlink spans non-interactive when no onHyperlinkClick is supplied', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildPptxTextLayer(layer as unknown as HTMLDivElement, [run({ text: 'link', hyperlink: EXTERNAL })], 960, 540);
    const span = layer.children[0].children[0];
    // Without a handler the span stays a plain, selectable text span.
    expect(span.style.cursor).toBe('text');
    expect(span.title).toBe('');
    expect(span._listeners.click ?? []).toHaveLength(0);
  });

  it('uses the internal ref as the tooltip for an internal slide-jump link', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const internal: HyperlinkTarget = { kind: 'internal', ref: '../slides/slide3.xml' };
    buildPptxTextLayer(layer as unknown as HTMLDivElement, [run({ text: 'go', hyperlink: internal })], 960, 540, onHyperlinkClick);
    const span = layer.children[0].children[0];
    expect(span.style.cursor).toBe('pointer');
    expect(span.title).toBe('../slides/slide3.xml');
    span.click();
    expect(onHyperlinkClick).toHaveBeenCalledWith(internal);
  });
});
