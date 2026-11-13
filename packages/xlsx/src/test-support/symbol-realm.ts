/**
 * A Canvas realm whose route sentinels are loaded (core route-sentinels.ts),
 * for tests of symbol-font text (symbol-paint.ts): a sentinel draws every
 * character at its own geometry, an `installed` family draws the characters
 * its coverage predicate accepts, and every other family is absent, so text
 * it cannot draw falls back to a generic face. Records every measured and
 * painted text with the font it was measured or painted in.
 */
const SENTINELS: Record<string, readonly [number, number, number, number]> = {
  __ooxml_route_sentinel_a_6d2e91__: [512, 8192, 716, 204],
  __ooxml_route_sentinel_b_6d2e91__: [768, 9216, 614, 307],
};

/** A symbol-encoded face: its F0-shifted codes, and its bare codes through
 * the symbol-cmap remapping. */
export const SYMBOL_ENCODED = (ch: string): boolean => {
  const cp = ch.codePointAt(0) ?? 0;
  return (cp >= 0xf020 && cp <= 0xf0ff) || (cp >= 0x20 && cp <= 0xff);
};

export interface PaintRecord { text: string; font: string; x: number; y: number }

export interface SymbolRealm {
  readonly ctx: CanvasRenderingContext2D;
  readonly fills: PaintRecord[];
  readonly measures: PaintRecord[];
}

function families(font: string): string[] {
  const list = /^(?:.*?)\d*\.?\d+px\s+(.+)$/su.exec(font)?.[1] ?? '';
  return list.split(',').map((entry) => entry.trim().replace(/^"(.*)"$/u, '$1').toLowerCase());
}

/** `installed`: lower-case family → whether it draws a character. */
export function symbolRealm(
  installed: Readonly<Record<string, (ch: string) => boolean>> = {},
  size: { width: number; height: number } = { width: 1200, height: 800 },
): SymbolRealm {
  const fills: PaintRecord[] = [];
  const measures: PaintRecord[] = [];
  const fontSet: FontFace[] = [];
  const make = (record: boolean, width: number, height: number): Record<string, unknown> => {
    let font = '10px sans-serif';
    let fillStyle = '#000000';
    let ink: [number, number, number] | null = null;
    const metric = (text: string) => {
      const px = Number.parseFloat(/(\d*\.?\d+)px/u.exec(font)?.[1] ?? '10');
      const n = [...text].length;
      for (const family of families(font)) {
        const sentinel = SENTINELS[family];
        if (sentinel) {
          const scale = px / 1024;
          return { width: n * sentinel[0] * scale, actualBoundingBoxAscent: sentinel[1] * scale,
            fontBoundingBoxAscent: sentinel[2] * scale, fontBoundingBoxDescent: sentinel[3] * scale,
            actualBoundingBoxDescent: 0, actualBoundingBoxLeft: 0, actualBoundingBoxRight: n * sentinel[0] * scale };
        }
        const draws = installed[family];
        if (draws && n > 0 && [...text].every(draws)) {
          return { width: n * px * 0.75, actualBoundingBoxAscent: px * 0.7, actualBoundingBoxDescent: px * 0.1,
            fontBoundingBoxAscent: px * 0.9, fontBoundingBoxDescent: px * 0.25, actualBoundingBoxLeft: 0,
            actualBoundingBoxRight: n * px * 0.75 };
        }
      }
      return { width: n * px * 0.5, actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2, actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: n * px * 0.5 };
    };
    const noop = () => {};
    const element: Record<string, unknown> = {
      width, height, style: {},
      ownerDocument: { fonts: fontSet, createElement: () => make(false, 48, 48) },
    };
    const target: Record<string | symbol, unknown> = {
      canvas: element,
      get font() { return font; },
      set font(value: string) { font = value; },
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string) { fillStyle = value; },
      direction: 'ltr',
      letterSpacing: '0px',
      globalAlpha: 1,
      textAlign: 'left',
      textBaseline: 'alphabetic',
      measureText: (text: string) => {
        if (record) measures.push({ text, font, x: 0, y: 0 });
        return metric(text);
      },
      fillText: (text: string, x: number, y: number) => {
        if (record) fills.push({ text, font, x, y });
        const hex = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/iu.exec(fillStyle);
        ink = hex ? [1, 2, 3].map((i) => parseInt(hex[i]!, 16) * 17) as [number, number, number] : null;
      },
      clearRect: () => { ink = null; },
      getImageData: (_x: number, _y: number, w: number, h: number) => {
        const data = new Uint8ClampedArray(Math.max(0, w * h * 4));
        if (ink) for (let i = 0; i < data.length; i += 4) data.set([...ink, 255], i);
        return { data };
      },
      createLinearGradient: () => ({ addColorStop: noop }),
      createRadialGradient: () => ({ addColorStop: noop }),
      createPattern: () => null,
      getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
      getLineDash: () => [],
    };
    const ctx = new Proxy(target, {
      get(object, key) { return key in object ? object[key] : noop; },
      set(object, key, value) { object[key] = value; return true; },
    });
    element.getContext = () => ctx;
    return element;
  };
  const canvas = make(true, size.width, size.height);
  return { ctx: (canvas.getContext as () => CanvasRenderingContext2D)(), fills, measures };
}
