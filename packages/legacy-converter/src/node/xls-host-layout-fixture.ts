// A metric-only Node canvas factory for the XLSX host-layout measurement: its
// digits measure `width` CSS pixels, and each measured font string is kept.
import type { NodeCanvasFactory } from './node-facade.js';

export function digitWidthFactory(width: number | (() => number)): NodeCanvasFactory & { readonly fonts: string[] } {
  const fonts: string[] = [];
  const measure = typeof width === 'function' ? width : () => width;
  const context = {
    set font(value: string) { fonts.push(value); },
    get font() { return fonts.at(-1) ?? ''; },
    save() {}, restore() {},
    measureText: () => ({ width: measure() }),
  };
  return {
    fonts,
    createCanvas: (w, h) => ({ width: w, height: h, getContext: () => context }),
    loadImage: async () => { throw new Error('metric-only factory decodes no images'); },
  };
}
