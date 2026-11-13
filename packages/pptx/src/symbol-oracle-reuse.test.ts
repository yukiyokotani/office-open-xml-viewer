/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { createSymbolFaceOracleCache } from '@silurus/ooxml-core';
import { deck, para, run, textBox } from './test-support/synthetic-deck.js';
import { initSync, PptxArchive } from './wasm/pptx_parser.js';
import { renderSlideWithEmbeddedFonts } from './renderer.js';
import type { Presentation } from './types';

// The symbol oracle of a presentation (main and worker) is kept across its
// slide renders while the realm's faces are unchanged: each authored symbol
// code is judged by the route sentinels once, not once per render.

const SENTINEL_A = '"__ooxml_route_sentinel_a_6d2e91__"';
const SENTINEL_B = '"__ooxml_route_sentinel_b_6d2e91__"';

/** A realm whose route sentinels are loaded: a sentinel draws every
 * character at its own geometry (route-sentinels.ts); no named family of
 * the deck is installed. Records the sentinel-terminated probes of U+F0B7. */
function realm(faces: { status: string }[]) {
  const probes: string[] = [];
  const canvas = (): unknown => {
    let font = '10px serif';
    const metric = (text: string) => {
      const px = Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
      const scale = px / 1024;
      const last = font.slice(font.lastIndexOf(', ') >= 0 ? font.lastIndexOf(', ') + 2 : font.indexOf('px ') + 3);
      const sentinel = last === SENTINEL_A ? [512, 8192, 716, 204] : last === SENTINEL_B ? [768, 9216, 614, 307] : null;
      if (sentinel) {
        if (text === '') probes.push(font);
        return { width: sentinel[0]! * scale, actualBoundingBoxAscent: sentinel[1]! * scale,
          fontBoundingBoxAscent: sentinel[2]! * scale, fontBoundingBoxDescent: sentinel[3]! * scale,
          actualBoundingBoxDescent: 0, actualBoundingBoxLeft: 0, actualBoundingBoxRight: sentinel[0]! * scale };
      }
      return { width: [...text].length * px * 0.5, actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2, actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: [...text].length * px * 0.5 };
    };
    const noop = () => {};
    const element: Record<string, unknown> = { width: 1, height: 1, style: {}, ownerDocument: { fonts: faces, createElement: canvas } };
    const ctx = new Proxy({
      canvas: element,
      get font() { return font; },
      set font(value: string) { font = value; },
      direction: 'ltr',
      letterSpacing: '0px',
      measureText: metric,
      createLinearGradient: () => ({ addColorStop: noop }),
      createRadialGradient: () => ({ addColorStop: noop }),
      createPattern: () => null,
      getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
      getLineDash: () => [],
      getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(Math.max(0, w * h * 4)) }),
    } as Record<string | symbol, unknown>, {
      get(target, key) { return key in target ? target[key] : noop; },
      set(target, key, value) { target[key] = value; return true; },
    });
    element.getContext = () => ctx;
    return element;
  };
  return { canvas, probes };
}

let pres: Presentation;
beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/pptx_parser_bg.wasm', import.meta.url)) });
  const archive = new PptxArchive(deck(textBox(para(run('Item'),
    '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Symbol"/><a:buChar char=""/></a:pPr>'))));
  try {
    pres = JSON.parse(new TextDecoder().decode(archive.parse())) as Presentation;
  } finally {
    archive.free();
  }
});

describe('PPTX symbol oracle across slide renders', () => {
  const renderSlides = async (count: number, symbolFaces?: ReturnType<typeof createSymbolFaceOracleCache>,
    faces: { status: string }[] = []) => {
    const { canvas, probes } = realm(faces);
    for (let index = 0; index < count; index++) {
      await renderSlideWithEmbeddedFonts(canvas() as HTMLCanvasElement, pres.slides[0]!, pres.slideWidth, pres.slideHeight, {
        width: 320, dpr: 1, majorFont: pres.majorFont, minorFont: pres.minorFont,
        ...(symbolFaces ? { symbolFaces } : {}),
      } as never);
    }
    return probes.length;
  };

  it('judges a symbol bullet once for five slides, and again once the faces change', async () => {
    const once = await renderSlides(1);
    expect(once).toBeGreaterThan(0);
    // Without the presentation's cache every render probes again.
    expect(await renderSlides(5)).toBe(5 * once);
    const faces = [{ status: 'loaded' }];
    const cache = createSymbolFaceOracleCache();
    expect(await renderSlides(5, cache, faces)).toBe(once);
    // A face that finished loading, or a face added, is a new font snapshot.
    const { canvas, probes } = realm(faces);
    const again = async () => {
      await renderSlideWithEmbeddedFonts(canvas() as HTMLCanvasElement, pres.slides[0]!, pres.slideWidth, pres.slideHeight,
        { width: 320, dpr: 1, symbolFaces: cache } as never);
      return probes.length;
    };
    expect(await again()).toBe(0);
    faces[0]!.status = 'unloaded';
    expect(await again()).toBe(once);
    faces.push({ status: 'loaded' });
    expect(await again()).toBe(2 * once);
    expect(await again()).toBe(2 * once);
    // A rewritten matching descriptor too.
    (faces[1] as { unicodeRange?: string }).unicodeRange = 'U+0-7F';
    expect(await again()).toBe(3 * once);
  });
});
