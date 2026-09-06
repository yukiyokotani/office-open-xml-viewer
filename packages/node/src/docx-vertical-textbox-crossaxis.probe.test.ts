/**
 * Vertical text-box cross-axis spacing on the REAL private sample-53 —
 * ECMA-376 §20.1.10.83 (bodyPr vert) + §21.1.2.1.1 (insets) + §17.3.3.25 (ruby).
 *
 * Word PDF ground truth (Letter 612×792 pt; all boxes lIns/rIns = 7.2 pt,
 * tIns/bIns = 3.6 pt; label runs 11 pt, body/base runs 15 pt, ruby 7.5 pt):
 *
 *   box (a) `vert="mongolianVert"` at x ∈ [28.8, 144.0] — columns flow L→R.
 *     Word puts the first (label) column's glyph box at x = 38.56..50.49, i.e.
 *     the flow-start gap MIRRORS the eaVert twin (box b: right gap 9.27 pt vs
 *     a: left gap 9.76 pt). The two historical bugs pulled the label ~7 pt left:
 *     (1) the reflected column stack started at physical bIns (3.6) instead of
 *     lIns (7.2), and (2) the line CENTERLINE was not mirrored within its band,
 *     so Yu-Mincho-style asc-heavy leading pushed ink LEFT of the lIns content
 *     edge (x < 36).
 *
 *   box (b) `vert="eaVert"` at x ∈ [172.8, 288.0] — columns flow R→L; a label
 *     line then a ruby-bearing base line. Word renders THREE distinct ink
 *     bands: base column 235..247, ruby band 254..259, label column 267..277.
 *     The historical bug dropped the layoutLines() ruby ascent reservation in
 *     the textbox re-measure, so the base column advanced by the plain line
 *     pitch and the furigana overprinted the label column (two merged bands).
 *
 * Assertions are font-robust invariants of those two fixes (the harness aliases
 * Hiragino for Yu Mincho, so exact Word ink positions shift by a few pt):
 *   a-1. box (a) first-column ink starts at/after the lIns content edge (36.0).
 *   a-2. …and not absurdly far in (< 48; Word ink starts at 39).
 *   b-1. box (b) has THREE separated ink bands (base / ruby / label).
 *   b-2. the ruby band clears the label band by ≥ 2 pt (no overprint).
 *
 * CI-safe: gated on docx WASM + skia-canvas + the PRIVATE sample + a macOS JP
 * font; skips when any is absent (never hard-fails for the private file).
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const SAMPLE = fileURLToPath(
  new URL('../../docx/public/private/docx/sample-53.docx', import.meta.url),
);
const MINCHO = '/System/Library/Fonts/ヒラギノ明朝 ProN.ttc';
const havePrereqs = existsSync(SAMPLE) && existsSync(MINCHO);

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

/** Contiguous x-ranges holding dark text ink inside [x0,x1) × [y0,y1). */
function inkBands(
  data: Uint8ClampedArray,
  W: number,
  x0: number,
  x1: number,
  y0: number,
  y1: number,
): Array<{ start: number; end: number }> {
  const bands: Array<{ start: number; end: number }> = [];
  let cur: { start: number; end: number } | null = null;
  for (let x = x0; x < x1; x++) {
    let ink = 0;
    for (let y = y0; y < y1; y++) {
      const i = (y * W + x) * 4;
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (lum < 150) ink++;
    }
    if (ink > 0) {
      if (cur && x - cur.end <= 3) cur.end = x;
      else bands.push((cur = { start: x, end: x }));
    }
  }
  return bands;
}

describe.skipIf(!skia || !docxMod || !rendererMod || !havePrereqs)(
  'docx vertical textbox cross-axis spacing (sample-53: mongolianVert inset+mirror, eaVert ruby reservation)',
  () => {
    it('places the mongolian first column inside lIns and separates base/ruby/label bands', async () => {
      for (const fam of ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP']) {
        FontLibrary.use(fam, [MINCHO]);
      }
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
      const { renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(readFileSync(SAMPLE));
      const W = Math.round(doc.section.pageWidth);
      const H = Math.round(doc.section.pageHeight);
      const canvas = new Canvas(W, H);
      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      try {
        await renderDocumentToCanvas(doc, canvas as Any, 0, {
          dpr: 1,
          width: doc.section.pageWidth,
        });
      } finally {
        rOff();
        rImg();
      }
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
      const { data } = ctx.getImageData(0, 0, W, H);

      // box (a) mongolianVert x ∈ [28.8, 144]: interior scan (skip the border).
      const a = inkBands(data, W, 30, 143, 80, 550);
      expect(a.length, `box(a) ink bands ${JSON.stringify(a)}`).toBeGreaterThanOrEqual(2);
      // a-1: no trespass over the lIns content edge (28.8 + 7.2 = 36; −1 for AA fringe).
      expect(a[0].start, `box(a) first ink ${a[0].start} respects lIns edge 36`).toBeGreaterThanOrEqual(35);
      // a-2: still starts near the edge (Word ink @39; Hiragino metrics shift a few pt).
      expect(a[0].start, `box(a) first ink ${a[0].start} stays near the edge`).toBeLessThan(48);

      // box (b) eaVert + ruby x ∈ [172.8, 288]: THREE separated bands.
      const b = inkBands(data, W, 174, 287, 80, 550);
      expect(b.length, `box(b) ink bands ${JSON.stringify(b)}`).toBe(3);
      const [, ruby, label] = b;
      // b-2: furigana clears the label column (the bug overprinted it).
      expect(label.start - ruby.end, `ruby→label gap ${label.start - ruby.end}`).toBeGreaterThanOrEqual(2);
    }, 120000);
  },
);
