/**
 * Table cells inside a vertical (tbRl) section — ECMA-376 §17.6.20 +
 * §17.4.80/§17.18.37, issue #988 batch-3 adjudication ④.
 *
 * Word ground truth (the batch-3 vertical-table fixture's PDF, Letter portrait
 * 612×792 pt, 1 in margins; two 1-cell tables, `tcW` 1440 twips fixed +
 * `tblLayout` fixed, same long CJK string; TABLE 1 `trHeight hRule="exact"`
 * 2880 twips, TABLE 2 auto):
 *
 *   - cell text lays out HORIZONTALLY (5 chars/line, wrapping downward) — the
 *     section's vertical direction does not propagate into cells;
 *   - fixed `tcW` 1 in ⇒ PHYSICAL horizontal width (border box 72.5 pt);
 *   - TABLE 1 (exact): border box exactly 2 in tall — x [417.1, 489.6],
 *     y [72, 216.5] — content clipped at the border (line 9 sliced mid-glyph);
 *   - TABLE 2 (auto): border box grew to enclose all 17 lines (293.8 pt at Yu
 *     Mincho's 17.2 pt natural line; a substituted face scales the pitch).
 *
 * TABLE 1's border box lands at its natural flow position and matches the PDF
 * on BOTH axes, so it is asserted exactly (±2.5 pt). TABLE 2's physical
 * placement is Word-IDIOSYNCRATIC (see the RE-ADJUDICATION below) — so for it
 * only the spec-grounded invariants are asserted: physical width, top-margin
 * pinning, and auto GROWTH well past the 144.5 pt exact height.
 *
 * ── RE-ADJUDICATION (2026-07-12, issue #988 follow-up) ──────────────────────
 * The user re-observed sample-52 and reported (1) "cells still vertical / cut
 * off" and (2) "Word keeps vertical text BELOW the table, we put it beside it".
 *
 * (1) does NOT reproduce. Cell text renders HORIZONTAL/upright in BOTH the node
 * (skia-canvas) render AND the browser (Chromium via Storybook) — verified with
 * onTextRun (transform:undefined, 5-char w≈60 runs advancing downward), the
 * cell-glyph ctm (identity — the +90° page rotation is un-done), and a pixel
 * band analysis. The earlier report was a STALE Storybook build predating #999
 * (merge 6e2a2cc); the low line-pitch of a substituted CJK face makes stacked
 * horizontal lines LOOK like vertical columns until magnified. So #999's
 * horizontal-cell behaviour is CORRECT and is regression-guarded below.
 *
 * (2) is real but NON-CAUSAL, so it stays a won't-fix (narrowed & measured).
 * Word GT (red-border boxes + pdftotext -bbox), physical page x[72,540] y[72,720]:
 *   • T1 (exact) box x[417,489] y[72,216]  — top-anchored, LEFT.
 *   • T2 (auto)  box x[504,576] y[72,365]  — top-anchored, RIGHT, OVERHANGS the
 *     right content margin (576 > 540).
 *   • p0 heading col x[519,531] y[366,563] and p1 col x[497,510] y[366,621]
 *     resume BELOW T2 (start at T2's bottom, in T2's x-band).
 *   • p4 col x[449,462] y[217,516] and p6 col x[425,438] y[217,620] resume
 *     BELOW T1 (start at T1's bottom, in T1's x-band).
 * So Word DOES flow vertical text below each top-anchored table. BUT the pairing
 * is non-causal: text BEFORE T1 (p0,p1) sits under the LATER table T2, and text
 * after T1 (p4,p6) sits under the EARLIER table T1 — text-segment i is capped by
 * table (n+1−i), the tables progress LEFT→RIGHT (reverse of the RTL text), and
 * T2 overhangs the margin. A later table displacing earlier text cannot arise in
 * a single forward layout pass; reproducing it needs a page-level place-all-
 * tables → register-exclusions → reflow model, and even that leaves the reversed
 * pairing + overhang undocumented. §17.6.20/§17.4.80 say nothing about it and a
 * plain block table has no wrapTopAndBottom, so a forward-only "text-below-table"
 * approximation would be sample-fitting (spec-first violation) AND a *different*
 * wrong layout (it would put p6 under T2, not T1). Independent review (Codex
 * gpt-5.6-sol) concurred. Decision: keep the renderer's RTL block-flow placement,
 * assert only the spec-grounded invariants, and CHARACTERIZE the residual (our
 * T2 lands LEFT of T1, the reverse of Word) so a future adjudication sweep on a
 * purpose-built fixture matrix notices any change.
 *
 * CI-safe: gated on docx WASM + skia-canvas + the PRIVATE sample + a macOS JP
 * font; skips when any is absent.
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
  new URL('../../docx/public/private/docx/sample-52.docx', import.meta.url),
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

interface RunInfo { text: string; x: number; y: number; w: number; transform?: string }

describe.skipIf(!skia || !docxMod || !rendererMod || !havePrereqs)(
  'docx vertical table cells render upright/horizontal (§17.6.20 + §17.4.80, #988 ④)',
  () => {
    it('matches the Word PDF: exact clips at 2 in, auto grows, cells horizontal', async () => {
      for (const fam of ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP']) {
        FontLibrary.use(fam, [MINCHO]);
      }
      const { materializeDocxDocument } = docxMod as { materializeDocxDocument: (b: Uint8Array) => Any };
      const { renderDocumentToCanvas } = rendererMod!;
      const doc = await materializeDocxDocument(readFileSync(SAMPLE));
      const W = Math.round(doc.section.pageWidth);
      const H = Math.round(doc.section.pageHeight);
      const canvas = new Canvas(W, H);
      const runs: RunInfo[] = [];
      const rImg = installImageBitmapShim(factory);
      const rOff = installOffscreenCanvasShim(factory);
      try {
        await renderDocumentToCanvas(doc, canvas as Any, 0, {
          dpr: 1,
          width: doc.section.pageWidth,
          onTextRun: (r: Any) =>
            runs.push({ text: r.text, x: r.x, y: r.y, w: r.w, transform: r.transform }),
        });
      } finally {
        rOff();
        rImg();
      }
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
      const { data } = ctx.getImageData(0, 0, W, H);

      // ── Red (C00000) border boxes ────────────────────────────────────────
      const isRed = (x: number, y: number): boolean => {
        const i = (y * W + x) * 4;
        return data[i] > 120 && data[i + 1] < 80 && data[i + 2] < 80;
      };
      // Vertical border lines: columns with a long red span.
      const vlines: Array<{ x: number; y0: number; y1: number }> = [];
      for (let x = 0; x < W; x++) {
        let count = 0, y0 = -1, y1 = -1;
        for (let y = 0; y < H; y++) {
          if (isRed(x, y)) {
            count++;
            if (y0 < 0) y0 = y;
            y1 = y;
          }
        }
        if (count > 40) vlines.push({ x, y0, y1 });
      }
      // Cluster adjacent columns into border edges.
      const edges: Array<{ x: number; y0: number; y1: number }> = [];
      for (const v of vlines) {
        const last = edges[edges.length - 1];
        if (last && v.x - last.x <= 2) {
          last.y0 = Math.min(last.y0, v.y0);
          last.y1 = Math.max(last.y1, v.y1);
          continue;
        }
        edges.push({ ...v });
      }
      expect(edges.length, 'two tables ⇒ four vertical border edges').toBe(4);

      // Pair edges into boxes by matching y extents (left/right of one table
      // share the same y span).
      const boxes: Array<{ x0: number; x1: number; y0: number; y1: number }> = [];
      const used = new Set<number>();
      for (let a = 0; a < edges.length; a++) {
        if (used.has(a)) continue;
        for (let b = a + 1; b < edges.length; b++) {
          if (used.has(b)) continue;
          if (Math.abs(edges[a].y1 - edges[b].y1) <= 3 && Math.abs(edges[a].y0 - edges[b].y0) <= 3) {
            boxes.push({
              x0: Math.min(edges[a].x, edges[b].x),
              x1: Math.max(edges[a].x, edges[b].x),
              y0: Math.min(edges[a].y0, edges[b].y0),
              y1: Math.max(edges[a].y1, edges[b].y1),
            });
            used.add(a).add(b);
            break;
          }
        }
      }
      expect(boxes.length, 'two table border boxes').toBe(2);
      const exactBox = boxes.find((b) => Math.abs(b.y1 - b.y0 - 144.5) < 6);
      const autoBox = boxes.find((b) => b !== exactBox);
      expect(exactBox, 'a 2-in-tall exact border box').toBeDefined();
      expect(autoBox).toBeDefined();

      // TABLE 1 (exact): matches the Word PDF on BOTH axes (±2.5 pt).
      expect(Math.abs(exactBox!.x0 - 417.1), 'exact left').toBeLessThanOrEqual(2.5);
      expect(Math.abs(exactBox!.x1 - 489.6), 'exact right').toBeLessThanOrEqual(2.5);
      expect(Math.abs(exactBox!.y0 - 72), 'exact top').toBeLessThanOrEqual(2.5);
      expect(Math.abs(exactBox!.y1 - 216.5), 'exact bottom').toBeLessThanOrEqual(2.5);

      // TABLE 2 (auto): spec-grounded invariants only (placement along the
      // flow axis is the Word-idiosyncratic partial-won't-fix candidate).
      expect(Math.abs(autoBox!.x1 - autoBox!.x0 - 72.5), 'auto physical width').toBeLessThanOrEqual(2.5);
      expect(Math.abs(autoBox!.y0 - 72), 'auto pinned at the top margin').toBeLessThanOrEqual(2.5);
      expect(autoBox!.y1 - autoBox!.y0, 'auto row grew past the exact 144.5').toBeGreaterThan(250);
      // Both tables share the SAME physical top edge (top content margin) — the
      // top-anchoring the RE-ADJUDICATION confirms matches Word for BOTH tables.
      expect(Math.abs(exactBox!.y0 - autoBox!.y0), 'both tables top-anchored together').toBeLessThanOrEqual(2.5);
      // CHARACTERIZATION of the documented won't-fix residual (see the header
      // RE-ADJUDICATION): our RTL block flow lays the LATER auto table entirely
      // to the LEFT of the earlier exact table. Word GT is the REVERSE (auto at
      // x[504,576], right of exact at x[417,489], overhanging the margin). If a
      // future change reproduces Word's page-level exclusion/reflow model this
      // flips and must be re-adjudicated on a purpose-built fixture matrix.
      expect(autoBox!.x1, 'residual: auto table lands LEFT of exact (reverse of Word)').toBeLessThan(exactBox!.x0);

      // ── Clip: no cell ink below TABLE 1's exact bottom border ───────────
      let inkBelow = 0;
      for (let y = Math.round(exactBox!.y1) + 3; y < Math.round(exactBox!.y1) + 90; y++) {
        for (let x = Math.round(exactBox!.x0) + 2; x < Math.round(exactBox!.x1) - 2; x++) {
          const i = (y * W + x) * 4;
          if (data[i] < 200 && data[i + 1] < 200 && data[i + 2] < 200) inkBelow++;
        }
      }
      expect(inkBelow, 'exact row content is clipped at the border').toBe(0);

      // ── Cell text is HORIZONTAL: same x per line, y advancing downward ──
      const cellLines = runs.filter(
        (r) => r.transform === undefined && Math.abs(r.w - 60) < 3 && r.x > exactBox!.x0 && r.x < exactBox!.x1,
      );
      expect(cellLines.length, 'horizontal cell lines inside TABLE 1').toBeGreaterThanOrEqual(8);
      expect(new Set(cellLines.map((r) => Math.round(r.x))).size).toBe(1);
      const ys = cellLines.map((r) => r.y);
      for (let i = 1; i < ys.length; i++) expect(ys[i]).toBeGreaterThan(ys[i - 1]);
      // Body text outside the tables keeps the +90° vertical overlay transform.
      expect(runs.some((r) => r.transform === 'rotate(90deg)')).toBe(true);
    }, 120000);
  },
);
