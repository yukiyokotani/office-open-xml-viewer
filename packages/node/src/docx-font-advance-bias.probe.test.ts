/**
 * Issue #794 probe harness — per-font advance bias (Regime A), two layers.
 *
 * Layer 1 — ACCEPTANCE, font-present ground-truth parity: on a host that has a
 * document's fonts, wrap positions must match Word. This layer is enforced by
 * the demo/sample-1 fidelity ratchet (Playwright VRT) and the
 * synthetic Word-verified §17.18.44 gates in
 * packages/docx/src/justify-shrink-overshoot.test.ts. It also covers behavior
 * that survives substitution, such as the non-justified drawable trailing-space
 * budget (sample-10 title below).
 *
 * Layer 2 — DIAGNOSTIC ONLY, font-absent reflow: when a requested face is
 * absent, the viewer substitutes the host fallback and REFLOWS, like Word on a
 * machine without the fonts. Line breaks and page counts may then differ from
 * the authoring machine; that is accepted behavior, not a defect (product
 * decision 2026-07-12; issue #855 won't-fix — PR #979's cross-font metric
 * emulation "Regime B" was reverted). The sample-4 page-count check below only
 * LOGS the substituted-environment result; it asserts nothing about parity.
 *
 * Environment: fonts are deliberately NOT registered. macOS-gated (the VRT
 * reference environment); skia + WASM gated like the other probes.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
} from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia ? await loadDocxRendererForTests() : null;

const samplePath = (n: number) => resolve(ROOT, `packages/docx/public/private/docx/sample-${n}.docx`);
const have = (n: number) => existsSync(samplePath(n));

interface Run { text: string; x: number; y: number; w: number; h: number; fontSize: number }

async function parse(n: number) {
  return await docxMod!.materializeDocxDocument(readFileSync(samplePath(n)));
}

/** Page count via retained layout (needs the OffscreenCanvas shim). */
async function pageCount(n: number, width = 595): Promise<number> {
  const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
  try {
    const { createLayoutServices, layoutDocument } = rendererMod!;
    const doc = await parse(n);
    const layoutServices = createLayoutServices(doc);
    return layoutDocument(doc, layoutServices, { currentDateMs: 0 }).pages.length;
  } finally {
    restore.forEach((r) => r());
  }
}

/** A recording ctx proxy over a real skia 2D context: forwards every property /
 *  method to the real ctx (so `measureText`, transforms, fills all behave as in
 *  production) but captures each `fillText`/`strokeText` call as a drawn text run
 *  in ABSOLUTE page coordinates (via the live CTM). Unlike `onTextRun` this also
 *  captures DrawingML textbox (`wps:txbx`) text, which the copyright block
 *  (sample-15) and centred title (sample-10) live in. */
function recordingCtx(real: CanvasRenderingContext2D, sink: Run[]): CanvasRenderingContext2D {
  return new Proxy(real, {
    get(target, prop, receiver) {
      if (prop === 'fillText' || prop === 'strokeText') {
        return (text: string, x: number, y: number, maxW?: number) => {
          if (text) {
            const m = target.getTransform();
            const ax = m.a * x + m.c * y + m.e;
            const ay = m.b * x + m.d * y + m.f;
            const w = target.measureText(text).width * m.a;
            sink.push({ text, x: ax, y: ay, w, h: 0, fontSize: 0 });
          }
          return (target[prop as 'fillText'] as (t: string, x: number, y: number, mw?: number) => void)
            .call(target, text, x, y, maxW as number);
        };
      }
      const v = Reflect.get(target, prop, target);
      return typeof v === 'function' ? v.bind(target) : v;
    },
    set(target, prop, value) {
      return Reflect.set(target, prop, value, target);
    },
  }) as unknown as CanvasRenderingContext2D;
}

/** Visual lines of page `page` (0-based), top-to-bottom, each the concatenated
 *  drawn text sharing a baseline y (captured at the fillText level so textbox
 *  content is included). Column-aware (splits on gutter x-gaps). */
async function pageLines(n: number, page: number, width = 595): Promise<string[]> {
  const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
  try {
    const { createLayoutServices, renderDocumentToCanvas } = rendererMod!;
    const doc = await parse(n);
    const layoutServices = createLayoutServices(doc);
    const runs: Run[] = [];
    const canvas = new Canvas(Math.round(width * 1.5), Math.round(width * 2));
    const realCtx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const recCtx = recordingCtx(realCtx, runs);
    // Return the recording ctx from getContext (leave native canvas.width/height
    // setters intact so skia's private fields are untouched by a Proxy).
    (canvas as unknown as { getContext: () => unknown }).getContext = () => recCtx;
    await renderDocumentToCanvas(doc, canvas as unknown as never, page, {
      width,
      dpr: 1,
      layoutServices,
      currentDate: 0,
      defaultCurrentDateMs: 0,
    });
    // Column-aware line reconstruction: bucket by baseline y, then within a row
    // split into separate visual lines wherever x jumps by more than ~15% of the
    // page width (a column gutter). A 2-column paper (sample-15) otherwise merges
    // its left/right columns into one string per y.
    const gap = width * 0.15;
    const byY = new Map<number, Run[]>();
    for (const r of runs) {
      const key = Math.round(r.y);
      let arr = byY.get(key);
      if (!arr) { arr = []; byY.set(key, arr); }
      arr.push(r);
    }
    const out: { x: number; y: number; text: string }[] = [];
    for (const [y, rs] of byY) {
      const sorted = rs.slice().sort((p, q) => p.x - q.x);
      let cur: Run[] = [];
      let lastRight = -Infinity;
      const flush = () => {
        if (cur.length) out.push({ x: cur[0].x, y, text: cur.map((r) => r.text).join('') });
        cur = [];
      };
      for (const r of sorted) {
        if (cur.length && r.x - lastRight > gap) flush();
        cur.push(r);
        lastRight = r.x + r.w;
      }
      flush();
    }
    // Order the reconstructed lines by column (x band) then y: left column top→
    // bottom, then right column. This keeps a paragraph's lines contiguous.
    return out
      .sort((a, b) => (a.x < width * 0.5 ? 0 : 1) - (b.x < width * 0.5 ? 0 : 1) || a.y - b.y)
      .map((l) => l.text);
  } finally {
    restore.forEach((r) => r());
  }
}

const macos = process.platform === 'darwin';
const gate = !!skia && !!docxMod && !!rendererMod && macos;

describe.skipIf(!gate)('issue #794 — per-font advance-bias probes', () => {
  it.skipIf(!have(4))('DIAGNOSTIC (non-acceptance): sample-4 font-absent page count', async () => {
    // Meiryo UI is absent here, so the full-width substitute reflows and the
    // document may paginate differently from Word on the authoring machine
    // (Word: 1 page; the substitute typically yields 2). Accepted behavior —
    // issue #855 won't-fix. Logged for regen bookkeeping only.
    const pages = await pageCount(4);
    // eslint-disable-next-line no-console
    console.log(`[#794 diagnostic] sample-4 font-absent pageCount = ${pages} (Word with fonts: 1)`);
    expect(pages).toBeGreaterThanOrEqual(1);
  });

  // The sample-15 #698 narrow justified column is encoded as the synthetic
  // Word-verified gate in packages/docx/src/justify-shrink-overshoot.test.ts
  // (3 tokens / 2 lines on justified lines). The real sample-15 copyright block
  // is a page-bottom-anchored <wps:txbx> that this body render path does not
  // draw, so it cannot be line-counted here; the synthetic gate +
  // demo/sample-1 VRT ratchet are its acceptance.

  it.skipIf(!have(10))('ACCEPTANCE: sample-10 p1 centred title stays on one line', async (context) => {
    const lines = await pageLines(10, 0, 595);
    const norm = (l: string) => l.replace(/\s+/g, '');
    const title = '横幹連合コンファレンスサンプル原稿';
    if (!norm(lines.join('')).includes(title)) {
      context.skip('the installed local fixture is not the centred-title corpus');
      return;
    }
    // Word truth (sample-10.pdf p1): the centred main title is one line:
    //   「第 11 回横幹連合コンファレンスサンプル原稿」
    // A substituted MS Mincho over-measures the CJK title by ~+5.96px; Word keeps
    // it on ONE line. A centred (non-justified) line keeps the drawable
    // trailing-space shrink budget, which absorbs the overflow — this pin guards
    // that retained path. Historical failure mode: the tail 「サンプル原稿」
    // wraps to a second line.
    const titleLine = lines.find((l) => norm(l).includes(title));
    // eslint-disable-next-line no-console
    console.log(`[#794 C3] sample-10 title one-line? ${!!titleLine}\n` +
      lines.slice(0, 8).map((l, i) => `  ${i}: ${JSON.stringify(l)}`).join('\n'));
    expect(titleLine, 'full title 第…コンファレンスサンプル原稿 on a single visual line').toBeTruthy();
  });
});
