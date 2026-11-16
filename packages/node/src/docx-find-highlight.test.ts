import { beforeAll, describe, it, expect } from 'vitest';
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

/**
 * IX2 findText END-TO-END on the real (committed) demo docx: parse → retain →
 * render each page capturing the real `onTextRun` runs → DocxFindController →
 * find a known word → assert its `{ page }` location + reconstructed text, then
 * feed the real match slices to buildDocxHighlightLayer with a REAL skia
 * text-measurer and assert the highlight box lands on the drawn glyph run
 * (numeric geometry, not eyeballing). This proves the whole find + highlight
 * pipeline against a real document, not stubs.
 */
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async () => {
    throw new Error('unused');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const FIND_PATH = resolve(ROOT, 'packages/docx/src/find.ts');
const HIGHLIGHT_PATH = resolve(ROOT, 'packages/docx/src/find-highlight-layer.ts');

const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia ? await loadDocxRendererForTests() : null;
const findMod = skia ? await importForTests(() => import(FIND_PATH), 'packages/docx/src/find.ts') : null;
const highlightMod = skia
  ? await importForTests(() => import(HIGHLIGHT_PATH), 'packages/docx/src/find-highlight-layer.ts')
  : null;

const DEMO = resolve(ROOT, 'packages/docx/public/demo/sample-1.docx');
const haveDemo = existsSync(DEMO);

// Minimal recording DOM used by buildDocxHighlightLayer in node (no jsdom).
interface FakeEl {
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
}
function makeEl(): FakeEl {
  const style: Record<string, string> = {};
  return {
    innerHTML: '',
    children: [],
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(t, p: string, v: string) {
        if (p === 'cssText') {
          for (const d of v.split(';')) {
            const i = d.indexOf(':');
            if (i > 0) t[d.slice(0, i).trim()] = d.slice(i + 1).trim();
          }
          t.cssText = v;
        } else t[p] = v;
        return true;
      },
    }),
    appendChild(c: FakeEl) {
      this.children.push(c);
    },
  };
}

describe.skipIf(!skia || !docxMod || !rendererMod || !findMod || !highlightMod || !haveDemo)(
  'IX2 docx findText + highlight on the demo fixture',
  () => {
    type Run = {
      text: string;
      x: number;
      y: number;
      w: number;
      h: number;
      font: string;
      fontSize: number;
    };

    // Render every page and collect its runs, exactly as DocxViewer does. Keep
    // the immutable recording for both assertions: rebuilding the complete
    // retained document twice made these behavior-identical tests contend with
    // the rest of the full-suite workers. Preparing the real DOCX in beforeAll
    // also keeps suite setup out of Vitest's shorter per-test timeout.
    async function collectAllPages(): Promise<Run[][]> {
      const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
      try {
        const { materializeDocxDocument } = docxMod!;
        const { createLayoutServices, layoutDocument, renderDocumentToCanvas } = rendererMod!;
        const doc = await materializeDocxDocument(readFileSync(DEMO));
        const layoutServices = createLayoutServices(doc);
        // layoutDocument consumes the normalized internal epoch value, while
        // renderDocumentToCanvas accepts the public Date|number override plus
        // the load-time default used to key its retained-layout variant store.
        // Pin all three views of the same instant so this probe is deterministic.
        const layout = layoutDocument(doc, layoutServices, { currentDateMs: 0 });
        const perPage: Run[][] = [];
        for (let p = 0; p < layout.pages.length; p++) {
          const canvas = new Canvas(800, 1000);
          const runs: Run[] = [];
          await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, p, {
            width: 800,
            dpr: 1,
            layoutServices,
            currentDate: 0,
            defaultCurrentDateMs: 0,
            onTextRun: (r: Run) => runs.push(r),
          });
          perPage.push(runs);
        }
        return perPage;
      } finally {
        restore.forEach((r) => r());
      }
    }

    let collectedPages: Run[][] | null = null;
    beforeAll(async () => {
      collectedPages = await collectAllPages();
    }, 30_000);

    function pages(): Run[][] {
      if (!collectedPages) throw new Error('DOCX run fixture was not prepared');
      return collectedPages;
    }

    it('finds a known word and reports its page + original-case text', async () => {
      const recordedPages = pages();
      const { DocxFindController } = findMod as {
        DocxFindController: new (
          count: () => number,
          collect: (p: number) => Promise<Run[]>,
        ) => {
          find: (q: string, o?: unknown) => Promise<{ matchIndex: number; text: string; location: { page: number } }[]>;
          next: () => { location: { page: number } } | null;
          pageHighlights: (p: number) => { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[];
        };
      };
      const controller = new DocxFindController(
        () => recordedPages.length,
        (p) => Promise.resolve(recordedPages[p] ?? []),
      );

      // "Cathedral" appears in the demo's feature headline.
      const matches = await controller.find('cathedral');
      expect(matches.length).toBeGreaterThanOrEqual(1);
      // Case-insensitive match, but the reported text is the document's case.
      expect(matches[0].text.toLowerCase()).toBe('cathedral');
      expect(matches[0].location.page).toBeGreaterThanOrEqual(0);
      expect(matches[0].location.page).toBeLessThan(recordedPages.length);

      // Activate the first match; its page must carry a highlight marked active.
      controller.next();
      const hp = controller.pageHighlights(matches[0].location.page);
      expect(hp.some((h) => h.active)).toBe(true);
    });

    it('matches whole words that end a paragraph, in a whole-word find and in a text-highlight set', async () => {
      const recordedPages = pages();
      type Match = { matchIndex: number; text: string; location: { page: number }; group?: string; term?: string };
      const { DocxFindController } = findMod as {
        DocxFindController: new (
          count: () => number,
          collect: (p: number) => Promise<Run[]>,
        ) => {
          find: (q: string, o?: { wholeWord?: boolean }) => Promise<Match[]>;
          highlight: (groups: unknown[]) => Promise<Match[]>;
          highlightNext: () => Match | null;
          pageHighlights: (p: number) => { active: boolean; color?: string }[];
        };
      };
      const controller = new DocxFindController(
        () => recordedPages.length,
        (p) => Promise.resolve(recordedPages[p] ?? []),
      );

      // "Living" ends the headline's first paragraph and "Cathedral" is the next
      // one, with no character between them in the page's run list.
      const joined = recordedPages.map((runs) => runs.map((run) => run.text).join('')).join('\n');
      expect(joined).toContain('LivingCathedral');
      expect((await controller.find('living', { wholeWord: true })).map((match) => match.text)).toContain('Living');
      expect((await controller.find('cathedral', { wholeWord: true })).length).toBeGreaterThanOrEqual(1);
      // A word fragment is found by a plain find only.
      expect((await controller.find('cathed')).length).toBeGreaterThanOrEqual(1);
      expect(await controller.find('cathed', { wholeWord: true })).toEqual([]);

      const matches = await controller.highlight([
        { id: 'headline', terms: ['living', 'cathedral'], wholeWord: true, color: 'rgba(0, 0, 255, 0.3)' },
        { id: 'fragment', terms: ['cathed'] },
      ]);
      const groupsOf = (text: string) =>
        matches.filter((match) => match.text.toLowerCase() === text).map((match) => match.group);
      expect(groupsOf('living')).toContain('headline');
      expect(groupsOf('cathedral')).toContain('headline');
      expect(groupsOf('cathed')).toContain('fragment');
      expect(matches.map((match) => match.matchIndex)).toEqual(matches.map((_, index) => index));
      // The find above stays as it was: no matches, and the highlights are drawn.
      const active = controller.highlightNext();
      expect(active?.matchIndex).toBe(0);
      const boxes = controller.pageHighlights(active!.location.page);
      expect(boxes.filter((box) => box.active)).toHaveLength(1);
      expect(boxes.every((box) => box.color !== undefined)).toBe(true);
    });

    it('places the highlight box on the drawn glyph run (real skia measure)', async () => {
      const recordedPages = pages();
      const { DocxFindController } = findMod as {
        DocxFindController: new (
          count: () => number,
          collect: (p: number) => Promise<Run[]>,
        ) => {
          find: (q: string) => Promise<{ location: { page: number } }[]>;
          next: () => unknown;
          pageHighlights: (p: number) => { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[];
        };
      };
      const { buildDocxHighlightLayer } = highlightMod as {
        buildDocxHighlightLayer: (
          layer: unknown,
          runs: Run[],
          matches: { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[],
          w: number,
          h: number,
          measureForFont: (font: string) => (s: string) => number,
        ) => void;
      };
      const controller = new DocxFindController(
        () => recordedPages.length,
        (p) => Promise.resolve(recordedPages[p] ?? []),
      );
      const matches = await controller.find('cathedral');
      const page = matches[0].location.page;
      const runs = recordedPages[page];
      const highlights = controller.pageHighlights(page);

      // Real skia measurer, primed per font (same as the viewer's _measureForFont).
      const measureCanvas = new Canvas(1, 1);
      const mctx = measureCanvas.getContext('2d') as unknown as {
        font: string;
        measureText: (s: string) => { width: number };
      };
      const measureForFont = (font: string) => {
        mctx.font = font;
        return (s: string) => mctx.measureText(s).width;
      };

      const layer = makeEl();
      // buildDocxHighlightLayer creates each box via document.createElement; no
      // jsdom in node, so provide a minimal recording stub for the call.
      const prevDoc = (globalThis as { document?: unknown }).document;
      (globalThis as { document?: unknown }).document = { createElement: () => makeEl() };
      try {
        buildDocxHighlightLayer(
          layer as unknown as HTMLDivElement,
          runs,
          highlights,
          800,
          1000,
          measureForFont,
        );
      } finally {
        (globalThis as { document?: unknown }).document = prevDoc;
      }

      expect(layer.children.length).toBeGreaterThanOrEqual(1);
      // Every box must sit within its matched run's drawn extent: left ≥ run.x,
      // right ≤ run.x + run.w (+1px slack for sub-pixel rounding), top == run.y,
      // and a positive width. This is the pixel-level "highlight lands on the
      // glyphs" assertion.
      //
      // buildDocxHighlightLayer positions each box as a PERCENTAGE of the
      // cssWidth/cssHeight passed above (800/1000) — the % denominators — so
      // convert the parsed percentage back to px against that same basis before
      // comparing to the run's real (px) render geometry.
      const CSS_W = 800;
      const CSS_H = 1000;
      const firstSlice = highlights.find((h) => h.slices.length)?.slices[0];
      expect(firstSlice).toBeDefined();
      const run = runs[firstSlice!.runIndex];
      const box = layer.children[0];
      const left = (parseFloat(box.style.left) / 100) * CSS_W;
      const width = (parseFloat(box.style.width) / 100) * CSS_W;
      const top = (parseFloat(box.style.top) / 100) * CSS_H;
      expect(width).toBeGreaterThan(0);
      expect(left).toBeGreaterThanOrEqual(run.x - 0.5);
      expect(left + width).toBeLessThanOrEqual(run.x + run.w + 1);
      expect(top).toBeCloseTo(run.y, 1);
      // The box background is the (translucent) highlight fill, not empty.
      expect(box.style.background).toMatch(/rgba/);
    });
  },
);
