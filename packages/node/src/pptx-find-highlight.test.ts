import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { installImageBitmapShim, installOffscreenCanvasShim, type NodeCanvasFactory } from './render.ts';
import { importForTests, loadSkiaForTests } from './test-imports';

/**
 * IX2 findText END-TO-END on the real demo pptx: parse → render each slide
 * capturing real `onTextRun` runs → PptxFindController → find a known word →
 * assert its `{ slide }` + reconstructed text, then feed the real match slices
 * to buildPptxHighlightLayer with a REAL skia measurer and assert the box lands
 * inside its shape group on the drawn glyph run.
 */
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async () => {
    throw new Error('unused');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const RENDERER_PATH = resolve(ROOT, 'packages/pptx/src/renderer.ts');
const FIND_PATH = resolve(ROOT, 'packages/pptx/src/find.ts');
const HIGHLIGHT_PATH = resolve(ROOT, 'packages/pptx/src/find-highlight-layer.ts');

const pptxMod = skia ? await importForTests(() => import('./pptx.ts'), './pptx.ts (pptx WASM)') : null;
const rendererMod = skia
  ? await importForTests(() => import(RENDERER_PATH), 'packages/pptx/src/renderer.ts')
  : null;
const findMod = skia ? await importForTests(() => import(FIND_PATH), 'packages/pptx/src/find.ts') : null;
const highlightMod = skia
  ? await importForTests(() => import(HIGHLIGHT_PATH), 'packages/pptx/src/find-highlight-layer.ts')
  : null;

const DEMO = resolve(ROOT, 'packages/pptx/public/demo/sample-1.pptx');
const haveDemo = existsSync(DEMO);

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

type Run = {
  text: string;
  inShapeX: number;
  inShapeY: number;
  w: number;
  h: number;
  font: string;
  fontSize: number;
  shapeX: number;
  shapeY: number;
  shapeW: number;
  shapeH: number;
  rotation: number;
};

describe.skipIf(!skia || !pptxMod || !rendererMod || !findMod || !highlightMod || !haveDemo)(
  'IX2 pptx findText + highlight on the demo fixture',
  () => {
    async function collectSlides(): Promise<Run[][]> {
      const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
      try {
        const { materializePptxPresentation } = pptxMod as { materializePptxPresentation: (b: Uint8Array) => Promise<{ slides: unknown[]; slideWidth: number; slideHeight: number; defaultTextColor?: string; majorFont?: unknown; minorFont?: unknown; hlinkColor?: unknown }> };
        const { renderSlide } = rendererMod as {
          renderSlide: (
            canvas: unknown,
            slide: unknown,
            sw: number,
            sh: number,
            opts: Record<string, unknown>,
            onTextRun?: (r: Run) => void,
          ) => Promise<unknown>;
        };
        const pres = await materializePptxPresentation(readFileSync(DEMO));
        const out: Run[][] = [];
        for (let s = 0; s < pres.slides.length; s++) {
          const canvas = new Canvas(960, 540);
          const runs: Run[] = [];
          await renderSlide(
            canvas,
            pres.slides[s],
            pres.slideWidth,
            pres.slideHeight,
            {
              width: 960,
              dpr: 1,
              defaultTextColor: pres.defaultTextColor,
              majorFont: pres.majorFont,
              minorFont: pres.minorFont,
              hlinkColor: pres.hlinkColor ?? null,
              fetchMedia: async () => new Blob([]),
              fetchImage: async () => new Blob([]),
              skipMediaControls: true,
            },
            (r: Run) => runs.push(r),
          );
          out.push(runs);
        }
        return out;
      } finally {
        restore.forEach((r) => r());
      }
    }

    it('finds a known word and reports its slide + original-case text', async () => {
      const slides = await collectSlides();
      const { PptxFindController } = findMod as {
        PptxFindController: new (
          count: () => number,
          collect: (s: number) => Promise<Run[]>,
        ) => {
          find: (q: string) => Promise<{ text: string; location: { slide: number } }[]>;
          next: () => unknown;
          slideHighlights: (s: number) => { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[];
        };
      };
      const ctrl = new PptxFindController(
        () => slides.length,
        (s) => Promise.resolve(slides[s] ?? []),
      );
      const matches = await ctrl.find('forest');
      expect(matches.length).toBeGreaterThanOrEqual(1);
      expect(matches[0].text.toLowerCase()).toBe('forest');
      expect(matches[0].location.slide).toBeGreaterThanOrEqual(0);
    });

    it('matches whole words that end a line or a shape, in a whole-word find and in a text-highlight set', async () => {
      const slides = await collectSlides();
      type Match = { text: string; location: { slide: number }; group?: string; term?: string };
      const { PptxFindController } = findMod as {
        PptxFindController: new (
          count: () => number,
          collect: (s: number) => Promise<Run[]>,
        ) => {
          find: (q: string, o?: { wholeWord?: boolean }) => Promise<Match[]>;
          highlight: (groups: unknown[]) => Promise<Match[]>;
        };
      };
      // A real seam where one run ends in a letter, the next starts with one,
      // and layout separates them (another line or another shape).
      const separated = (previous: Run, next: Run): boolean =>
        previous.shapeX !== next.shapeX || previous.shapeY !== next.shapeY ||
        Math.abs(previous.inShapeY - next.inShapeY) > 0.5;
      let seam: { slide: number; word: string } | undefined;
      for (let slide = 0; slide < slides.length && !seam; slide++) {
        const runs = slides[slide].filter((run) => run.text.length > 0);
        for (let i = 1; i < runs.length; i++) {
          const word = /\p{L}{4,}$/u.exec(runs[i - 1].text)?.[0];
          const alone = word !== undefined && !/\p{L}/u.test(runs[i - 1].text.slice(-word.length - 1, -word.length));
          if (word && alone && /^\p{L}/u.test(runs[i].text) && separated(runs[i - 1], runs[i])) {
            seam = { slide, word };
            break;
          }
        }
      }
      expect(seam).toBeDefined();
      const ctrl = new PptxFindController(
        () => slides.length,
        (s) => Promise.resolve(slides[s] ?? []),
      );
      const plain = (await ctrl.find(seam!.word)).filter((match) => match.location.slide === seam!.slide);
      const whole = (await ctrl.find(seam!.word, { wholeWord: true }))
        .filter((match) => match.location.slide === seam!.slide);
      expect(whole.length).toBeGreaterThanOrEqual(1);
      expect(whole.length).toBeLessThanOrEqual(plain.length);

      const highlighted = await ctrl.highlight([{ id: 'seam', terms: [seam!.word], wholeWord: true }]);
      expect(highlighted.filter((match) => match.location.slide === seam!.slide)).toHaveLength(whole.length);
      expect(highlighted.every((match) => match.group === 'seam' && match.term === seam!.word)).toBe(true);
    });

    it('places the highlight box inside its shape group on the drawn run', async () => {
      const slides = await collectSlides();
      const { PptxFindController } = findMod as {
        PptxFindController: new (
          count: () => number,
          collect: (s: number) => Promise<Run[]>,
        ) => {
          find: (q: string) => Promise<{ location: { slide: number } }[]>;
          slideHighlights: (s: number) => { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[];
        };
      };
      const { buildPptxHighlightLayer } = highlightMod as {
        buildPptxHighlightLayer: (
          layer: unknown,
          runs: Run[],
          matches: { slices: { runIndex: number; start: number; end: number }[]; active: boolean }[],
          w: number,
          h: number,
          measureForFont: (font: string) => (s: string) => number,
        ) => void;
      };
      const ctrl = new PptxFindController(
        () => slides.length,
        (s) => Promise.resolve(slides[s] ?? []),
      );
      const matches = await ctrl.find('forest');
      const slide = matches[0].location.slide;
      const runs = slides[slide];
      const highlights = ctrl.slideHighlights(slide);

      const mcanvas = new Canvas(1, 1);
      const mctx = mcanvas.getContext('2d') as unknown as { font: string; measureText: (s: string) => { width: number } };
      const measureForFont = (font: string) => {
        mctx.font = font;
        return (s: string) => mctx.measureText(s).width;
      };

      const layer = makeEl();
      const prevDoc = (globalThis as { document?: unknown }).document;
      (globalThis as { document?: unknown }).document = { createElement: () => makeEl() };
      try {
        buildPptxHighlightLayer(layer as unknown as HTMLDivElement, runs, highlights, 960, 540, measureForFont);
      } finally {
        (globalThis as { document?: unknown }).document = prevDoc;
      }

      // One or more shape group divs, each with box children.
      expect(layer.children.length).toBeGreaterThanOrEqual(1);
      const firstSlice = highlights.find((h) => h.slices.length)?.slices[0];
      expect(firstSlice).toBeDefined();
      const run = runs[firstSlice!.runIndex];
      const shapeDiv = layer.children.find((d) => d.children.length > 0);
      expect(shapeDiv).toBeDefined();
      const box = shapeDiv!.children[0];
      // buildPptxHighlightLayer positions a box as a PERCENTAGE of its shape
      // frame (run.shapeW/run.shapeH) — the % denominators — so convert the
      // parsed percentage back to px against that same basis before comparing
      // to the run's real (px) in-shape geometry.
      const left = (parseFloat(box.style.left) / 100) * run.shapeW;
      const width = (parseFloat(box.style.width) / 100) * run.shapeW;
      const top = (parseFloat(box.style.top) / 100) * run.shapeH;
      // The box sits in the shape's own frame at inShapeX + slice extent.
      expect(width).toBeGreaterThan(0);
      expect(left).toBeGreaterThanOrEqual(run.inShapeX - 0.5);
      expect(top).toBeCloseTo(run.inShapeY, 1);
      expect(box.style.background).toMatch(/rgba/);
    });
  },
);
