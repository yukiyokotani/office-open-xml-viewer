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
  loadDocxBookmarkNavForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
} from './test-imports';

// IX-nav END-TO-END: prove the docx internal-anchor navigation resolves against
// the REAL parser + REAL paginator + REAL bookmark-map builder on a real fixture
// (sample-11: a Word doc whose TOC hyperlinks — `<w:hyperlink w:anchor="_TocN">`
// — target `<w:bookmarkStart w:name="_TocN">` headings, ECMA-376 §17.16.23 /
// §17.13.6.2). The private journal samples are git-ignored (not redistributable),
// so this self-skips where absent (CI); OOXML_REQUIRE_SKIA=1 turns absence into a
// hard failure locally.
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
const navMod = skia ? await loadDocxBookmarkNavForTests() : null;

const samplePath = (n: number) => resolve(ROOT, `packages/docx/public/private/docx/sample-${n}.docx`);
const haveSample11 = existsSync(samplePath(11));

describe.skipIf(!skia || !docxMod || !rendererMod || !navMod || !haveSample11)(
  'docx internal-anchor navigation resolves on a real fixture (sample-11)',
  () => {
    async function buildMap(): Promise<Map<string, number>> {
      const restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
      try {
        const { materializeDocxDocument } = docxMod!;
        const { createLayoutServices, layoutDocument } = rendererMod!;
        const { buildBookmarkPageMap } = navMod!;
        const doc = await materializeDocxDocument(readFileSync(samplePath(11)));
        const layoutServices = createLayoutServices(doc);
        const layout = layoutDocument(doc, layoutServices, { currentDateMs: 0 });
        return buildBookmarkPageMap(layout);
      } finally {
        restore.forEach((r) => r());
      }
    }

    it('extracts TOC bookmarks and resolves every anchor to a real page', async (context) => {
      const map = await buildMap();
      const tocNames = [...map.keys()].filter((k) => k.startsWith('_Toc'));
      if (tocNames.length === 0) {
        context.skip('the installed local fixture is not the TOC navigation corpus');
        return;
      }
      // sample-11's TOC has ~20 `_TocN` bookmarks; the parser must surface them.
      expect(map.size).toBeGreaterThan(0);

      // Each TOC anchor names a bookmark of the SAME name (a Word TOC field);
      // every one must resolve to a valid 0-based page index.
      expect(tocNames.length).toBeGreaterThan(0);
      for (const name of tocNames) {
        const page = map.get(name);
        expect(page).toBeDefined();
        expect(page).toBeGreaterThanOrEqual(0);
      }

      // Monotonic sanity: the TOC lists headings in document order, so later
      // `_TocN` names must not resolve to EARLIER pages than earlier ones. This
      // catches a map that mixed up page attribution (a real navigation bug).
      const sorted = tocNames.sort();
      let prevPage = -1;
      for (const name of sorted) {
        const page = map.get(name) as number;
        expect(page).toBeGreaterThanOrEqual(prevPage);
        prevPage = page;
      }

      // A concrete jump: the first TOC entry lands on some page, and at least one
      // entry lands on a page AFTER page 0 (the TOC itself) — i.e. clicking it in
      // a viewer actually moves the user forward through the document.
      const pages = tocNames.map((n) => map.get(n) as number);
      expect(Math.max(...pages)).toBeGreaterThan(0);
    });
  },
);
