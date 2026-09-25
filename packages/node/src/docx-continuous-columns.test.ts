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
  type DocxRendererModule,
} from './test-imports';

// skia-canvas is a devDependency, so `pnpm install` provides it in CI as well as
// locally; the private journal samples are git-ignored (not redistributable), so
// this suite still self-skips where they are absent. Load skia through the shared
// helper: absent → skip cleanly (local), OOXML_REQUIRE_SKIA=1 (CI) → hard failure.
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
// The WASM-backed docx parser + renderer are only loaded when skia is present.
// Both statically import git-ignored WASM glue, so they need `pnpm build:wasm`
// first; under OOXML_REQUIRE_SKIA=1 a failure to load is a hard error.
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia ? await loadDocxRendererForTests() : null;

type DocxLayout = ReturnType<DocxRendererModule['layoutDocument']>;
type RetainedBodyNode = DocxLayout['pages'][number]['layers']['body'][number];

const samplePath = (n: number) =>
  resolve(ROOT, `packages/docx/public/private/docx/sample-${n}.docx`);
const haveSamples =
  existsSync(samplePath(5)) && existsSync(samplePath(12)) && existsSync(samplePath(13));

/** Private sample numbers are local filenames, not stable fixture identities.
 * Only apply a Word-ground-truth assertion when the retained paragraph graph
 * proves that the installed file is the corpus document the assertion names. */
function retainedParagraphTexts(layout: DocxLayout): string[] {
  const texts: string[] = [];
  const seen = new WeakSet<object>();
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    const candidate = value as Readonly<{
      kind?: unknown;
      lines?: readonly Readonly<{ placements?: readonly Readonly<{ kind?: unknown; text?: unknown }>[] }>[];
    }>;
    if (candidate.kind === 'paragraph' && Array.isArray(candidate.lines)) {
      texts.push(candidate.lines.flatMap((line) => line.placements ?? [])
        .flatMap((placement) => placement.kind === 'text' && typeof placement.text === 'string'
          ? [placement.text]
          : [])
        .join(''));
    }
    for (const child of Object.values(value)) {
      if (ArrayBuffer.isView(child) || child instanceof ArrayBuffer) continue;
      if (Array.isArray(child)) child.forEach(visit);
      else visit(child);
    }
  };
  visit(layout.pages);
  return texts;
}

function isExpectedPrivateFixture(layout: DocxLayout, fragments: readonly string[]): boolean {
  const paragraphs = retainedParagraphTexts(layout)
    .map((text) => text.replace(/\s+/g, ' ').trim());
  return fragments.every((fragment) => paragraphs.some((text) => text.includes(fragment)));
}

// ECMA-376 §17.6.4 newspaper columns + §17.18.79 "continuous" section marks.
// Both journal templates flow their body through `continuous` section breaks
// that flip the column count (1 ⇄ 2) mid-page. The paginator must place each
// multi-column region's later columns at the REGION top (where the section
// began on the page), not the page content top — otherwise the second column
// overprints the preceding single-column content and the page absorbs too much
// content (sample-12 collapsed 3 Word pages into 2). Ground truth = the Word
// PDF page counts next to each .docx.
describe.skipIf(!skia || !docxMod || !rendererMod || !haveSamples)(
  'continuous column-count section breaks (sample-12/13)',
  () => {
    let restore: Array<() => void> = [];
    const paginate = async (n: number) => {
      restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
      try {
        const { materializeDocxDocument } = docxMod!;
        const { createLayoutServices, layoutDocument } = rendererMod!;
        const doc = await materializeDocxDocument(readFileSync(samplePath(n)));
        const layoutServices = createLayoutServices(doc);
        return layoutDocument(doc, layoutServices, { currentDateMs: 0 });
      } finally {
        restore.forEach((r) => r());
      }
    };

    // Tier 1 (column-region top tracking): the second column of a continuous
    // mid-page multi-column section starts at the region top, not the page top —
    // so the overprint is gone and sample-12 flows across its 3 Word pages.
    it('sample-12 paginates to 3 pages (Word ground truth)', async (context) => {
      const pages = await paginate(12);
      if (!isExpectedPrivateFixture(pages, ['Figure 1 This is a Sample Figure'])) {
        context.skip('the installed local sample-12 is not the continuous-column corpus');
        return;
      }
      expect(pages.pages.length).toBe(3);
    });

    // Text of a retained paragraph's placements. Enough to locate a caption /
    // heading paragraph by content; tables not needed here.
    const elementText = (el: RetainedBodyNode): string =>
      el.kind === 'paragraph'
        ? el.lines.flatMap((line) => line.placements)
            .flatMap((placement) => placement.kind === 'text' ? [placement.text] : [])
            .join('')
        : '';
    const findParaPage = (layout: DocxLayout, startsWith: string): number =>
      layout.pages.findIndex((page) =>
        page.layers.body.some((el) =>
          elementText(el).replace(/\s+/g, ' ').trim().startsWith(startsWith)),
      );

    // ECMA-376 §17.3.1.29 + §20.4.2.17 (regression from #676): the
    // figure on sample-12 p.2 is a wrapSquare anchor offset ~70pt into the column,
    // so both side gaps are ~61–64pt — BELOW 1 inch. #676 replaced the empty
    // paragraph-mark line-start threshold with the 1-inch CONTENT-line rule, which
    // flowed the figure's nine trailing blank-line marks BELOW the float band and
    // pushed the caption "Figure 1 …" + the 4. CONCLUSION heading onto page 3. An
    // empty paragraph mark stays beside a float whenever the gap holds the pilcrow
    // (paragraphMarkEmPx), dropping below only for a full-width band — so the
    // caption and CONCLUSION belong on page 2 (0-indexed page 1), matching the
    // Word-exported PDF (sample-12.pdf p.2).
    it('sample-12 keeps the figure caption + CONCLUSION on page 2 (#676 regression)', async (context) => {
      const pages = await paginate(12);
      if (!isExpectedPrivateFixture(pages, [
        'Figure 1 This is a Sample Figure',
        'The other first headings can be Research',
      ])) {
        context.skip('the installed local sample-12 is not the continuous-column corpus');
        return;
      }
      expect(findParaPage(pages, 'Figure 1 This is a Sample Figure')).toBe(1);
      expect(findParaPage(pages, 'The other first headings can be Research')).toBe(1);
    });

    // 5 pages, matching Word. The intro 2-col section opens with a "continuous"
    // section break, so it stays on the title page (§17.6.22: the break is
    // governed by the upcoming section's start type, not the title section's
    // nextPage). Restored after the sample-5 cover overprint was fixed at its
    // real root — a PageBreak after the "Cover Pages" building block (§17.5.2) —
    // instead of forcing every nextPage→continuous boundary to break a page.
    it('sample-13 paginates to 5 pages (Word ground truth)', async (context) => {
      const pages = await paginate(13);
      if (!isExpectedPrivateFixture(pages, ['Journal homepage:'])) {
        context.skip('the installed local sample-13 is not the journal corpus');
        return;
      }
      expect(pages.pages.length).toBe(5);
    });

    // sample-5 (夢十夜): the cover is a "Cover Pages" building block (§17.5.2)
    // whose text flow is empty — the page is filled by page-anchored cover
    // graphics. Word places it on its own page and starts the novel body on
    // page 2, even though the body section opens with a "continuous" break. The
    // parser emits a PageBreak after the cover content so the cover stands alone:
    // 7 pages. Were the cover detection to fail, the continuous body would flow
    // up onto page 1 and the document would collapse to 6 pages.
    it('sample-5 cover page stands alone — 7 pages (Word ground truth)', async (context) => {
      const pages = await paginate(5);
      if (!isExpectedPrivateFixture(pages, ['夢十夜'])) {
        context.skip('the installed local sample-5 is not the cover-page corpus');
        return;
      }
      expect(pages.pages.length).toBe(7);
    });
  },
);
