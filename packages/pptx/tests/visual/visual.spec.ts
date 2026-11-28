import { test, expect, type Page } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import {
  captureOrCompareSelfVrtItem,
  clearSelfVrtCandidateOutput,
  prepareSelfVrtCorpus,
  selfVrtCorpusFiles,
  selfVrtInputBody,
  selfVrtInputPath,
  verifySelfVrtItemManifest,
} from '../../../../tests/visual/private-corpus.mjs';
import { loadFastTierSelection, privateCorpusTier } from '../../../../tests/visual/corpus-features.mjs';

test.afterEach(async ({ page }) => {
  await page.evaluate(() => {
    const destroy = (globalThis as unknown as {
      destroyPptxVrt?: () => void;
    }).destroyPptxVrt;
    destroy?.();
  }).catch(() => {
    // Navigation or a load failure may already have disposed the page.
  });
});

// ── Fidelity targets ──────────────────────────────────────────────────────────
// Tracked demo decks with committed reference images (references/{name}/) and
// fidelity scores (references/{name}/scores.json). Each entry needs:
//   name       : public path stem (loads /{name}.pptx, reads references/{name}/)
//   slideCount : exact slide count; it must equal the renderer's slide count
// Regression (self-VRT) coverage is not listed here: `pnpm vrt` renders every
// deck in public/demo/ and `pnpm vrt:private` every deck in
// public/private/pptx/ against previous-renderer baselines.
const PPTX_FILES: { name: string; slideCount: number }[] = [
  { name: 'demo/sample-1', slideCount: 9 },
];

// Per-pixel color tolerance for pixelmatch (0 = exact, 1 = fully lenient)
// 0.20 absorbs font hinting / sub-pixel differences between PowerPoint and Canvas
const PIXEL_THRESHOLD = 0.20;

// Fail the fidelity test when the diff against the reference exceeds this percentage.
const FAIL_ABOVE_PCT = 20;
// Fidelity-score ratchet: fail if a slide's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the PowerPoint ground truth even while
// staying under the coarse FAIL_ABOVE_PCT ceiling.
const RATCHET_DROP_PCT = 0.5;

// UPDATE_REFS=1 pnpm vrt:fidelity → adopt the current canvas output as the new
// reference. Only with explicit user approval (see AGENTS.md).
const UPDATE_REFS = process.env.UPDATE_REFS === '1';
// UPDATE_SCORES=1 pnpm vrt:fidelity → record the current fidelity match-% into
// references/<name>/scores.json WITHOUT touching the reference PNGs. This is how
// the committed demo scores are (re)generated from a clean latest-main checkout
// (AGENTS.md); it never rewrites ground truth.
const UPDATE_SCORES = process.env.UPDATE_SCORES === '1';
const SNAPSHOT = process.env.VRT_SNAPSHOT === '1';
const RUN_MODE = process.env.VRT_MODE === 'regression' ? 'regression' : 'fidelity';
// `vrt` / `vrt:snapshot` run the self-VRT corpora; `vrt:fidelity` compares the
// listed decks with their references. The two never share an oracle.
const SELF_VRT = RUN_MODE === 'regression' || SNAPSHOT;

// Per-sample fidelity scores live next to the reference PNGs
// (references/<name>/scores.json), so they inherit the exact same commit policy:
// demo scores are tracked, private scores are gitignored. Keyed by item id
// (e.g. "slide-3") → match-% (2 dp). Read-modify-write is safe because the VRT
// config runs sequentially (fullyParallel: false).
function scoresPathFor(name: string): string {
  return `tests/visual/references/${name}/scores.json`;
}
function readScores(name: string): Record<string, number> {
  const p = scoresPathFor(name);
  if (!existsSync(p)) return {};
  try {
    return JSON.parse(readFileSync(p, 'utf8')) as Record<string, number>;
  } catch {
    return {};
  }
}
function writeScore(name: string, key: string, matchPct: number): void {
  const scores = readScores(name);
  scores[key] = Math.round(matchPct * 100) / 100;
  mkdirSync(`tests/visual/references/${name}`, { recursive: true });
  const ordered = Object.fromEntries(Object.entries(scores).sort(([a], [b]) => a.localeCompare(b)));
  writeFileSync(scoresPathFor(name), JSON.stringify(ordered, null, 2) + '\n');
}

// ── Fidelity tests ────────────────────────────────────────────────────────────
test.describe('visual fidelity', () => {
  for (const { name, slideCount } of SELF_VRT ? [] : PPTX_FILES) {
    for (let i = 0; i < slideCount; i++) {
      const slideNum = i + 1;

      test(`${name} › slide ${slideNum}`, async ({ page }) => {
        // ── Load the fixture and wait for rendering to complete ────────────
        await page.goto(`/tests/visual/fixture.html?pptx=${name}&slide=${i}`);

        await page.waitForFunction(
          () =>
            document.body.dataset.status === 'ready' ||
            document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} slide ${slideNum}: ${msg}`);
        }

        const actualSlideCount = Number(await page.evaluate(() => document.body.dataset.slideCount));
        if (actualSlideCount !== slideCount) {
          throw new Error(
            `${name}: the renderer reports ${actualSlideCount} slide(s) but PPTX_FILES declares ` +
            `${slideCount}; keep the declared count exact so no slide goes unchecked.`
          );
        }

        // ── Capture the canvas via toDataURL ──────────────────────────────
        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} slide ${slideNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/slide-${slideNum}.png`, actualBuf);

        // ── UPDATE_REFS mode: adopt the current canvas output as the new reference ─
        if (UPDATE_REFS) {
          mkdirSync(`tests/visual/references/${name}`, { recursive: true });
          writeFileSync(`tests/visual/references/${name}/slide-${slideNum}.png`, actualBuf);
          console.log(`  ${name} slide ${slideNum}: reference updated`);
          return;
        }

        const refPath = `tests/visual/references/${name}/slide-${slideNum}.png`;
        if (!existsSync(refPath)) {
          throw new Error(`missing fidelity reference: ${refPath}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          console.error(
            `  ${name} slide ${slideNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ` +
            `ref=${refW}×${refH}`
          );
        }

        const w = Math.max(actualPng.width, refW);
        const h = Math.max(actualPng.height, refH);

        const pad = (png: ReturnType<typeof PNG.sync.read>, tw: number, th: number) => {
          if (png.width === tw && png.height === th) return png;
          const out = new PNG({ width: tw, height: th });
          out.data.fill(255);
          PNG.bitblt(png, out, 0, 0, png.width, png.height, 0, 0);
          return out;
        };
        const refPadded = pad(refPng, w, h);
        const actualPadded = pad(actualPng, w, h);

        // ── Pixel comparison ───────────────────────────────────────────────
        const diff = new PNG({ width: w, height: h });
        const diffPixels = pixelmatch(
          refPadded.data,
          actualPadded.data,
          diff.data,
          w, h,
          { threshold: PIXEL_THRESHOLD, includeAA: true }
        );
        mkdirSync(`tests/visual/diffs/${name}`, { recursive: true });
        writeFileSync(`tests/visual/diffs/${name}/slide-${slideNum}.png`, PNG.sync.write(diff));

        const totalPx  = w * h;
        const diffPct  = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        // ── Report ─────────────────────────────────────────────────────────
        console.log(
          `  ${name} slide ${slideNum}: ` +
          `match=${matchPct.toFixed(1)}%  ` +
          `diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        if (diffPct > FAIL_ABOVE_PCT) {
          throw new Error(
            `${name} slide ${slideNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${FAIL_ABOVE_PCT}%`
          );
        }

        // Fidelity-score ratchet. UPDATE_SCORES rewrites the stored score;
        // otherwise a committed score is a floor.
        const key = `slide-${slideNum}`;
        if (UPDATE_SCORES) {
          writeScore(name, key, matchPct);
        } else {
          const prior = readScores(name)[key];
          if (prior === undefined) {
            throw new Error(`${name} ${key} has no recorded fidelity score in ${scoresPathFor(name)}`);
          }
          if (matchPct < prior - RATCHET_DROP_PCT) {
            throw new Error(
              `${name} ${key} fidelity regressed: match ${matchPct.toFixed(2)}% ` +
              `is >${RATCHET_DROP_PCT}pt below the recorded ${prior.toFixed(2)}%`
            );
          }
        }
      });
    }
  }
});

// ── Self-VRT (previous-renderer regression) ───────────────────────────────────
// Every deck of a corpus is rendered completely and compared pixel-for-pixel
// with the previous renderer's images, bound by manifest to
// VRT_BASELINE_REVISION. `demo` runs under `pnpm vrt`; `private` under
// `pnpm vrt:private` (VRT_PRIVATE_CORPUS=1).
type SelfVrtCorpus = 'demo' | 'corpus' | 'private' | 'private-fast';

// Fast-tier selection (tests/visual/corpus-features.mjs): per file, the
// slides that add structural XML features, out of the file's slide count.
type FastTierSelection = {
  digest: string;
  files: Map<string, { unitCount: number; units: number[] }>;
};

// Serve the renderer the exact bytes whose SHA-256 the corpus manifest records
// (selfVrtInputBody), not whatever the dev server finds on disk later.
// The route matches the raw encoded path the fixture requests (each segment
// through encodeURIComponent), so a name such as `A%20B` cannot be confused
// with a sibling `A B`. The returned function reports how many times the route
// served the input; a test asserts it after the document loads.
async function serveSelfVrtInput(page: Page, corpus: SelfVrtCorpus | 'demo-scroll', file: string): Promise<() => number> {
  const inputPath = '/' + selfVrtInputPath({ corpus, file }).split('/').map(encodeURIComponent).join('/');
  let served = 0;
  await page.route(
    (url) => url.pathname === inputPath,
    (route) => {
      served += 1;
      return route.fulfill({ status: 200, body: selfVrtInputBody({ corpus, file }) });
    },
  );
  return () => served;
}

// `query` holds extra fixture parameters (the public specimens load their
// Roboto / Noto Sans JP faces offline with ?fonts=corpus). `selection` limits
// a `private-fast` run to the selected slides; a file without any is skipped.
function describeSelfRegression(
  title: string,
  corpus: SelfVrtCorpus,
  files: string[],
  { query = '', selection }: { query?: string; selection?: FastTierSelection } = {},
): void {
  test.describe(title, () => {
    if ((corpus === 'demo' || corpus === 'corpus') && files.length > 0) {
      test.beforeAll(() => {
        prepareSelfVrtCorpus({ corpus, format: 'pptx', files, snapshot: SNAPSHOT });
      });
    }
    for (const file of files) {
      const selected = selection?.files.get(file);
      if (selection && !selected) throw new Error(`${file} is missing from the fast-tier selection`);
      if (selected && selected.units.length === 0) continue;
      test(file, async ({ page }) => {
        test.setTimeout(600_000);
        const stem = file.slice(0, -'.pptx'.length);
        if (!SNAPSHOT) clearSelfVrtCandidateOutput({ corpus, stem, itemKind: 'slide' });
        const inputServed = await serveSelfVrtInput(page, corpus, file);
        const openSlide = async (slideIndex: number) => {
          await page.goto(
            `/tests/visual/fixture.html?pptx=${encodeURIComponent(selfVrtInputPath({ corpus, file: stem }))}`
            + `&slide=${slideIndex}${query}`,
          );
          await page.waitForFunction(
            () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
            { timeout: 120_000 },
          );
          const status = await page.evaluate(() => document.body.dataset.status);
          if (status === 'error') {
            const message = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
            throw new Error(`${stem} slide ${slideIndex + 1}: ${message}`);
          }
        };

        const firstSlide = selected?.units[0] ?? 0;
        await openSlide(firstSlide);
        expect(inputServed(), `${stem}: the verified input route never served the document`).toBeGreaterThan(0);
        const slideCount = Number(await page.evaluate(() => document.body.dataset.slideCount));
        expect(slideCount, `${stem} must report its complete slide count`).toBeGreaterThan(0);
        if (selected && selected.unitCount !== slideCount) {
          throw new Error(
            `${stem}: renderer reports ${slideCount} slides but the fast-tier selection `
            + `expects ${selected.unitCount}; regenerate it with pnpm vrt:private:fast:select`,
          );
        }
        const slideIndices = selected?.units ?? Array.from({ length: slideCount }, (_, index) => index);
        const differences: string[] = [];
        for (const slideIndex of slideIndices) {
          if (slideIndex !== firstSlide) {
            await page.evaluate(async (index) => {
              const render = (globalThis as unknown as {
                renderPptxVrtSlide(slideIndex: number): Promise<void>;
              }).renderPptxVrtSlide;
              await render(index);
            }, slideIndex);
          }
          const dataUrl = await page.evaluate(() =>
            (document.querySelector('canvas') as HTMLCanvasElement | null)?.toDataURL('image/png'));
          if (!dataUrl) throw new Error(`${stem} slide ${slideIndex + 1}: no canvas`);
          const actual = Buffer.from(dataUrl.split(',')[1], 'base64');
          const difference = captureOrCompareSelfVrtItem({
            corpus, stem, itemKind: 'slide', itemIndex: slideIndex, actual, snapshot: SNAPSHOT,
          });
          if (difference) differences.push(difference);
        }
        verifySelfVrtItemManifest({
          corpus, format: 'pptx', stem, itemKind: 'slide', itemCount: slideCount, snapshot: SNAPSHOT,
          ...(selected ? { itemIndices: selected.units } : {}),
        });
        expect(differences, differences.join('\n')).toEqual([]);
      });
    }
  });
}

const PPTX_DEMO_CORPUS = SELF_VRT ? selfVrtCorpusFiles({ corpus: 'demo', format: 'pptx' }) : [];
// Public specimen corpus (tests/corpus/, generated into public/corpus/): runs
// under `pnpm vrt` with the demo corpus, or alone under `pnpm vrt:corpus`.
const PPTX_PUBLIC_CORPUS = SELF_VRT ? selfVrtCorpusFiles({ corpus: 'corpus', format: 'pptx' }) : [];
const PPTX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? selfVrtCorpusFiles({ corpus: 'private', format: 'pptx' })
  : [];

// VRT_PRIVATE_TIER=fast renders only the slides chosen by
// tests/visual/corpus-features.mjs (structural XML feature cover) into the
// separate `private-fast` baseline tree; the default full tier renders every slide.
const PPTX_PRIVATE_SELF_VRT: SelfVrtCorpus = privateCorpusTier() === 'fast' ? 'private-fast' : 'private';
const PPTX_FAST_SELECTION: FastTierSelection | undefined =
  process.env.VRT_PRIVATE_CORPUS === '1' && PPTX_PRIVATE_SELF_VRT === 'private-fast'
    ? loadFastTierSelection({ format: 'pptx', files: PPTX_PRIVATE_CORPUS })
    : undefined;

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  prepareSelfVrtCorpus({
    corpus: PPTX_PRIVATE_SELF_VRT, format: 'pptx', files: PPTX_PRIVATE_CORPUS, snapshot: SNAPSHOT,
    selectionDigest: PPTX_FAST_SELECTION?.digest,
  });
}

describeSelfRegression('demo corpus self regression', 'demo', PPTX_DEMO_CORPUS);
describeSelfRegression('public corpus self regression', 'corpus', PPTX_PUBLIC_CORPUS, {
  query: '&fonts=corpus',
});

// The demo decks through PptxScrollViewer in its default options: every slot
// canvas is compared with the previous renderer's, so the scroll path is
// covered by pixels and not only by the single-canvas viewer.
test.describe('demo corpus scroll viewer self regression', () => {
  const corpus = 'demo-scroll';
  if (PPTX_DEMO_CORPUS.length > 0) {
    test.beforeAll(() => {
      prepareSelfVrtCorpus({ corpus, format: 'pptx', files: PPTX_DEMO_CORPUS, snapshot: SNAPSHOT });
    });
  }
  for (const file of PPTX_DEMO_CORPUS) {
    test(file, async ({ page }) => {
      test.setTimeout(600_000);
      const stem = file.slice(0, -'.pptx'.length);
      if (!SNAPSHOT) clearSelfVrtCandidateOutput({ corpus, stem, itemKind: 'slide' });
      const inputServed = await serveSelfVrtInput(page, corpus, file);
      await page.goto(
        `/tests/visual/fixture.html?viewer=scroll&pptx=${encodeURIComponent(selfVrtInputPath({ corpus, file: stem }))}`,
      );
      await page.waitForFunction(
        () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
        { timeout: 120_000 },
      );
      const status = await page.evaluate(() => document.body.dataset.status);
      if (status === 'error') {
        throw new Error(`${stem}: ${await page.evaluate(() => document.body.dataset.errorMessage ?? '')}`);
      }
      expect(inputServed(), `${stem}: the verified input route never served the document`).toBeGreaterThan(0);
      const slideCount = Number(await page.evaluate(() => document.body.dataset.slideCount));
      expect(slideCount, `${stem} must report its complete slide count`).toBeGreaterThan(0);
      const differences: string[] = [];
      for (let slideIndex = 0; slideIndex < slideCount; slideIndex++) {
        const dataUrl = await page.evaluate((index) => (globalThis as unknown as {
          pptxVrtSlideDataUrl(slideIndex: number): string | null;
        }).pptxVrtSlideDataUrl(index), slideIndex);
        if (!dataUrl) throw new Error(`${stem} slide ${slideIndex + 1}: no mounted scroll slot`);
        const actual = Buffer.from(dataUrl.split(',')[1], 'base64');
        const difference = captureOrCompareSelfVrtItem({
          corpus, stem, itemKind: 'slide', itemIndex: slideIndex, actual, snapshot: SNAPSHOT,
        });
        if (difference) differences.push(difference);
      }
      verifySelfVrtItemManifest({
        corpus, format: 'pptx', stem, itemKind: 'slide', itemCount: slideCount, snapshot: SNAPSHOT,
      });
      expect(differences, differences.join('\n')).toEqual([]);
    });
  }
});
describeSelfRegression('private corpus self regression', PPTX_PRIVATE_SELF_VRT, PPTX_PRIVATE_CORPUS, {
  selection: PPTX_FAST_SELECTION,
});
