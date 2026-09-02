import { test, expect } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, readdirSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import {
  captureOrComparePrivateItem,
  preparePrivateCorpus,
  verifyPrivateItemManifest,
} from '../../../../tests/visual/private-corpus.mjs';

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

// ── Test targets ──────────────────────────────────────────────────────────────
// Add entries here to include additional PPTX files.
// Each entry needs:
//   name       : filename stem (loads /{name}.pptx, reads references/{name}/)
//   slideCount : number of slides to test (must have matching reference images)
const PPTX_FILES: { name: string; slideCount: number }[] = [
  { name: 'private/sample-1', slideCount: 5 },
  { name: 'private/sample-2', slideCount: 17 },
  { name: 'private/sample-3', slideCount: 21 },
  { name: 'private/sample-4', slideCount: 6 },
  { name: 'private/sample-5', slideCount: 16 },
  { name: 'private/sample-6', slideCount: 13 },
  { name: 'private/sample-7', slideCount: 2 },
  { name: 'private/sample-8', slideCount: 1 },
  { name: 'private/sample-9', slideCount: 2 },
  { name: 'private/sample-10', slideCount: 5 },
  { name: 'private/sample-11', slideCount: 3 },
  { name: 'demo/sample-1', slideCount: 9 },
];

// Per-pixel color tolerance for pixelmatch (0 = exact, 1 = fully lenient)
// 0.20 absorbs font hinting / sub-pixel differences between PowerPoint and Canvas
const PIXEL_THRESHOLD = 0.20;

// Set to a number (e.g. 20) to fail the test when diff exceeds that percentage.
// Set to null to always pass (report-only mode).
const FAIL_ABOVE_PCT = 20;
const REGRESSION_PCT = 0.5;
// Fidelity-score ratchet: fail if a slide's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the PowerPoint ground truth even while
// staying under the coarse FAIL_ABOVE_PCT ceiling.
const RATCHET_DROP_PCT = 0.5;

// UPDATE_REFS=1 pnpm vrt → adopt the current canvas output as the new reference.
// Skips diff comparison and writes the screenshot straight into references/.
const UPDATE_REFS = process.env.UPDATE_REFS === '1';
// UPDATE_SCORES=1 pnpm vrt → record the current fidelity match-% into
// references/<name>/scores.json WITHOUT touching the reference PNGs. This is how
// the committed demo scores are (re)generated; it never rewrites ground truth.
const UPDATE_SCORES = process.env.UPDATE_SCORES === '1';
const SNAPSHOT = process.env.VRT_SNAPSHOT === '1';
const RUN_MODE = process.env.VRT_MODE === 'regression' ? 'regression' : 'fidelity';

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

// ── Tests ─────────────────────────────────────────────────────────────────────
test.describe('visual regression', () => {
  for (const { name, slideCount } of PPTX_FILES) {
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

        // Give the browser one extra frame to flush composite layers
        await page.waitForTimeout(200);

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
        if (SNAPSHOT) {
          mkdirSync(`tests/visual/baseline/${name}`, { recursive: true });
          writeFileSync(`tests/visual/baseline/${name}/slide-${slideNum}.png`, actualBuf);
          console.log(`  ${name} slide ${slideNum}: baseline captured`);
          return;
        }

        const targetRoot = RUN_MODE === 'regression' ? 'baseline' : 'references';
        const refPath = `tests/visual/${targetRoot}/${name}/slide-${slideNum}.png`;
        if (!existsSync(refPath)) {
          if (RUN_MODE === 'regression') {
            throw new Error(`missing regression baseline: ${refPath}`);
          }
          test.skip(true, `no ${targetRoot} image for ${name} slide ${slideNum}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          if (RUN_MODE === 'regression') {
            throw new Error(
              `${name} slide ${slideNum}: regression dimensions changed from ` +
              `${refW}×${refH} to ${actualPng.width}×${actualPng.height}`,
            );
          }
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

        // ── Optional hard failure ──────────────────────────────────────────
        const limit = RUN_MODE === 'regression' ? REGRESSION_PCT : FAIL_ABOVE_PCT;
        if (diffPct > limit) {
          throw new Error(
            `${name} slide ${slideNum} pixel diff ${diffPct.toFixed(1)}% exceeds ` +
            `${limit}% in ${RUN_MODE} mode`
          );
        }

        // Fidelity-score ratchet (fidelity mode only; the regression mode above
        // already gates against the captured baseline). UPDATE_SCORES rewrites
        // the stored score; otherwise a committed score is a floor.
        if (RUN_MODE === 'fidelity') {
          const key = `slide-${slideNum}`;
          if (UPDATE_SCORES) {
            writeScore(name, key, matchPct);
          } else {
            const prior = readScores(name)[key];
            if (prior !== undefined && matchPct < prior - RATCHET_DROP_PCT) {
              throw new Error(
                `${name} ${key} fidelity regressed: match ${matchPct.toFixed(2)}% ` +
                `is >${RATCHET_DROP_PCT}pt below the recorded ${prior.toFixed(2)}%`
              );
            }
          }
        }
      });
    }
  }
});

const PPTX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? readdirSync('public/private')
      .filter((file) => file.endsWith('.pptx') && !file.startsWith('~$'))
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
  : [];

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  preparePrivateCorpus({ format: 'pptx', files: PPTX_PRIVATE_CORPUS, snapshot: SNAPSHOT });
}

test.describe('private corpus self regression', () => {
  for (const file of PPTX_PRIVATE_CORPUS) {
    test(file, async ({ page }) => {
      test.setTimeout(600_000);
      const stem = file.slice(0, -'.pptx'.length);
      const openSlide = async (slideIndex: number) => {
        await page.goto(
          `/tests/visual/fixture.html?pptx=${encodeURIComponent(`private/${stem}`)}`
          + `&slide=${slideIndex}`,
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
        await page.waitForTimeout(200);
      };

      await openSlide(0);
      const slideCount = Number(await page.evaluate(() => document.body.dataset.slideCount));
      expect(slideCount, `${stem} must report its complete slide count`).toBeGreaterThan(0);
      const differences: string[] = [];
      for (let slideIndex = 0; slideIndex < slideCount; slideIndex++) {
        if (slideIndex > 0) {
          await page.evaluate(async (index) => {
            const render = (globalThis as unknown as {
              renderPptxVrtSlide(slideIndex: number): Promise<void>;
            }).renderPptxVrtSlide;
            await render(index);
          }, slideIndex);
          await page.waitForTimeout(200);
        }
        const dataUrl = await page.evaluate(() =>
          (document.querySelector('canvas') as HTMLCanvasElement | null)?.toDataURL('image/png'));
        if (!dataUrl) throw new Error(`${stem} slide ${slideIndex + 1}: no canvas`);
        const actual = Buffer.from(dataUrl.split(',')[1], 'base64');
        const difference = captureOrComparePrivateItem({
          stem, itemKind: 'slide', itemIndex: slideIndex, actual, snapshot: SNAPSHOT,
        });
        if (difference) differences.push(difference);
      }
      verifyPrivateItemManifest({
        format: 'pptx', stem, itemKind: 'slide', itemCount: slideCount, snapshot: SNAPSHOT,
      });
      expect(differences, differences.join('\n')).toEqual([]);
    });
  }
});
