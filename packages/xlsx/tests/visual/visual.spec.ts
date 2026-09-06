import { test, expect } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, readdirSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import {
  captureOrComparePrivateItem,
  preparePrivateCorpus,
  verifyPrivateItemManifest,
} from '../../../../tests/visual/private-corpus.mjs';

// ── Test targets ──────────────────────────────────────────────────────────────
// Each entry needs:
//   name       : path stem (loads /{name}.xlsx, reads references/{name}/)
//   sheetCount : number of sheets to test (sheet-1.png .. sheet-N.png in references/)
const XLSX_FILES: { name: string; sheetCount: number }[] = [
  { name: 'demo/sample-1', sheetCount: 5 },
  { name: 'private/sample-1', sheetCount: 3 },
  { name: 'private/sample-2', sheetCount: 4 },
  { name: 'private/sample-3', sheetCount: 2 },
  { name: 'private/sample-4', sheetCount: 1 },
  { name: 'private/sample-5', sheetCount: 2 },
  { name: 'private/sample-6', sheetCount: 1 },
  { name: 'private/sample-7', sheetCount: 1 },
  { name: 'private/sample-8', sheetCount: 1 },
  { name: 'private/sample-9', sheetCount: 3 },
  { name: 'private/sample-10', sheetCount: 1 },
  { name: 'private/sample-11', sheetCount: 1 },
  { name: 'private/sample-12', sheetCount: 8 },
  { name: 'private/sample-13', sheetCount: 2 },
  { name: 'private/sample-14', sheetCount: 2 },
  // Keep Plot Area, Legend, Titles, and Labels in the private self-VRT. The
  // fifth sheet is source data without chart-renderer coverage.
  { name: 'private/sample-15', sheetCount: 4 },
  { name: 'private/sample-16', sheetCount: 2 },
  { name: 'private/sample-17', sheetCount: 2 },
  { name: 'private/sample-18', sheetCount: 2 },
  { name: 'private/sample-19', sheetCount: 2 },
  { name: 'private/sample-20', sheetCount: 2 },
  { name: 'private/sample-21', sheetCount: 2 },
  { name: 'private/sample-22', sheetCount: 2 },
  { name: 'private/sample-23', sheetCount: 2 },
  { name: 'private/sample-24', sheetCount: 2 },
  { name: 'private/sample-25', sheetCount: 4 },
  { name: 'private/sample-26', sheetCount: 2 },
  { name: 'private/sample-27', sheetCount: 1 },
  // sample-28: four sheets, each an OMML equation text box (Fourier series,
  // cone volume, circle area). Adds regression coverage for shape-equation
  // rendering, which had none (issue #877). References are self-baseline
  // (renderer output), regenerated locally with UPDATE_REFS — no Excel export.
  { name: 'private/sample-28', sheetCount: 4 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT = 20;
const REGRESSION_PCT = 0.5;
// Fidelity-score ratchet: fail if a page's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the Excel ground truth even while staying
// under the coarse FAIL_ABOVE_PCT ceiling.
const RATCHET_DROP_PCT = 0.5;

// UPDATE_REFS=1 pnpm vrt → adopt the current canvas output as the new reference.
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
// (e.g. "sheet-3") → match-% (2 dp). Read-modify-write is safe because the VRT
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
  // Stable key order keeps the committed JSON diff-friendly.
  const ordered = Object.fromEntries(Object.entries(scores).sort(([a], [b]) => a.localeCompare(b)));
  writeFileSync(scoresPathFor(name), JSON.stringify(ordered, null, 2) + '\n');
}

test.describe('xlsx visual regression', () => {
  for (const { name, sheetCount } of XLSX_FILES) {
    for (let i = 0; i < sheetCount; i++) {
      const sheetNum = i + 1;

      test(`${name} › sheet ${sheetNum}`, async ({ page }) => {
        await page.goto(`/tests/visual/fixture.html?file=${name}.xlsx&sheet=${i}`);

        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} sheet ${sheetNum}: ${msg}`);
        }

        await page.waitForTimeout(200);

        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} sheet ${sheetNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/sheet-${sheetNum}.png`, actualBuf);

        if (UPDATE_REFS) {
          mkdirSync(`tests/visual/references/${name}`, { recursive: true });
          writeFileSync(`tests/visual/references/${name}/sheet-${sheetNum}.png`, actualBuf);
          console.log(`  ${name} sheet ${sheetNum}: reference updated`);
          return;
        }
        if (SNAPSHOT) {
          mkdirSync(`tests/visual/baseline/${name}`, { recursive: true });
          writeFileSync(`tests/visual/baseline/${name}/sheet-${sheetNum}.png`, actualBuf);
          console.log(`  ${name} sheet ${sheetNum}: baseline captured`);
          return;
        }

        const targetRoot = RUN_MODE === 'regression' ? 'baseline' : 'references';
        const refPath = `tests/visual/${targetRoot}/${name}/sheet-${sheetNum}.png`;
        if (!existsSync(refPath)) {
          if (RUN_MODE === 'regression') {
            throw new Error(`missing regression baseline: ${refPath}`);
          }
          test.skip(true, `no ${targetRoot} image for ${name} sheet ${sheetNum}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          if (RUN_MODE === 'regression') {
            throw new Error(
              `${name} sheet ${sheetNum}: regression dimensions changed from ` +
              `${refW}×${refH} to ${actualPng.width}×${actualPng.height}`,
            );
          }
          console.warn(
            `  ${name} sheet ${sheetNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ref=${refW}×${refH}`
          );
        }

        const w = Math.max(actualPng.width, refW);
        const h = Math.max(actualPng.height, refH);

        const pad = (png: ReturnType<typeof PNG.sync.read>, tw: number, th: number) => {
          if (png.width === tw && png.height === th) return png;
          const out = new PNG({ width: tw, height: th });
          out.data.fill(255);
          for (let y = 0; y < Math.min(png.height, th); y++) {
            for (let x = 0; x < Math.min(png.width, tw); x++) {
              const src = (y * png.width + x) * 4;
              const dst = (y * tw + x) * 4;
              out.data[dst]     = png.data[src];
              out.data[dst + 1] = png.data[src + 1];
              out.data[dst + 2] = png.data[src + 2];
              out.data[dst + 3] = png.data[src + 3];
            }
          }
          return out;
        };
        const refPadded    = pad(refPng,    w, h);
        const actualPadded = pad(actualPng, w, h);

        const diff = new PNG({ width: w, height: h });
        const diffPixels = pixelmatch(
          refPadded.data, actualPadded.data, diff.data, w, h,
          { threshold: PIXEL_THRESHOLD, includeAA: true }
        );
        mkdirSync(`tests/visual/diffs/${name}`, { recursive: true });
        writeFileSync(`tests/visual/diffs/${name}/sheet-${sheetNum}.png`, PNG.sync.write(diff));

        const totalPx = w * h;
        const diffPct = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        console.log(
          `  ${name} sheet ${sheetNum}: ` +
          `match=${matchPct.toFixed(1)}%  diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        const limit = RUN_MODE === 'regression' ? REGRESSION_PCT : FAIL_ABOVE_PCT;
        if (diffPct > limit) {
          throw new Error(
            `${name} sheet ${sheetNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${limit}% in ${RUN_MODE} mode`
          );
        }

        // Fidelity-score ratchet (fidelity mode only; the regression mode above
        // already gates against the captured baseline). UPDATE_SCORES rewrites
        // the stored score; otherwise a committed score is a floor.
        if (RUN_MODE === 'fidelity') {
          const key = `sheet-${sheetNum}`;
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

const XLSX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? readdirSync('public/private/xlsx')
      .filter((file) => file.endsWith('.xlsx') && !file.startsWith('~$'))
      .map((file) => `xlsx/${file}`)
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
  : [];

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  preparePrivateCorpus({ format: 'xlsx', files: XLSX_PRIVATE_CORPUS, snapshot: SNAPSHOT });
}

test.describe('private corpus self regression', () => {
  for (const file of XLSX_PRIVATE_CORPUS) {
    test(file, async ({ page }) => {
      test.setTimeout(600_000);
      const stem = file.slice(0, -'.xlsx'.length);
      const openSheet = async (sheetIndex: number) => {
        await page.goto(
          `/tests/visual/fixture.html?file=${encodeURIComponent(`private/${file}`)}`
          + `&sheet=${sheetIndex}`,
        );
        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 120_000 },
        );
        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const message = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`${stem} sheet ${sheetIndex + 1}: ${message}`);
        }
        await page.waitForTimeout(200);
      };

      await openSheet(0);
      const sheetCount = Number(await page.evaluate(() => document.body.dataset.sheetCount));
      expect(sheetCount, `${stem} must report its complete sheet count`).toBeGreaterThan(0);
      const differences: string[] = [];
      for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
        if (sheetIndex > 0) {
          await page.evaluate(async (index) => {
            const render = (globalThis as unknown as {
              renderXlsxVrtSheet(sheetIndex: number): Promise<void>;
            }).renderXlsxVrtSheet;
            await render(index);
          }, sheetIndex);
          await page.waitForTimeout(200);
        }
        const dataUrl = await page.evaluate(() =>
          (document.querySelector('canvas') as HTMLCanvasElement | null)?.toDataURL('image/png'));
        if (!dataUrl) throw new Error(`${stem} sheet ${sheetIndex + 1}: no canvas`);
        const actual = Buffer.from(dataUrl.split(',')[1], 'base64');
        const difference = captureOrComparePrivateItem({
          stem, itemKind: 'sheet', itemIndex: sheetIndex, actual, snapshot: SNAPSHOT,
        });
        if (difference) differences.push(difference);
      }
      verifyPrivateItemManifest({
        format: 'xlsx', stem, itemKind: 'sheet', itemCount: sheetCount, snapshot: SNAPSHOT,
      });
      expect(differences, differences.join('\n')).toEqual([]);
    });
  }
});
