import { test, expect } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, readdirSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import {
  captureOrComparePrivateItem,
  preparePrivateCorpus,
  verifyPrivateItemManifest,
} from '../../../../tests/visual/private-corpus.mjs';

const DOCX_FILES: { name: string; pageCount: number; width: number }[] = [
  { name: 'private/sample-1', pageCount: 1, width: 612 },
  { name: 'private/sample-2', pageCount: 1, width: 595 },
  { name: 'private/sample-3', pageCount: 3, width: 595 },
  { name: 'private/sample-4', pageCount: 1, width: 595 },
  { name: 'private/sample-5', pageCount: 7, width: 595 },
  // Multi-column (§17.6.4) section carrying both wrapTopAndBottom (§20.4.2.20,
  // anchored in the single-column title area) and column-anchored wrapSquare
  // floats — pixel coverage for the float column-scope semantics (#907).
  // Reference is private (gitignored) and generated locally with UPDATE_REFS=1.
  { name: 'private/sample-10', pageCount: 2, width: 595 },
  // CH13 chart coverage: sample-24 p.3 carries a stockChart (hi-lo-close);
  // sample-25 is a pie3DChart. References are private (gitignored) and generated
  // locally with UPDATE_REFS=1 — they are never committed.
  { name: 'private/sample-24', pageCount: 3, width: 595 },
  { name: 'private/sample-25', pageCount: 1, width: 595 },
  // XF9 vertical writing (§17.6.20 tbRl): a landscape vertical-Japanese
  // newspaper. width = physical page width (842pt, A4 landscape). Reference is
  // private (gitignored) and generated locally with UPDATE_REFS=1.
  // 2 pages per the Word PDF (page 2 carries only the spill of the final
  // column): the deterministic untabled EA docGrid cell height (1.3 em — the
  // sample-9 regression fix) restores the Word page count from the crammed
  // 1-page state the 2026-07-13a baseline captured.
  { name: 'private/sample-26', pageCount: 2, width: 842 },
  // Multilingual / section coverage (references private + gitignored, generated
  // locally with UPDATE_REFS=1):
  // sample-27 = continuous section-break page-number restart fixture (US Letter, 612pt).
  { name: 'private/sample-27', pageCount: 2, width: 612 },
  // sample-28 = Arabic RTL long-form (A4).
  { name: 'private/sample-28', pageCount: 23, width: 595 },
  // sample-29 = Thai script (A4). 11 pages since the #989 baseline-grazing
  // page fit (adjudicated best-fidelity state in the #981 follow-up); a stale
  // higher count silently clamps out-of-range page requests back to page 0,
  // snapshotting duplicate first pages.
  { name: 'private/sample-29', pageCount: 11, width: 595 },
  // sample-30 = Korean script (A4).
  { name: 'private/sample-30', pageCount: 4, width: 595 },
  // sample-31 = Russian / Cyrillic (A4).
  { name: 'private/sample-31', pageCount: 12, width: 595 },
  { name: 'demo/sample-1', pageCount: 6, width: 595 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT = 20;
const REGRESSION_PCT = 0.5;
// Fidelity-score ratchet: fail if a page's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the Word ground truth even while staying
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
// (e.g. "page-3") → match-% (2 dp). Read-modify-write is safe because the VRT
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

test.describe('docx visual regression', () => {
  for (const { name, pageCount, width } of DOCX_FILES) {
    for (let i = 0; i < pageCount; i++) {
      const pageNum = i + 1;

      test(`${name} › page ${pageNum}`, async ({ page }) => {
        await page.goto(
          `/tests/visual/fixture.html?file=${name}.docx&page=${i}&width=${width}`
        );

        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} page ${pageNum}: ${msg}`);
        }

        // Loud out-of-range page guard. `i` is the requested page index and
        // `pageCount` above is this VRT's DECLARED count; the fixture reports the
        // renderer's REAL page count via dataset.pageCount. renderPage() silently
        // clamps an out-of-range index back to page 0 (renderer.ts:
        // `pages[pageIndex] ?? pages[0]`), so a stale declared count that exceeds
        // the real pagination keeps snapshotting duplicate first pages under a
        // green status — the #993 regression, where a private sample dropped
        // 14→11 pages yet ~15 PRs of VRT stayed green on triplicated page-0 refs.
        // Fail the moment a requested index is out of range. This runs BEFORE the
        // UPDATE_REFS branch below on purpose: the silent duplication happened
        // during reference refreshes, so the guard must fire there too.
        const actualPageCount = Number(await page.evaluate(() => document.body.dataset.pageCount));
        if (!Number.isInteger(actualPageCount) || actualPageCount <= 0) {
          throw new Error(
            `${name}: fixture did not report a valid page count (got "${actualPageCount}")`
          );
        }
        if (i >= actualPageCount) {
          throw new Error(
            `${name}: requested page index ${i} (page ${pageNum}) is out of range — the renderer ` +
            `produced only ${actualPageCount} page(s). renderPage() would clamp it back to page 0 ` +
            `and snapshot a duplicate first page. Lower the DOCX_FILES pageCount for ${name} to ` +
            `${actualPageCount}.`
          );
        }

        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} page ${pageNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/page-${pageNum}.png`, actualBuf);

        if (UPDATE_REFS) {
          mkdirSync(`tests/visual/references/${name}`, { recursive: true });
          writeFileSync(`tests/visual/references/${name}/page-${pageNum}.png`, actualBuf);
          console.log(`  ${name} page ${pageNum}: reference updated`);
          return;
        }
        if (SNAPSHOT) {
          mkdirSync(`tests/visual/baseline/${name}`, { recursive: true });
          writeFileSync(`tests/visual/baseline/${name}/page-${pageNum}.png`, actualBuf);
          console.log(`  ${name} page ${pageNum}: baseline captured`);
          return;
        }

        const targetRoot = RUN_MODE === 'regression' ? 'baseline' : 'references';
        const refPath = `tests/visual/${targetRoot}/${name}/page-${pageNum}.png`;
        if (!existsSync(refPath)) {
          if (RUN_MODE === 'regression') {
            throw new Error(`missing regression baseline: ${refPath}`);
          }
          test.skip(true, `no ${targetRoot} image for ${name} page ${pageNum}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          if (RUN_MODE === 'regression') {
            throw new Error(
              `${name} page ${pageNum}: regression dimensions changed from ` +
              `${refW}×${refH} to ${actualPng.width}×${actualPng.height}`,
            );
          }
          console.warn(
            `  ${name} page ${pageNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ` +
            `ref=${refW}×${refH}`
          );
        }

        const w = Math.max(actualPng.width, refW);
        const h = Math.max(actualPng.height, refH);

        // Pad both images to same size so pixelmatch doesn't throw
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
        writeFileSync(`tests/visual/diffs/${name}/page-${pageNum}.png`, PNG.sync.write(diff));

        const totalPx = w * h;
        const diffPct = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        console.log(
          `  ${name} page ${pageNum}: ` +
          `match=${matchPct.toFixed(1)}%  diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        const limit = RUN_MODE === 'regression' ? REGRESSION_PCT : FAIL_ABOVE_PCT;
        if (diffPct > limit) {
          throw new Error(
            `${name} page ${pageNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${limit}% in ${RUN_MODE} mode`
          );
        }

        // Fidelity-score ratchet (fidelity mode only; the regression mode above
        // already gates against the captured baseline). UPDATE_SCORES rewrites
        // the stored score; otherwise a committed score is a floor.
        if (RUN_MODE === 'fidelity') {
          const key = `page-${pageNum}`;
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

const DOCX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? readdirSync('public/private/docx')
      .filter((file) => file.endsWith('.docx') && !file.startsWith('~$'))
      .map((file) => `docx/${file}`)
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
  : [];

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  preparePrivateCorpus({ format: 'docx', files: DOCX_PRIVATE_CORPUS, snapshot: SNAPSHOT });
}

test.describe('private corpus self regression', () => {
  for (const file of DOCX_PRIVATE_CORPUS) {
    test(file, async ({ page }) => {
      test.setTimeout(600_000);
      const stem = file.slice(0, -'.docx'.length);
      const openPage = async (pageIndex: number) => {
        await page.goto(
          `/tests/visual/fixture.html?file=${encodeURIComponent(`private/${file}`)}`
          + `&page=${pageIndex}&width=612`,
        );
        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 120_000 },
        );
        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const message = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`${stem} page ${pageIndex + 1}: ${message}`);
        }
      };

      await openPage(0);
      const pageCount = Number(await page.evaluate(() => document.body.dataset.pageCount));
      expect(pageCount, `${stem} must report its complete page count`).toBeGreaterThan(0);
      const differences: string[] = [];
      for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
        if (pageIndex > 0) {
          await page.evaluate(async (index) => {
            const render = (globalThis as unknown as {
              renderDocxVrtPage(pageIndex: number): Promise<void>;
            }).renderDocxVrtPage;
            await render(index);
          }, pageIndex);
        }
        const dataUrl = await page.evaluate(() =>
          (document.querySelector('canvas') as HTMLCanvasElement | null)?.toDataURL('image/png'));
        if (!dataUrl) throw new Error(`${stem} page ${pageIndex + 1}: no canvas`);
        const actual = Buffer.from(dataUrl.split(',')[1], 'base64');
        const difference = captureOrComparePrivateItem({
          stem, itemKind: 'page', itemIndex: pageIndex, actual, snapshot: SNAPSHOT,
        });
        if (difference) differences.push(difference);
      }
      verifyPrivateItemManifest({
        format: 'docx', stem, itemKind: 'page', itemCount: pageCount, snapshot: SNAPSHOT,
      });
      expect(differences, differences.join('\n')).toEqual([]);
    });
  }
});
