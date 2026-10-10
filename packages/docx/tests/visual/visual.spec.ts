import { test, expect } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import {
  captureOrCompareSelfVrtItem,
  clearSelfVrtCandidateOutput,
  prepareSelfVrtCorpus,
  selfVrtCorpusFiles,
  selfVrtInputPath,
  verifySelfVrtItemManifest,
} from '../../../../tests/visual/private-corpus.mjs';

// ── Fidelity targets ──────────────────────────────────────────────────────────
// Tracked demo documents with committed reference images (references/{name}/)
// and fidelity scores (references/{name}/scores.json). Each entry needs:
//   name      : public path stem (loads /{name}.docx, reads references/{name}/)
//   pageCount : exact page count; it must equal the renderer's page count
//   width     : render width in CSS px (the reference image width)
// Regression (self-VRT) coverage is not listed here: `pnpm vrt` renders every
// document in public/demo/ and `pnpm vrt:private` every document in
// public/private/docx/ against previous-renderer baselines.
const DOCX_FILES: { name: string; pageCount: number; width: number }[] = [
  { name: 'demo/sample-1', pageCount: 6, width: 595 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT = 20;
// Fidelity-score ratchet: fail if a page's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the Word ground truth even while staying
// under the coarse FAIL_ABOVE_PCT ceiling.
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
// listed files with their references. The two never share an oracle.
const SELF_VRT = RUN_MODE === 'regression' || SNAPSHOT;

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

test.describe('docx visual fidelity', () => {
  for (const { name, pageCount, width } of SELF_VRT ? [] : DOCX_FILES) {
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

        // Exact page-count guard. The fixture reports the renderer's REAL page
        // count via dataset.pageCount. renderPage() silently clamps an
        // out-of-range index back to page 0 (`pages[pageIndex] ?? pages[0]`), so
        // a stale declared count that exceeds the real pagination would keep
        // snapshotting duplicate first pages under a green status (#993), and
        // one below it would leave pages unchecked. It runs BEFORE the
        // UPDATE_REFS branch on purpose: the silent duplication happened during
        // reference refreshes.
        const actualPageCount = Number(await page.evaluate(() => document.body.dataset.pageCount));
        if (actualPageCount !== pageCount) {
          throw new Error(
            `${name}: the renderer reports ${actualPageCount} page(s) but DOCX_FILES declares ` +
            `${pageCount}; keep the declared count exact so no page is duplicated or unchecked.`
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
        const refPath = `tests/visual/references/${name}/page-${pageNum}.png`;
        if (!existsSync(refPath)) {
          throw new Error(`missing fidelity reference: ${refPath}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
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

        if (diffPct > FAIL_ABOVE_PCT) {
          throw new Error(
            `${name} page ${pageNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${FAIL_ABOVE_PCT}%`
          );
        }

        // Fidelity-score ratchet. UPDATE_SCORES rewrites the stored score;
        // otherwise a committed score is a floor.
        const key = `page-${pageNum}`;
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
// Every file of a corpus is rendered completely and compared pixel-for-pixel
// with the previous renderer's images, bound by manifest to
// VRT_BASELINE_REVISION. `demo` runs under `pnpm vrt`; `private` under
// `pnpm vrt:private` (VRT_PRIVATE_CORPUS=1).
type SelfVrtCorpus = 'demo' | 'private';

function describeSelfRegression(title: string, corpus: SelfVrtCorpus, files: string[]): void {
  test.describe(title, () => {
    if (corpus === 'demo' && files.length > 0) {
      test.beforeAll(() => {
        prepareSelfVrtCorpus({ corpus, format: 'docx', files, snapshot: SNAPSHOT });
      });
    }
    for (const file of files) {
      test(file, async ({ page }) => {
        test.setTimeout(600_000);
        const stem = file.slice(0, -'.docx'.length);
        if (!SNAPSHOT) clearSelfVrtCandidateOutput({ corpus, stem, itemKind: 'page' });
        const openPage = async (pageIndex: number) => {
          await page.goto(
            `/tests/visual/fixture.html?file=${encodeURIComponent(selfVrtInputPath({ corpus, file }))}`
            + `&page=${pageIndex}&width=612`,
          );
          await page.waitForFunction(
            () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
            undefined,
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
          const difference = captureOrCompareSelfVrtItem({
            corpus, stem, itemKind: 'page', itemIndex: pageIndex, actual, snapshot: SNAPSHOT,
          });
          if (difference) differences.push(difference);
        }
        verifySelfVrtItemManifest({
          corpus, format: 'docx', stem, itemKind: 'page', itemCount: pageCount, snapshot: SNAPSHOT,
        });
        expect(differences, differences.join('\n')).toEqual([]);
      });
    }
  });
}

const DOCX_DEMO_CORPUS = SELF_VRT ? selfVrtCorpusFiles({ corpus: 'demo', format: 'docx' }) : [];
const DOCX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? selfVrtCorpusFiles({ corpus: 'private', format: 'docx' })
  : [];

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  prepareSelfVrtCorpus({ corpus: 'private', format: 'docx', files: DOCX_PRIVATE_CORPUS, snapshot: SNAPSHOT });
}

describeSelfRegression('demo corpus self regression', 'demo', DOCX_DEMO_CORPUS);
describeSelfRegression('private corpus self regression', 'private', DOCX_PRIVATE_CORPUS);
