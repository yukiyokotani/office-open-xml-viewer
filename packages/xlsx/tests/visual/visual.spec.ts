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

// ── Fidelity targets ──────────────────────────────────────────────────────────
// Tracked demo workbooks with committed reference images (references/{name}/)
// and fidelity scores (references/{name}/scores.json). Each entry needs:
//   name       : public path stem (loads /{name}.xlsx, reads references/{name}/)
//   sheetCount : exact sheet count; it must equal the renderer's sheet count
// Regression (self-VRT) coverage is not listed here: `pnpm vrt` renders every
// workbook in public/demo/ and `pnpm vrt:private` every workbook in
// public/private/xlsx/ against previous-renderer baselines.
const XLSX_FILES: { name: string; sheetCount: number }[] = [
  { name: 'demo/sample-1', sheetCount: 5 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT = 20;
// Fidelity-score ratchet: fail if a page's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the Excel ground truth even while staying
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

test.describe('xlsx visual fidelity', () => {
  for (const { name, sheetCount } of SELF_VRT ? [] : XLSX_FILES) {
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

        const actualSheetCount = Number(await page.evaluate(() => document.body.dataset.sheetCount));
        if (actualSheetCount !== sheetCount) {
          throw new Error(
            `${name}: the renderer reports ${actualSheetCount} sheet(s) but XLSX_FILES declares ` +
            `${sheetCount}; keep the declared count exact so no sheet goes unchecked.`
          );
        }

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
        const refPath = `tests/visual/references/${name}/sheet-${sheetNum}.png`;
        if (!existsSync(refPath)) {
          throw new Error(`missing fidelity reference: ${refPath}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
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

        if (diffPct > FAIL_ABOVE_PCT) {
          throw new Error(
            `${name} sheet ${sheetNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${FAIL_ABOVE_PCT}%`
          );
        }

        // Fidelity-score ratchet. UPDATE_SCORES rewrites the stored score;
        // otherwise a committed score is a floor.
        const key = `sheet-${sheetNum}`;
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
type SelfVrtCorpus = 'demo' | 'corpus' | 'private' | 'private-fast';

// Fast-tier selection (tests/visual/corpus-features.mjs): per file, the
// sheets that add structural XML features, out of the file's sheet count.
type FastTierSelection = {
  digest: string;
  files: Map<string, { unitCount: number; units: number[] }>;
};

// Serve the renderer the exact bytes whose SHA-256 the corpus manifest records
// (selfVrtInputBody), not whatever the dev server finds on disk later.
async function serveSelfVrtInput(page: Page, corpus: SelfVrtCorpus, file: string): Promise<void> {
  const inputPath = `/${selfVrtInputPath({ corpus, file })}`;
  await page.route(
    (url) => decodeURIComponent(url.pathname) === inputPath,
    (route) => route.fulfill({ status: 200, body: selfVrtInputBody({ corpus, file }) }),
  );
}

// `query` holds extra fixture parameters (the public specimens load their
// Roboto / Noto Sans JP faces offline with ?fonts=corpus). `selection` limits
// a `private-fast` run to the selected sheets; a file without any is skipped.
function describeSelfRegression(
  title: string,
  corpus: SelfVrtCorpus,
  files: string[],
  { query = '', selection }: { query?: string; selection?: FastTierSelection } = {},
): void {
  test.describe(title, () => {
    if ((corpus === 'demo' || corpus === 'corpus') && files.length > 0) {
      test.beforeAll(() => {
        prepareSelfVrtCorpus({ corpus, format: 'xlsx', files, snapshot: SNAPSHOT });
      });
    }
    for (const file of files) {
      const selected = selection?.files.get(file);
      if (selection && !selected) throw new Error(`${file} is missing from the fast-tier selection`);
      if (selected && selected.units.length === 0) continue;
      test(file, async ({ page }) => {
        test.setTimeout(600_000);
        const stem = file.slice(0, -'.xlsx'.length);
        if (!SNAPSHOT) clearSelfVrtCandidateOutput({ corpus, stem, itemKind: 'sheet' });
        await serveSelfVrtInput(page, corpus, file);
        const openSheet = async (sheetIndex: number) => {
          await page.goto(
            `/tests/visual/fixture.html?file=${encodeURIComponent(selfVrtInputPath({ corpus, file }))}`
            + `&sheet=${sheetIndex}${query}`,
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

        const firstSheet = selected?.units[0] ?? 0;
        await openSheet(firstSheet);
        const sheetCount = Number(await page.evaluate(() => document.body.dataset.sheetCount));
        expect(sheetCount, `${stem} must report its complete sheet count`).toBeGreaterThan(0);
        if (selected && selected.unitCount !== sheetCount) {
          throw new Error(
            `${stem}: renderer reports ${sheetCount} sheets but the fast-tier selection `
            + `expects ${selected.unitCount}; regenerate it with pnpm vrt:private:fast:select`,
          );
        }
        const sheetIndices = selected?.units ?? Array.from({ length: sheetCount }, (_, index) => index);
        const differences: string[] = [];
        for (const sheetIndex of sheetIndices) {
          if (sheetIndex !== firstSheet) {
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
          const difference = captureOrCompareSelfVrtItem({
            corpus, stem, itemKind: 'sheet', itemIndex: sheetIndex, actual, snapshot: SNAPSHOT,
          });
          if (difference) differences.push(difference);
        }
        verifySelfVrtItemManifest({
          corpus, format: 'xlsx', stem, itemKind: 'sheet', itemCount: sheetCount, snapshot: SNAPSHOT,
          ...(selected ? { itemIndices: selected.units } : {}),
        });
        expect(differences, differences.join('\n')).toEqual([]);
      });
    }
  });
}

const XLSX_DEMO_CORPUS = SELF_VRT ? selfVrtCorpusFiles({ corpus: 'demo', format: 'xlsx' }) : [];
// Public specimen corpus (tests/corpus/, generated into public/corpus/): runs
// under `pnpm vrt` with the demo corpus, or alone under `pnpm vrt:corpus`.
const XLSX_PUBLIC_CORPUS = SELF_VRT ? selfVrtCorpusFiles({ corpus: 'corpus', format: 'xlsx' }) : [];
const XLSX_PRIVATE_CORPUS = process.env.VRT_PRIVATE_CORPUS === '1'
  ? selfVrtCorpusFiles({ corpus: 'private', format: 'xlsx' })
  : [];

// VRT_PRIVATE_TIER=fast renders only the sheets chosen by
// tests/visual/corpus-features.mjs (structural XML feature cover) into the
// separate `private-fast` baseline tree; the default full tier renders every sheet.
const XLSX_PRIVATE_SELF_VRT: SelfVrtCorpus = privateCorpusTier() === 'fast' ? 'private-fast' : 'private';
const XLSX_FAST_SELECTION: FastTierSelection | undefined =
  process.env.VRT_PRIVATE_CORPUS === '1' && XLSX_PRIVATE_SELF_VRT === 'private-fast'
    ? loadFastTierSelection({ format: 'xlsx', files: XLSX_PRIVATE_CORPUS })
    : undefined;

if (process.env.VRT_PRIVATE_CORPUS === '1') {
  prepareSelfVrtCorpus({
    corpus: XLSX_PRIVATE_SELF_VRT, format: 'xlsx', files: XLSX_PRIVATE_CORPUS, snapshot: SNAPSHOT,
    selectionDigest: XLSX_FAST_SELECTION?.digest,
  });
}

describeSelfRegression('demo corpus self regression', 'demo', XLSX_DEMO_CORPUS);
describeSelfRegression('public corpus self regression', 'corpus', XLSX_PUBLIC_CORPUS, {
  query: '&rows=40&cols=10&width=1280&height=960&fonts=corpus',
});
describeSelfRegression('private corpus self regression', XLSX_PRIVATE_SELF_VRT, XLSX_PRIVATE_CORPUS, {
  selection: XLSX_FAST_SELECTION,
});
