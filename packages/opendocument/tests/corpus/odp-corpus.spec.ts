// Local ODP corpus harness (#1934). For every ODP listed in the run directory
// it loads the file through the `odp` model source in the PPTX renderer,
// records whether it opened, and diffs each shown slide against the matching
// page of the LibreOffice PDF of the same file (LibreOffice exports shown
// slides only). One JSON result per file lands in <run>/results/; run-odp.mjs
// aggregates them.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

interface CorpusFile {
  readonly id: string;
  readonly rel: string;
  readonly odp: string;
  readonly pdf: string;
  readonly hidden: readonly number[];
}

interface BrowserResult {
  readonly status: 'opened' | 'rejected' | 'error';
  readonly message?: string;
  readonly slideCount?: number;
  readonly slides?: readonly { index: number; dataUrl?: string; error?: string }[];
}

const runDir = resolve(process.env.ODF_HARNESS_RUN_DIR!);
const checkout = resolve(process.env.ODF_HARNESS_CHECKOUT!);
const maxSlides = Number(process.env.ODF_HARNESS_MAX_PAGES ?? 0) || Infinity;
const saveMin = Number(process.env.ODF_HARNESS_SAVE_MIN ?? 10);
const files = JSON.parse(readFileSync(join(runDir, 'files.json'), 'utf8')) as CorpusFile[];
const resultsDir = join(runDir, 'results');
const imagesDir = join(runDir, 'images');
mkdirSync(resultsDir, { recursive: true });
mkdirSync(imagesDir, { recursive: true });

const inputsHash = (file: CorpusFile): string => createHash('sha256').update(readFileSync(file.odp))
  .update(existsSync(file.pdf) ? readFileSync(file.pdf) : '').digest('hex');

function pdfPages(pdf: string, prefix: string): PNG[] {
  const limit = Number.isFinite(maxSlides) ? ['-l', String(maxSlides)] : [];
  execFileSync('pdftoppm', ['-png', '-r', '72', ...limit, pdf, prefix], { stdio: ['ignore', 'pipe', 'pipe'] });
  const dir = resolve(prefix, '..');
  const base = prefix.slice(dir.length + 1);
  return readdirSync(dir)
    .filter((name) => name.startsWith(`${base}-`) && name.endsWith('.png'))
    .sort((a, b) => Number(/-(\d+)\.png$/u.exec(a)![1]) - Number(/-(\d+)\.png$/u.exec(b)![1]))
    .map((name) => PNG.sync.read(readFileSync(join(dir, name))));
}

function pdfPageCount(pdf: string): number {
  const info = execFileSync('pdfinfo', [pdf], { encoding: 'utf8' });
  return Number(/^Pages:\s+(\d+)/mu.exec(info)?.[1] ?? 0);
}

function padded(source: PNG, width: number, height: number): PNG {
  if (source.width === width && source.height === height) return source;
  const result = new PNG({ width, height });
  result.data.fill(255);
  PNG.bitblt(source, result, 0, 0, Math.min(source.width, width), Math.min(source.height, height), 0, 0);
  return result;
}

/**
 * Edge map: pixels whose colour differs from a neighbour. Slides are mostly
 * full-bleed colour, so "ink" here means structure (edges of shapes, glyphs
 * and images) rather than non-white area.
 */
function edges(png: PNG): Uint8Array {
  const { width, height, data } = png;
  const out = new Uint8Array(width * height);
  const at = (x: number, y: number) => (y * width + x) * 4;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = at(x, y);
      let edge = false;
      for (const [nx, ny] of [[x + 1, y], [x, y + 1]] as const) {
        if (nx >= width || ny >= height) continue;
        const j = at(nx, ny);
        const d = Math.abs(data[i]! - data[j]!) + Math.abs(data[i + 1]! - data[j + 1]!) + Math.abs(data[i + 2]! - data[j + 2]!);
        if (d > 60) { edge = true; break; }
      }
      out[y * width + x] = edge ? 1 : 0;
    }
  }
  return out;
}

function dilate(mask: Uint8Array, width: number, height: number, radius: number): Uint8Array {
  const horizontal = new Uint8Array(mask.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (!mask[y * width + x]) continue;
      for (let dx = Math.max(0, x - radius); dx <= Math.min(width - 1, x + radius); dx += 1) horizontal[y * width + dx] = 1;
    }
  }
  const out = new Uint8Array(mask.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (!horizontal[y * width + x]) continue;
      for (let dy = Math.max(0, y - radius); dy <= Math.min(height - 1, y + radius); dy += 1) out[dy * width + x] = 1;
    }
  }
  return out;
}

/**
 * Layout score: share of edge pixels (either side) with no edge on the other
 * side within `radius` pixels — missing, extra or displaced structure, not
 * antialiasing or 1–2 px offsets.
 */
function layoutDiff(ref: PNG, act: PNG, radius = 2): number {
  const { width, height } = ref;
  const refInk = edges(ref);
  const actInk = edges(act);
  const refNear = dilate(refInk, width, height, radius);
  const actNear = dilate(actInk, width, height, radius);
  let union = 0;
  let unmatched = 0;
  for (let i = 0; i < refInk.length; i += 1) {
    if (!refInk[i] && !actInk[i]) continue;
    union += 1;
    if ((refInk[i] && !actNear[i]) || (actInk[i] && !refNear[i])) unmatched += 1;
  }
  // A nearly empty slide has too few edges for a ratio; measure against at
  // least 0.5% of the slide so a handful of pixels cannot score 100%.
  return unmatched / Math.max(union, width * height * 0.005) * 100;
}

function sideBySide(reference: PNG, actual: PNG, diff: PNG): PNG {
  const gap = 8;
  const out = new PNG({ width: reference.width * 3 + gap * 2, height: reference.height });
  out.data.fill(128);
  PNG.bitblt(reference, out, 0, 0, reference.width, reference.height, 0, 0);
  PNG.bitblt(actual, out, 0, 0, actual.width, actual.height, reference.width + gap, 0);
  PNG.bitblt(diff, out, 0, 0, diff.width, diff.height, (reference.width + gap) * 2, 0);
  return out;
}

for (const file of files) {
  test(file.rel, async ({ page }) => {
    const resultPath = join(resultsDir, `${file.id}.json`);
    const inputs = inputsHash(file);
    if (existsSync(resultPath)) {
      if (JSON.parse(readFileSync(resultPath, 'utf8')).inputs === inputs) return;
      rmSync(resultPath);
    }
    const started = Date.now();
    const pdfCount = existsSync(file.pdf) ? pdfPageCount(file.pdf) : 0;
    await page.route('**/__odf_harness/input', (route) => route.fulfill({
      status: 200, contentType: 'application/octet-stream', body: readFileSync(file.odp),
    }));
    await page.goto('/tests/visual/non-ooxml-fixture.html');
    const loaded = await page.evaluate(async ({ odpModule }) => {
      const { PptxPresentation } = await import('/src/index.ts');
      const { createOdp } = await import(odpModule);
      try {
        const presentation = await PptxPresentation.load('/__odf_harness/input', { odp: createOdp() });
        (window as unknown as { odfHarness: unknown }).odfHarness = presentation;
        return { status: 'opened', slideCount: presentation.slideCount };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const code = (error as { code?: unknown }).code;
        return { status: code === 'unsupported-content' ? 'rejected' : 'error', message: `${String(code ?? '')} ${message}`.trim() };
      }
    }, { odpModule: `/@fs${checkout}/packages/opendocument/src/odp.ts` });

    let result: BrowserResult = loaded as BrowserResult;
    let references: PNG[] = [];
    const tmp = join(runDir, 'tmp', file.id);
    const hidden = new Set(file.hidden);
    if (loaded.status === 'opened') {
      rmSync(tmp, { recursive: true, force: true });
      mkdirSync(tmp, { recursive: true });
      references = pdfCount > 0 ? pdfPages(file.pdf, join(tmp, 'ref')) : [];
      const shown = Array.from({ length: loaded.slideCount ?? 0 }, (_, i) => i).filter((i) => !hidden.has(i))
        .slice(0, Number.isFinite(maxSlides) ? maxSlides : undefined);
      const width = references[0]?.width ?? 960;
      result = await page.evaluate(async ({ shown, width }) => {
        const presentation = (window as unknown as { odfHarness: {
          slideCount: number;
          renderSlide(canvas: HTMLCanvasElement, index: number, options: { width: number; dpr: number }): Promise<void>;
          destroy(): void;
        } }).odfHarness;
        const slides = [];
        for (const index of shown) {
          const canvas = document.createElement('canvas');
          try {
            await document.fonts.ready;
            await presentation.renderSlide(canvas, index, { width, dpr: 1 });
            await document.fonts.ready;
            await presentation.renderSlide(canvas, index, { width, dpr: 1 });
            slides.push({ index, dataUrl: canvas.toDataURL('image/png') });
          } catch (error) {
            slides.push({ index, error: error instanceof Error ? error.message : String(error) });
          }
        }
        const slideCount = presentation.slideCount;
        presentation.destroy();
        return { status: 'opened', slideCount, slides };
      }, { shown, width }) as BrowserResult;
    }

    const slides = [];
    for (const [position, rendered] of (result.slides ?? []).entries()) {
      const reference = references[position];
      if (!reference || !rendered.dataUrl) {
        slides.push({ slide: rendered.index + 1, error: rendered.error ?? 'no LibreOffice page' });
        continue;
      }
      const actual = PNG.sync.read(Buffer.from(rendered.dataUrl.replace(/^data:image\/png;base64,/u, ''), 'base64'));
      const width = Math.max(reference.width, actual.width);
      const height = Math.max(reference.height, actual.height);
      const ref = padded(reference, width, height);
      const act = padded(actual, width, height);
      const diff = new PNG({ width, height });
      const different = pixelmatch(ref.data, act.data, diff.data, width, height, { threshold: 0.2, includeAA: false });
      const diffPct = different / (width * height) * 100;
      const layoutDiffPct = layoutDiff(ref, act);
      let image: string | undefined;
      if (layoutDiffPct >= saveMin) {
        image = join(imagesDir, `${file.id}-s${rendered.index + 1}.png`);
        writeFileSync(image, PNG.sync.write(sideBySide(ref, act, diff)));
      }
      slides.push({
        slide: rendered.index + 1,
        diffPct: Number(diffPct.toFixed(4)),
        layoutDiffPct: Number(layoutDiffPct.toFixed(4)),
        sizeDeltaPx: [actual.width - reference.width, actual.height - reference.height],
        ...(image ? { image } : {}),
      });
    }
    rmSync(tmp, { recursive: true, force: true });
    const shownCount = result.slideCount === undefined ? undefined : result.slideCount - hidden.size;
    writeFileSync(resultPath, `${JSON.stringify({
      id: file.id,
      rel: file.rel,
      inputs,
      status: result.status,
      ...(result.message ? { message: result.message } : {}),
      pdfPageCount: pdfCount,
      ...(shownCount !== undefined ? { slideCount: result.slideCount, shownCount, slideCountMatch: shownCount === pdfCount } : {}),
      ...(result.slides ? { slides } : {}),
      ms: Date.now() - started,
    }, null, 2)}\n`);
  });
}
