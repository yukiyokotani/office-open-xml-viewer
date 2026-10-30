// Local ODT corpus harness (#1934). For every ODT listed in the run directory
// it loads the file through the `odt` model source in the DOCX renderer,
// records whether it opened, and diffs each rendered page against the
// LibreOffice PDF of the same file. One JSON result per file lands in
// <run>/results/; report.mjs aggregates them. Start it with run.mjs.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

interface CorpusFile {
  readonly id: string;
  readonly rel: string;
  readonly odt: string;
  readonly pdf: string;
}

interface BrowserResult {
  readonly status: 'opened' | 'rejected' | 'error';
  readonly message?: string;
  readonly pageCount?: number;
  readonly pages?: readonly { dataUrl?: string; widthPt: number; heightPt: number; error?: string }[];
}

const runDir = resolve(process.env.ODF_HARNESS_RUN_DIR!);
const checkout = resolve(process.env.ODF_HARNESS_CHECKOUT!);
const wasmPath = process.env.ODF_HARNESS_WASM;
const maxPages = Number(process.env.ODF_HARNESS_MAX_PAGES ?? 20);
/** Pages at or above this layout diff keep a side-by-side image. */
const saveMin = Number(process.env.ODF_HARNESS_SAVE_MIN ?? 5);
const files = JSON.parse(readFileSync(join(runDir, 'files.json'), 'utf8')) as CorpusFile[];
const resultsDir = join(runDir, 'results');
const imagesDir = join(runDir, 'images');
mkdirSync(resultsDir, { recursive: true });
mkdirSync(imagesDir, { recursive: true });

function pdfPages(pdf: string, prefix: string): PNG[] {
  execFileSync('pdftoppm', ['-png', '-r', '72', '-l', String(maxPages), pdf, prefix], { stdio: ['ignore', 'pipe', 'pipe'] });
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

/** Pixels that are not near-white (any channel below 235) after alpha over white. */
function ink(png: PNG, index: number): boolean {
  const a = png.data[index + 3]! / 255;
  for (let c = 0; c < 3; c += 1) {
    if (png.data[index + c]! * a + 255 * (1 - a) < 235) return true;
  }
  return false;
}

/** Ink mask dilated by `radius` pixels (square neighbourhood). */
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
 * Layout score: share of inked pixels (either side) with no ink on the other
 * side within `radius` pixels. Unlike a pixel diff it ignores antialiasing
 * and 1–2 px offsets, so it tracks missing, extra or displaced content.
 */
function layoutDiff(ref: PNG, act: PNG, radius = 2): number {
  const { width, height } = ref;
  const refInk = new Uint8Array(width * height);
  const actInk = new Uint8Array(width * height);
  for (let i = 0; i < refInk.length; i += 1) {
    refInk[i] = ink(ref, i * 4) ? 1 : 0;
    actInk[i] = ink(act, i * 4) ? 1 : 0;
  }
  const refNear = dilate(refInk, width, height, radius);
  const actNear = dilate(actInk, width, height, radius);
  let union = 0;
  let unmatched = 0;
  for (let i = 0; i < refInk.length; i += 1) {
    if (!refInk[i] && !actInk[i]) continue;
    union += 1;
    if ((refInk[i] && !actNear[i]) || (actInk[i] && !refNear[i])) unmatched += 1;
  }
  return union === 0 ? 0 : unmatched / union * 100;
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
    if (existsSync(resultPath)) return;
    const started = Date.now();

    const pdfCount = existsSync(file.pdf) ? pdfPageCount(file.pdf) : 0;

    await page.route('**/__odf_harness/input', (route) => route.fulfill({
      status: 200, contentType: 'application/octet-stream', body: readFileSync(file.odt),
    }));
    if (wasmPath) {
      await page.route('**/__odf_harness/odt.wasm', (route) => route.fulfill({
        status: 200, contentType: 'application/wasm', body: readFileSync(wasmPath),
      }));
    }
    // Any page served by the DOCX dev server gives the module graph its origin;
    // the fixture's own default load is irrelevant here.
    await page.goto('/tests/visual/fixture.html?file=__odf_harness/none');

    const loaded = await page.evaluate(async ({ odtModule, useCustomWasm }) => {
      const { DocxDocument } = await import('/src/index.ts');
      const { createOdt } = await import(odtModule);
      const odt = createOdt(useCustomWasm ? { wasmUrl: new URL('/__odf_harness/odt.wasm', location.href).href } : {});
      try {
        const document = await DocxDocument.load('/__odf_harness/input', { odt });
        (window as unknown as { odfHarnessDocument: unknown }).odfHarnessDocument = document;
        return { status: 'opened', pageCount: document.pageCount };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        // OoxmlError('unsupported-content', { reason: 'source-subset' }): the
        // reader declined the input; the native pass supplies its code.
        const code = (error as { code?: unknown }).code;
        return { status: code === 'unsupported-content' ? 'rejected' : 'error', message: `${String(code ?? '')} ${message}`.trim() };
      }
    }, { odtModule: `/@fs${checkout}/packages/opendocument/src/odt.ts`, useCustomWasm: Boolean(wasmPath) });

    let result: BrowserResult = loaded as BrowserResult;
    let references: PNG[] = [];
    const tmp = join(runDir, 'tmp', file.id);
    if (loaded.status === 'opened') {
      rmSync(tmp, { recursive: true, force: true });
      mkdirSync(tmp, { recursive: true });
      references = pdfCount > 0 ? pdfPages(file.pdf, join(tmp, 'ref')) : [];
      const widths = references.map((png) => png.width);
      result = await page.evaluate(async ({ widths, maxPages }) => {
        const { stableCanvasRender } = await import('/tests/visual/stable-canvas-render.mjs');
        const document = (window as unknown as { odfHarnessDocument: {
          pageCount: number;
          pageSize(index: number): { widthPt: number; heightPt: number };
          renderPage(canvas: HTMLCanvasElement, index: number, options: { width: number; dpr: number }): Promise<void>;
          destroy(): void;
        } }).odfHarnessDocument;
        const pages = [];
        const count = Math.min(document.pageCount, maxPages);
        for (let index = 0; index < count; index += 1) {
          const size = document.pageSize(index);
          const canvas = window.document.createElement('canvas');
          try {
            await stableCanvasRender({
              render: () => document.renderPage(canvas, index, { width: widths[index] ?? Math.round(size.widthPt), dpr: 1 }),
              fonts: window.document.fonts,
              nextFrame: () => new Promise((resolveFrame) => requestAnimationFrame(resolveFrame)),
            });
            pages.push({ dataUrl: canvas.toDataURL('image/png'), widthPt: size.widthPt, heightPt: size.heightPt });
          } catch (error) {
            pages.push({ widthPt: size.widthPt, heightPt: size.heightPt, error: error instanceof Error ? error.message : String(error) });
          }
        }
        const pageCount = document.pageCount;
        document.destroy();
        return { status: 'opened', pageCount, pages };
      }, { widths, maxPages }) as BrowserResult;
    }

    const pages = [];
    for (let index = 0; index < (result.pages?.length ?? 0); index += 1) {
      const rendered = result.pages![index]!;
      const reference = references[index];
      if (!reference || !rendered.dataUrl) {
        pages.push({ page: index + 1, error: rendered.error ?? 'no reference page' });
        continue;
      }
      const actual = PNG.sync.read(Buffer.from(rendered.dataUrl.replace(/^data:image\/png;base64,/u, ''), 'base64'));
      const width = Math.max(reference.width, actual.width);
      const height = Math.max(reference.height, actual.height);
      const ref = padded(reference, width, height);
      const act = padded(actual, width, height);
      const diff = new PNG({ width, height });
      const different = pixelmatch(ref.data, act.data, diff.data, width, height, { threshold: 0.2, includeAA: false });
      let inkUnion = 0;
      for (let i = 0; i < ref.data.length; i += 4) if (ink(ref, i) || ink(act, i)) inkUnion += 1;
      const diffPct = different / (width * height) * 100;
      // Share of the inked area (either side) that differs: whitespace-heavy
      // pages otherwise score near zero whatever their content does.
      const inkDiffPct = inkUnion === 0 ? 0 : Math.min(100, different / inkUnion * 100);
      const layoutDiffPct = layoutDiff(ref, act);
      let image: string | undefined;
      if (layoutDiffPct >= saveMin) {
        image = join(imagesDir, `${file.id}-p${index + 1}.png`);
        writeFileSync(image, PNG.sync.write(sideBySide(ref, act, diff)));
      }
      pages.push({
        page: index + 1,
        diffPct: Number(diffPct.toFixed(4)),
        inkDiffPct: Number(inkDiffPct.toFixed(4)),
        layoutDiffPct: Number(layoutDiffPct.toFixed(4)),
        sizeDeltaPx: [actual.width - reference.width, actual.height - reference.height],
        ...(image ? { image } : {}),
      });
    }
    rmSync(tmp, { recursive: true, force: true });

    writeFileSync(resultPath, `${JSON.stringify({
      id: file.id,
      rel: file.rel,
      status: result.status,
      ...(result.message ? { message: result.message } : {}),
      pdfPageCount: pdfCount,
      ...(result.pageCount !== undefined ? { pageCount: result.pageCount, pageCountMatch: result.pageCount === pdfCount } : {}),
      ...(result.pages ? { pages } : {}),
      ms: Date.now() - started,
    }, null, 2)}\n`);
  });
}
