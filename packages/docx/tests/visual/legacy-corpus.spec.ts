// Local-only legacy DOC fidelity survey against Word-exported PDFs.
// Renders every private `.doc` through the direct DOC source (no OOXML
// generation) and writes per-page reference/actual/diff PNGs plus a summary
// into LEGACY_CORPUS_OUT. It reports; it does not gate or update references.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

const here = dirname(fileURLToPath(import.meta.url));
const enabled = process.env.LEGACY_CORPUS === '1';
const corpus = resolve(here, '../../public/private/doc');
const filter = process.env.LEGACY_CORPUS_FILTER;
const names = enabled && existsSync(corpus)
  ? readdirSync(corpus)
    .filter((name) => name.toLowerCase().endsWith('.doc') && !name.startsWith('~$'))
    .filter((name) => !filter || name.includes(filter))
    .sort()
  : [];

interface Rendered {
  readonly error?: string;
  readonly pages: readonly string[];
}

function pdfPages(pdf: string, prefix: string): PNG[] {
  execFileSync('pdftoppm', ['-png', '-r', '72', pdf, prefix], { stdio: 'ignore' });
  const directory = resolve(prefix, '..');
  const stem = basename(prefix);
  return readdirSync(directory)
    .filter((name) => name.startsWith(`${stem}-`) && name.endsWith('.png'))
    .sort((a, b) => Number(/-(\d+)\.png$/u.exec(a)![1]) - Number(/-(\d+)\.png$/u.exec(b)![1]))
    .map((name) => PNG.sync.read(readFileSync(resolve(directory, name))));
}

function padded(source: PNG, width: number, height: number): PNG {
  const result = new PNG({ width, height });
  result.data.fill(255);
  PNG.bitblt(source, result, 0, 0, Math.min(source.width, width), Math.min(source.height, height), 0, 0);
  return result;
}

function sideBySide(left: PNG, right: PNG): PNG {
  const height = Math.max(left.height, right.height);
  const result = new PNG({ width: left.width + right.width + 8, height });
  result.data.fill(128);
  PNG.bitblt(left, result, 0, 0, left.width, left.height, 0, 0);
  PNG.bitblt(right, result, 0, 0, right.width, right.height, left.width + 8, 0);
  return result;
}

test.describe('legacy DOC corpus survey', () => {
  test.skip(!enabled, 'Set LEGACY_CORPUS=1 and LEGACY_CORPUS_OUT to run');
  test.describe.configure({ mode: 'serial' });

  for (const name of names) {
    test(name, async ({ page }) => {
      test.setTimeout(600_000);
      const out = resolve(process.env.LEGACY_CORPUS_OUT!, 'doc', name.replace(/[^\w.-]+/gu, '_'));
      mkdirSync(out, { recursive: true });
      const pdf = resolve(corpus, name.replace(/\.doc$/iu, '.pdf'));
      const reference = existsSync(pdf) ? pdfPages(pdf, resolve(out, 'word')) : [];
      const width = reference[0]?.width ?? 816;
      const directDoc = resolve(here, '../../../legacy-converter/src/direct-doc.ts');
      await page.goto('/tests/visual/fixture.html');
      const rendered = await page.evaluate(async ({ file, width: requested, module }) => {
        const pages: string[] = [];
        try {
          const { DocxDocument } = await import('/src/document.ts');
          const { math } = await import('/tests/visual/math-engine.ts');
          const { createLegacyDocSource } = await import(/* @vite-ignore */ module);
          const bytes = await (await fetch(`/private/doc/${encodeURIComponent(file)}`)).arrayBuffer();
          const document = await DocxDocument.load(bytes, {
            useGoogleFonts: false,
            math,
            legacyConversion: { doc: { source: createLegacyDocSource() } },
          });
          try {
            for (let index = 0; index < document.pageCount; index += 1) {
              const canvas = window.document.createElement('canvas');
              await document.renderPage(canvas, index, { width: requested, dpr: 1 });
              pages.push(canvas.toDataURL('image/png'));
            }
          } finally {
            document.destroy();
          }
          return { pages };
        } catch (error) {
          return { pages, error: String(error instanceof Error ? error.message : error) };
        }
      }, { file: name, width, module: `/@fs${directDoc}` }) as Rendered;

      const pagesReport = [];
      const count = Math.max(reference.length, rendered.pages.length);
      for (let index = 0; index < count; index += 1) {
        const actual = rendered.pages[index]
          ? PNG.sync.read(Buffer.from(rendered.pages[index]!.split(',')[1]!, 'base64'))
          : undefined;
        const expected = reference[index];
        const label = String(index + 1).padStart(3, '0');
        if (actual) writeFileSync(resolve(out, `actual-${label}.png`), PNG.sync.write(actual));
        if (!actual || !expected) {
          pagesReport.push({ page: index + 1, missing: actual ? 'reference' : 'actual' });
          continue;
        }
        const w = Math.max(actual.width, expected.width);
        const h = Math.max(actual.height, expected.height);
        const a = padded(actual, w, h);
        const e = padded(expected, w, h);
        const diff = new PNG({ width: w, height: h });
        const different = pixelmatch(e.data, a.data, diff.data, w, h, { threshold: 0.2, includeAA: false });
        writeFileSync(resolve(out, `pair-${label}.png`), PNG.sync.write(sideBySide(e, a)));
        pagesReport.push({ page: index + 1, matchPct: Number((100 - different / (w * h) * 100).toFixed(3)) });
      }
      const summary = {
        name,
        error: rendered.error ?? null,
        referencePages: reference.length,
        renderedPages: rendered.pages.length,
        pages: pagesReport,
      };
      writeFileSync(resolve(out, 'summary.json'), JSON.stringify(summary, null, 1));
      console.log(`${name}: ${rendered.error ? `ERROR ${rendered.error}` : `${rendered.pages.length}/${reference.length} pages`}`);
    });
  }
});
