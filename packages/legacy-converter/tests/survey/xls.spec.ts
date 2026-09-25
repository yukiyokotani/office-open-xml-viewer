// Local-only legacy XLS survey against Office-exported PDFs, run from the
// legacy converter package (see playwright.config.ts). The page is the
// XLSX package's own VRT fixture on its dev server; this package supplies
// the direct XLS source, so no OOXML package test imports the converter.
import { existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { packagesDir, viewerOrigin, padded, pdfPages, sideBySide, type Rendered } from './survey.js';

const enabled = process.env.LEGACY_CORPUS === '1';
const corpus = resolve(packagesDir, 'xlsx/public/private/xls');
const filter = process.env.LEGACY_CORPUS_FILTER;
const names = enabled && existsSync(corpus)
  ? readdirSync(corpus)
    .filter((name) => name.toLowerCase().endsWith('.xls') && !name.startsWith('~$'))
    .filter((name) => !filter || name.includes(filter))
    .sort()
  : [];

test.describe('legacy XLS corpus survey', () => {
  test.skip(!enabled, 'Set LEGACY_CORPUS=1 and LEGACY_CORPUS_OUT to run');
  test.describe.configure({ mode: 'serial' });

  for (const name of names) {
    test(name, async ({ page }) => {
      test.setTimeout(600_000);
      const out = resolve(process.env.LEGACY_CORPUS_OUT!, 'xls', name.replace(/[^\w.-]+/gu, '_'));
      mkdirSync(out, { recursive: true });
      const pdf = resolve(corpus, name.replace(/\.xls$/iu, '.pdf'));
      const reference = existsSync(pdf) ? pdfPages(pdf, resolve(out, 'excel')) : [];
      const width = 1100;
      const directXls = resolve(packagesDir, 'legacy-converter/src/direct-xls.ts');
      await page.goto(`${viewerOrigin('xls')}/tests/visual/fixture.html`);
      const rendered = await page.evaluate(async ({ file, width: requested, module }) => {
        const pages: string[] = [];
        try {
          const { XlsxWorkbook } = await import('/src/workbook.ts');
          const { createLegacyXlsSource } = await import(/* @vite-ignore */ module);
          // The dev server decodes paths with decodeURI, which keeps reserved
          // escapes such as %2B; encodeURI leaves those characters literal.
          const response = await fetch(`/private/xls/${encodeURI(file)}`);
          if (!response.ok) throw new Error(`fetch failed: ${response.status}`);
          const bytes = await response.arrayBuffer();
          // Excel column widths depend on the Normal font's maximum digit
          // width in whole pixels (ECMA-376 §18.3.1.13). The library default
          // measures only an installed face; this survey measures whatever
          // face the browser resolves so drawings remain reviewable.
          const measure = (font: { family: string; sizePoints: number; bold: boolean; italic: boolean }) => {
            const context = document.createElement('canvas').getContext('2d')!;
            const px = font.sizePoints * 96 / 72;
            context.font = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${px}px "${font.family}"`;
            let widest = 0;
            for (const digit of '0123456789') widest = Math.max(widest, context.measureText(digit).width);
            return Math.max(1, Math.round(widest));
          };
          const workbook = await XlsxWorkbook.load(bytes, {
            legacyConversion: { xls: { source: createLegacyXlsSource() } },
            measureLegacyXlsNormalFont: measure,
          });
          try {
            for (let index = 0; index < workbook.sheetCount; index += 1) {
              const canvas = window.document.createElement('canvas');
              await workbook.renderViewport(
                canvas,
                index,
                { row: 1, col: 1, rows: 45, cols: 14 },
                { width: requested, height: 850, dpr: 1 },
              );
              pages.push(canvas.toDataURL('image/png'));
            }
          } finally {
            workbook.destroy();
          }
          return { pages };
        } catch (error) {
          return { pages, error: String(error instanceof Error ? error.message : error) };
        }
      }, { file: name, width, module: `/@fs${directXls}` }) as Rendered;

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
