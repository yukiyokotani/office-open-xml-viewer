// Opt-in measurement of the ODS reader against the local, uncommitted
// corpus: every `.ods` under packages/opendocument/public/private/ods (each
// converted by LibreOffice from the XLSX file at the same relative path under
// packages/xlsx/public/private) is opened through the ODS source, and its
// displayed cell text is compared cell by cell with the original XLSX opened
// by the XLSX reader. The first sheet of both is also painted and compared.
//
//   OOXML_ODS_CORPUS=1 npx vitest run packages/opendocument/src/node/ods-private-corpus.test.ts
//
// Optional: ODS_CORPUS_FILTER=<substring>, ODS_CORPUS_LIMIT=<n>,
// ODS_CORPUS_VISUAL=0 (text only), ODS_CORPUS_DUMP=1 (keep the paintings). Results go to the main checkout's
// git-ignored packages/opendocument/public/private/harness/ods-<label>/.
// Full-corpus runs hold /private/tmp/ooxml-vrt.lock while they run.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import pixelmatch from 'pixelmatch';
import { describe, expect, it } from 'vitest';
import { testOdsSource } from '../test-sources.js';
import {
  formatCellValue, installImageBitmapShim, installOffscreenCanvasShim, materializeXlsxWorkbook,
  renderWorksheetViewport, resolveSharedStrings, skia, skiaFactory,
} from './ods-facade.js';

const run = process.env.OOXML_ODS_CORPUS === '1';
const here = dirname(fileURLToPath(import.meta.url));
const checkout = join(here, '../../../..');

function mainCheckout(): string {
  const common = execFileSync('git', ['rev-parse', '--path-format=absolute', '--git-common-dir'], {
    cwd: checkout, encoding: 'utf8',
  }).trim();
  return dirname(common);
}

type AnyCell = { col: number; row: number; value: { type: string }; styleIndex?: number };
type AnySheet = { name: string; rows: Array<{ index: number; cells: AnyCell[] }> };
type Materialized = Awaited<ReturnType<typeof materializeXlsxWorkbook>>;

/** Take the shared full-corpus lock (a directory), waiting for its holder. */
async function acquire(lock: string): Promise<void> {
  for (;;) {
    try {
      mkdirSync(lock);
      writeFileSync(join(lock, 'owner'), `ods-reader harness ${new Date().toString()}\n`);
      return;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
      await new Promise(resolve => setTimeout(resolve, 15_000));
    }
  }
}

const VIEW = { row: 1, col: 1, rows: 40, cols: 14 } as const;
const VIEW_SIZE = { width: 1100, height: 760 } as const;

async function walk(directory: string, base = directory, out: string[] = []): Promise<string[]> {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    if (entry.name.startsWith('.') || entry.name.startsWith('~$')) continue;
    const path = join(directory, entry.name);
    if (entry.isDirectory()) await walk(path, base, out);
    else if (entry.isFile() && entry.name.endsWith('.ods')) out.push(relative(base, path));
  }
  return out.sort();
}

/** Displayed text of every nonempty cell, keyed `sheet!row:col`. */
function texts(workbook: Materialized): Map<string, string> {
  const out = new Map<string, string>();
  const { styles, workbook: index, sharedStrings } = workbook.workbookIndex as unknown as {
    styles: Parameters<typeof formatCellValue>[1];
    workbook: { date1904?: boolean };
    sharedStrings: unknown[];
  };
  workbook.worksheets.forEach((ws, sheet) => {
    resolveSharedStrings(ws, sharedStrings as never);
    for (const row of (ws as unknown as AnySheet).rows) {
      for (const cell of row.cells) {
        if (cell.value.type === 'empty') continue;
        const text = formatCellValue(cell as never, styles, null, index.date1904 ?? false);
        if (text.trim() !== '') out.set(`${sheet}!${row.index}:${cell.col}`, text);
      }
    }
  });
  return out;
}

async function paint(workbook: Materialized, dump?: string): Promise<Uint8ClampedArray | null> {
  if (!skia || workbook.worksheets.length === 0) return null;
  const factory = skiaFactory();
  const canvas = factory.createCanvas(VIEW_SIZE.width, VIEW_SIZE.height);
  const restoreImage = installImageBitmapShim(factory);
  const restoreOffscreen = installOffscreenCanvasShim(factory);
  try {
    const mdw = (workbook.workbookIndex as { layoutMetrics?: { maximumDigitWidth: number } }).layoutMetrics?.maximumDigitWidth;
    await renderWorksheetViewport(
      { ws: workbook.worksheets[0]!, styles: workbook.workbookIndex.styles } as never,
      canvas as unknown as HTMLCanvasElement, VIEW,
      { ...VIEW_SIZE, dpr: 1, ...(mdw ? { authoritativeMdw: mdw } : {}) } as never,
    );
    if (dump) {
      await writeFile(dump, await (canvas as unknown as { toBuffer(format: string): Promise<Buffer> }).toBuffer('png'));
    }
    const context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    return context.getImageData(0, 0, VIEW_SIZE.width, VIEW_SIZE.height).data;
  } finally {
    restoreOffscreen();
    restoreImage();
  }
}

/** Share of inked pixels (on either side) that differ. */
function visualDiff(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
  const diff = new Uint8ClampedArray(a.length);
  const changed = pixelmatch(a, b, diff, VIEW_SIZE.width, VIEW_SIZE.height, { threshold: 0.2 });
  let inked = 0;
  for (let i = 0; i < a.length; i += 4) {
    const ink = (p: Uint8ClampedArray) => p[i]! < 240 || p[i + 1]! < 240 || p[i + 2]! < 240;
    if (ink(a) || ink(b)) inked++;
  }
  return inked === 0 ? 0 : changed / inked;
}

interface Result {
  file: string;
  opened: boolean;
  error?: string;
  xlsxError?: string;
  sheets?: [number, number];
  cells?: number;
  matched?: number;
  mismatches?: Array<[string, string, string]>;
  visual?: number;
}

describe.skipIf(!run)('ODS private corpus', () => {
  it('opens every workbook and matches the original XLSX cell text', async () => {
    const main = mainCheckout();
    const corpus = join(main, 'packages/opendocument/public/private/ods');
    const originals = join(main, 'packages/xlsx/public/private');
    const filter = process.env.ODS_CORPUS_FILTER;
    const limit = Number(process.env.ODS_CORPUS_LIMIT ?? Infinity);
    let files = (await walk(corpus)).filter(f => !filter || f.includes(filter));
    files = files.slice(0, limit);
    const full = !filter && !Number.isFinite(limit);
    const lock = '/private/tmp/ooxml-vrt.lock';
    if (full) await acquire(lock);
    try {
      const sha = execFileSync('git', ['rev-parse', '--short', 'HEAD'], { cwd: checkout, encoding: 'utf8' }).trim();
      const out = join(main, 'packages/opendocument/public/private/harness', `ods-${process.env.ODS_CORPUS_LABEL ?? sha}`);
      mkdirSync(join(out, 'results'), { recursive: true });
      const results: Result[] = [];
      const visual = process.env.ODS_CORPUS_VISUAL !== '0';
      for (const file of files) {
        const result: Result = { file, opened: false };
        let ods: Materialized | undefined;
        try {
          ods = await materializeXlsxWorkbook(new Uint8Array(await readFile(join(corpus, file))), {
            ods: testOdsSource(), ...(skia ? { factory: skiaFactory() } : {}),
          });
          result.opened = true;
        } catch (error) {
          result.error = error instanceof Error ? `${error.name}: ${error.message}`.slice(0, 300) : String(error);
        }
        const original = join(originals, file.replace(/\.ods$/, '.xlsx'));
        let xlsx: Materialized | undefined;
        if (existsSync(original)) {
          try {
            xlsx = await materializeXlsxWorkbook(new Uint8Array(await readFile(original)),
              skia ? { factory: skiaFactory() } : {});
          } catch (error) {
            result.xlsxError = error instanceof Error ? error.message.slice(0, 200) : String(error);
          }
        }
        if (ods && xlsx) {
          result.sheets = [ods.worksheets.length, xlsx.worksheets.length];
          const a = texts(ods);
          const b = texts(xlsx);
          const keys = new Set([...a.keys(), ...b.keys()]);
          let matched = 0;
          const mismatches: Array<[string, string, string]> = [];
          for (const key of keys) {
            const left = a.get(key) ?? '';
            const right = b.get(key) ?? '';
            if (left === right) matched++;
            else if (mismatches.length < 50) mismatches.push([key, left, right]);
          }
          result.cells = keys.size;
          result.matched = matched;
          result.mismatches = mismatches;
          if (visual) {
            // ODS_CORPUS_DUMP=1 keeps both paintings as PNG for inspection.
            const dump = process.env.ODS_CORPUS_DUMP === '1' ? join(out, 'png', file.replace(/[\\/]/g, '__')) : undefined;
            if (dump) mkdirSync(join(out, 'png'), { recursive: true });
            const [pa, pb] = [await paint(ods, dump && `${dump}.ods.png`), await paint(xlsx, dump && `${dump}.xlsx.png`)];
            if (pa && pb) result.visual = visualDiff(pa, pb);
          }
        }
        results.push(result);
        await writeFile(join(out, 'results', `${file.replace(/[\\/]/g, '__')}.json`), JSON.stringify(result, null, 1));
      }
      const opened = results.filter(r => r.opened).length;
      const compared = results.filter(r => r.cells !== undefined);
      const cells = compared.reduce((s, r) => s + (r.cells ?? 0), 0);
      const matched = compared.reduce((s, r) => s + (r.matched ?? 0), 0);
      const visuals = results.flatMap(r => (r.visual === undefined ? [] : [r.visual])).sort((x, y) => x - y);
      const summary = {
        files: results.length,
        opened,
        openedPct: (100 * opened) / Math.max(1, results.length),
        compared: compared.length,
        cells,
        matched,
        cellTextMatchPct: (100 * matched) / Math.max(1, cells),
        filesFullyMatching: compared.filter(r => r.cells === r.matched).length,
        visualMedian: visuals[Math.floor(visuals.length / 2)] ?? null,
        visualMean: visuals.reduce((s, v) => s + v, 0) / Math.max(1, visuals.length),
        errors: results.filter(r => !r.opened).map(r => [r.file, r.error]),
      };
      await writeFile(join(out, 'summary.json'), JSON.stringify(summary, null, 1));
      const worst = [...compared].sort((x, y) => (x.matched! / x.cells! || 1) - (y.matched! / y.cells! || 1)).slice(0, 25);
      const lines = [
        `# ODS corpus (${sha})`, '',
        `- opened: ${opened}/${results.length} (${summary.openedPct.toFixed(1)}%)`,
        `- cell text match: ${matched}/${cells} (${summary.cellTextMatchPct.toFixed(2)}%) over ${compared.length} workbooks; ${summary.filesFullyMatching} match fully`,
        `- first-sheet visual diff (share of inked pixels): median ${summary.visualMedian?.toFixed(3)}, mean ${summary.visualMean.toFixed(3)}`,
        '', '## Load errors', ...summary.errors.map(([f, e]) => `- ${f}: ${e}`),
        '', '## Lowest cell-text match',
        ...worst.map(r => `- ${r.file}: ${r.matched}/${r.cells}; e.g. ${JSON.stringify(r.mismatches?.slice(0, 3))}`),
      ];
      await writeFile(join(out, 'report.md'), lines.join('\n'));
      console.info(lines.slice(0, 5).join('\n'));
      expect(results.length).toBeGreaterThan(0);
    } finally {
      if (full) rmSync(lock, { recursive: true, force: true });
    }
  }, 3_600_000);
});
