// Shared helpers of the local-only legacy corpus surveys (Office PDFs).
//
// Fidelity method: the visual reference is Office opening the actual legacy
// file and exporting a PDF (tools/legacy-office-export.applescript). A legacy
// file paired with its original OOXML, or Office's upgraded OOXML, helps map
// binary records to XML but is not an oracle: saving to or from the old format
// can change or drop features, so a down-saved corpus member is not a lossless
// copy. Loading without an error is a smoke test, not reader completion. Keep
// renderer self-regression against the previous renderer separate from this
// Office comparison. The corpus stays local and is not redistributed.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { basename, dirname, posix, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PNG } from 'pngjs';

/** Repository `packages/` directory. */
export const packagesDir = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

export interface Rendered {
  readonly error?: string;
  readonly pages: readonly string[];
}

/**
 * Corpus inputs with `extension` under `corpus`, subfolders included, as
 * sorted `/`-separated relative paths. `filter` keeps paths containing it.
 * Symbolic links are not followed.
 */
export function corpusFiles(corpus: string, extension: string, filter?: string): string[] {
  if (!existsSync(corpus)) return [];
  const names: string[] = [];
  const walk = (relative: string): void => {
    for (const entry of readdirSync(resolve(corpus, relative), { withFileTypes: true })) {
      const path = posix.join(relative, entry.name);
      if (entry.isDirectory()) {
        if (!entry.name.startsWith('.')) walk(path);
      } else if (
        entry.isFile()
        && entry.name.toLowerCase().endsWith(`.${extension}`)
        && !entry.name.startsWith('~$')
      ) names.push(path);
    }
  };
  walk('');
  return names.filter((name) => !filter || name.includes(filter)).sort();
}

/**
 * Office PDF of the corpus input `name`: the PDF next to it, otherwise the
 * one in a `windows` or `windows-<n>` export folder beside it (the folder
 * with the highest number when several hold it).
 */
export function referencePdf(corpus: string, name: string): string | undefined {
  const directory = resolve(corpus, dirname(name));
  const pdf = `${basename(name).replace(/\.[^.]+$/u, '')}.pdf`;
  if (existsSync(resolve(directory, pdf))) return resolve(directory, pdf);
  return readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && /^windows(?:-\d+)?$/u.test(entry.name))
    .map((entry) => entry.name)
    .sort((a, b) => b.localeCompare(a, 'en', { numeric: true }))
    .map((folder) => resolve(directory, folder, pdf))
    .find((candidate) => existsSync(candidate));
}

/**
 * Output directory of the corpus input `name` under `root`: its own folders,
 * then its sanitized file name, so inputs of different folders never share
 * one.
 */
export function surveyOutput(root: string, name: string): string {
  return resolve(root, dirname(name), basename(name).replace(/[^\w.-]+/gu, '_'));
}

export function pdfPages(pdf: string, prefix: string): PNG[] {
  execFileSync('pdftoppm', ['-png', '-r', '72', pdf, prefix], { stdio: 'ignore' });
  const directory = resolve(prefix, '..');
  const stem = basename(prefix);
  return readdirSync(directory)
    .filter((name) => name.startsWith(`${stem}-`) && name.endsWith('.png'))
    .sort((a, b) => Number(/-(\d+)\.png$/u.exec(a)![1]) - Number(/-(\d+)\.png$/u.exec(b)![1]))
    .map((name) => PNG.sync.read(readFileSync(resolve(directory, name))));
}

export function padded(source: PNG, width: number, height: number): PNG {
  const result = new PNG({ width, height });
  result.data.fill(255);
  PNG.bitblt(source, result, 0, 0, Math.min(source.width, width), Math.min(source.height, height), 0, 0);
  return result;
}

export function sideBySide(left: PNG, right: PNG): PNG {
  const height = Math.max(left.height, right.height);
  const result = new PNG({ width: left.width + right.width + 8, height });
  result.data.fill(128);
  PNG.bitblt(left, result, 0, 0, left.width, left.height, 0, 0);
  PNG.bitblt(right, result, 0, 0, right.width, right.height, left.width + 8, 0);
  return result;
}


/** Dev-server origin of the viewer package that renders `format`. */
export function viewerOrigin(format: 'doc' | 'ppt' | 'xls'): string {
  return `http://127.0.0.1:${surveyPorts()[format]}`;
}

/** VRT_PORT (default 5351) serves DOC, +1 PPT and +2 XLS. */
export function surveyPorts(): Record<'doc' | 'ppt' | 'xls', number> {
  const base = Number(process.env.VRT_PORT ?? 5351);
  if (!Number.isInteger(base) || base < 1 || base > 65_533) {
    throw new Error(`invalid VRT_PORT: ${process.env.VRT_PORT}`);
  }
  return { doc: base, ppt: base + 1, xls: base + 2 };
}
