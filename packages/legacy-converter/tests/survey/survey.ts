// Shared helpers of the local-only legacy corpus surveys (Office PDFs).
import { execFileSync } from 'node:child_process';
import { readFileSync, readdirSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PNG } from 'pngjs';

/** Repository `packages/` directory. */
export const packagesDir = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

export interface Rendered {
  readonly error?: string;
  readonly pages: readonly string[];
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
