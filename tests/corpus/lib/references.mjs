// Office reference PDFs. Each specimen is kept with the PDFs Office exports
// from it, next to it in packages/<fmt>/public/corpus/:
//   <stem>.win.pdf   Office for Windows
//   <stem>.mac.pdf   Office for Mac
// DOCX and PPTX specimens all have them; XLSX specimens only when they hold a
// chart (cell layout is not compared with a PDF). tests/corpus/references.json
// records, per PDF, the SHA-256 of the specimen it was exported from and of
// the PDF itself, so a rebuilt specimen marks its PDFs stale.

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const VARIANTS = ['win', 'mac'];
export const MANIFEST = join(dirname(fileURLToPath(import.meta.url)), '..', 'references.json');

export const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
export const stemOf = (file) => file.replace(/\.(docx|pptx|xlsx)$/, '');
export const pdfName = (file, variant) => `${stemOf(file)}.${variant}.pdf`;
export const isReferencePdf = (name) => /\.(win|mac)\.pdf$/.test(name);

/** Whether a built specimen needs Office reference PDFs. */
export function needsReference(format, built, readZip) {
  if (format !== 'xlsx') return true;
  return readZip(built).names.some((name) => name.startsWith('xl/charts/chart'));
}

export function readManifest() {
  return existsSync(MANIFEST) ? JSON.parse(readFileSync(MANIFEST, 'utf8')) : {};
}

export function writeManifest(manifest) {
  const sorted = Object.fromEntries(Object.entries(manifest).sort(([a], [b]) => (a < b ? -1 : 1)));
  writeFileSync(MANIFEST, `${JSON.stringify(sorted, null, 2)}\n`);
}
