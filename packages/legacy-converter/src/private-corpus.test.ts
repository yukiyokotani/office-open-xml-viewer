import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import initDocx, { docx_to_markdown } from '../../docx/src/wasm/docx_parser.js';
import initXlsx, { xlsx_to_markdown } from '../../xlsx/src/wasm/xlsx_parser.js';
import initPptx, { pptx_to_markdown } from '../../pptx/src/wasm/pptx_parser.js';
import {
  HARD_MAX_LEGACY_CONVERSION_BYTES,
  validateConvertedOoxml,
  type LegacyOfficeFormat,
} from '@silurus/ooxml-core';
import { createLegacyOfficeWasmConverter } from './index.js';

const runPrivateCorpus = process.env.OOXML_LEGACY_CORPUS === '1';
const corpusTest = runPrivateCorpus ? describe : describe.skip;
const corpusRoot = process.env.OOXML_LEGACY_CORPUS_ROOT;

async function corpusFiles(format: (typeof formats)[number]): Promise<{ directory: string; names: string[] }> {
  const directory = corpusRoot ? join(corpusRoot, 'packages', format.to, 'public', 'private') : fileURLToPath(format.directory);
  const names: string[] = [];
  // Retain relative directories: equal basenames are distinct corpus inputs.
  async function walk(relative: string): Promise<void> {
    for (const entry of await readdir(join(directory, relative), { withFileTypes: true })) {
      if (entry.name.startsWith('.') || entry.name.startsWith('~$')) continue;
      const path = join(relative, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (entry.isFile()) names.push(path);
      // Do not follow symlinks out of the selected corpus.
    }
  }
  await walk('');
  return { directory, names: names.sort() };
}

const formats = [
  {
    from: 'doc',
    to: 'docx',
    directory: new URL('../../docx/public/private/', import.meta.url),
    parse: docx_to_markdown,
  },
  {
    from: 'xls',
    to: 'xlsx',
    directory: new URL('../../xlsx/public/private/', import.meta.url),
    parse: xlsx_to_markdown,
  },
  {
    from: 'ppt',
    to: 'pptx',
    directory: new URL('../../pptx/public/private/', import.meta.url),
    parse: pptx_to_markdown,
  },
] as const satisfies ReadonlyArray<{
  from: LegacyOfficeFormat;
  to: 'docx' | 'xlsx' | 'pptx';
  directory: URL;
  parse: (bytes: Uint8Array) => string;
}>;

corpusTest('local Office-produced legacy corpus', () => {
  it('discovers nonempty binary inputs without requiring a lossless OOXML round trip', async () => {
    for (const format of formats) {
      const { names } = await corpusFiles(format);
      const legacy = new Set(names.filter((name) => extension(name) === format.from));

      // Native binary documents may not have an Office-upgraded counterpart.
      expect(legacy.size).toBeGreaterThan(0);
    }
  });

  it('converts every legacy sample into OOXML accepted by the existing parsers', async () => {
    const [converterWasm, docxWasm, xlsxWasm, pptxWasm] = await Promise.all([
      readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)),
      readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)),
      readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)),
      readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)),
    ]);
    await Promise.all([
      initDocx({ module_or_path: docxWasm }),
      initXlsx({ module_or_path: xlsxWasm }),
      initPptx({ module_or_path: pptxWasm }),
    ]);
    const converter = createLegacyOfficeWasmConverter({ wasm: converterWasm });

    for (const format of formats) {
      const { directory, names: allNames } = await corpusFiles(format);
      const names = allNames.filter((name) => extension(name) === format.from);
      expect(names.length).toBeGreaterThan(0);

      for (const name of names) {
        const source = await readFile(join(directory, name));
        const result = await converter.convert({
          bytes: source,
          from: format.from,
          to: format.to,
          maxOutputBytes: HARD_MAX_LEGACY_CONVERSION_BYTES,
          signal: new AbortController().signal,
        });
        const bytes = result.bytes instanceof Uint8Array
          ? result.bytes
          : new Uint8Array(result.bytes);
        await validateConvertedOoxml(bytes, format.to);
        expect(() => format.parse(bytes)).not.toThrow();
      }
    }
  }, 300_000);
});

function extension(name: string): string {
  return name.slice(name.lastIndexOf('.') + 1).toLowerCase();
}
