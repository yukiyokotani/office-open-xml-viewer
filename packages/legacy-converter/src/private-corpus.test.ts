import { readdir, readFile } from 'node:fs/promises';
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
  it('has a legacy counterpart for every modern private sample', async () => {
    for (const format of formats) {
      const names = (await readdir(format.directory)).filter(isCorpusFile);
      const modern = names.filter((name) => extension(name) === format.to);
      const legacy = new Set(names.filter((name) => extension(name) === format.from));

      expect(legacy.size).toBe(modern.length);
      for (const name of modern) {
        expect(legacy.has(replaceExtension(name, format.from))).toBe(true);
      }
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
      const names = (await readdir(format.directory))
        .filter((name) => isCorpusFile(name) && extension(name) === format.from)
        .sort();
      expect(names.length).toBeGreaterThan(0);

      for (const name of names) {
        const source = await readFile(new URL(name, format.directory));
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

function isCorpusFile(name: string): boolean {
  return !name.startsWith('~$') && name.includes('.');
}

function extension(name: string): string {
  return name.slice(name.lastIndexOf('.') + 1).toLowerCase();
}

function replaceExtension(name: string, next: LegacyOfficeFormat): string {
  return `${name.slice(0, name.lastIndexOf('.'))}.${next}`;
}
