import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { nonOoxmlInputs, type NonOoxmlFormat } from '@silurus/ooxml-core/testing';
import {
  docxToMarkdown,
  initDocxFromBytes,
  initPptxFromBytes,
  initXlsxFromBytes,
  pptxToMarkdown,
  xlsxToMarkdown,
} from './index.js';

/**
 * Each projection must surface the parser's fail-closed admission rejection as
 * the shared typed `OoxmlError('not-ooxml')`, not the raw `OOXML_NOT_OOXML:`
 * envelope string. Runs against the real parser WASM (git-ignored build
 * output; CI builds it before `pnpm test`), skipping when it is absent.
 */
const formats: ReadonlyArray<{
  format: NonOoxmlFormat;
  wasm: URL;
  init: (bytes: Uint8Array) => void;
  convert: (bytes: Uint8Array) => string;
}> = [
  {
    format: 'docx',
    wasm: new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url),
    init: initDocxFromBytes,
    convert: docxToMarkdown,
  },
  {
    format: 'pptx',
    wasm: new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url),
    init: initPptxFromBytes,
    convert: pptxToMarkdown,
  },
  {
    format: 'xlsx',
    wasm: new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url),
    init: initXlsxFromBytes,
    convert: xlsxToMarkdown,
  },
];

describe.each(formats)('$format markdown projection', ({ format, wasm, init, convert }) => {
  const path = fileURLToPath(wasm);
  it.skipIf(!existsSync(path)).each(nonOoxmlInputs(format))(
    'rejects %s with OoxmlError not-ooxml',
    (_name, input) => {
      init(readFileSync(path));
      let error: unknown;
      try {
        convert(input);
      } catch (caught) {
        error = caught;
      }
      expect(error).toBeInstanceOf(OoxmlError);
      expect(error).toMatchObject({ code: 'not-ooxml' });
    },
  );
});
