import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { nonOoxmlInputs, type NonOoxmlFormat } from '@silurus/ooxml-core/testing';

/**
 * Each projection must surface the parser's fail-closed admission rejection as
 * the shared typed `OoxmlError('not-ooxml')`, not the raw `OOXML_NOT_OOXML:`
 * envelope string. Runs against the real parser WASM (git-ignored build
 * output; CI builds it before `pnpm test`), skipping when it is absent.
 */
const wasm = (format: NonOoxmlFormat): string =>
  fileURLToPath(new URL(`../../${format}/src/wasm/${format}_parser_bg.wasm`, import.meta.url));
const formats: ReadonlyArray<NonOoxmlFormat> = ['docx', 'pptx', 'xlsx'];
// The adapter statically imports every format's generated wasm glue, so import
// it only once all of it exists; a static import would fail collection instead
// of skipping.
const ready = formats.every((format) => existsSync(wasm(format)));

describe.skipIf(!ready).each(formats)('%s markdown projection', (format) => {
  it.each(nonOoxmlInputs(format))('rejects %s with OoxmlError not-ooxml', async (_name, input) => {
    const adapter = await import('./index.js');
    const { init, convert } = {
      docx: { init: adapter.initDocxFromBytes, convert: adapter.docxToMarkdown },
      pptx: { init: adapter.initPptxFromBytes, convert: adapter.pptxToMarkdown },
      xlsx: { init: adapter.initXlsxFromBytes, convert: adapter.xlsxToMarkdown },
    }[format];
    init(readFileSync(wasm(format)));
    let error: unknown;
    try {
      convert(input);
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(OoxmlError);
    expect(error).toMatchObject({ code: 'not-ooxml' });
  });
});
