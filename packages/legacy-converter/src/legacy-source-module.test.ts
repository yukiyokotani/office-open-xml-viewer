import { describe, expect, it } from 'vitest';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { openModelSource as openDoc } from './legacy-doc-source-module.js';
import { openModelSource as openPpt } from './legacy-ppt-source-module.js';
import { openModelSource as openXls } from './legacy-xls-source-module.js';
import { buildDocFixture, buildPptFixture, buildXlsFixture } from './test-fixtures.js';
import { TEST_SOURCE_URLS } from './test-sources.js';

// The real source modules with the real direct-reader WASM, called the way
// the renderer's archive realm calls them.
const modules = [
  { family: 'doc', open: openDoc, fixture: () => buildDocFixture(), stream: 'WordDocument', cursor: 'open_document_cursor' },
  { family: 'xls', open: openXls, fixture: () => buildXlsFixture(), stream: 'Workbook', cursor: 'open_sheet_cursor' },
  { family: 'ppt', open: openPpt, fixture: () => buildPptFixture(), stream: 'PowerPoint Document', cursor: 'presentation_bootstrap' },
] as const;

describe.each(modules)('legacy $family source module with the direct reader', ({ family, open, fixture, stream, cursor }) => {
  const config = { wasmUrl: TEST_SOURCE_URLS[family].wasmUrl };

  it('rejects a malformed container fail-closed without poisoning the runtime', async () => {
    const malformed = new Uint8Array(buildCfbFixture(['Root Entry', stream]));
    await expect(open(malformed, config)).rejects.toSatisfy((error: unknown) => {
      expect(error).not.toBeInstanceOf(WebAssembly.RuntimeError);
      return true;
    });
    const opened = await open(fixture(), config);
    opened.close();
  });

  it('opens a minimal authored file into a closeable renderer archive', async () => {
    const opened = await open(fixture(), config);
    expect(typeof (opened.archive as unknown as Record<string, unknown>)[cursor]).toBe('function');
    if (family === 'doc') {
      // No revision marks and no fRMPrint: the document requests no view change.
      expect((opened as Awaited<ReturnType<typeof openDoc>>).viewDefaults).toEqual({});
    }
    opened.close();
    opened.close();
    expect(() => (opened.archive as unknown as Record<string, unknown>)[cursor]).toThrow(/closed/);
  });
});
