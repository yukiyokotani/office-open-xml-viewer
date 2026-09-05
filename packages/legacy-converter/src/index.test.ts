import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import initDocx, { docx_to_markdown } from '../../docx/src/wasm/docx_parser.js';
import initXlsx, { xlsx_to_markdown } from '../../xlsx/src/wasm/xlsx_parser.js';
import initPptx, { pptx_to_markdown } from '../../pptx/src/wasm/pptx_parser.js';
import { validateConvertedOoxml } from '@silurus/ooxml-core';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, buildXlsFixture, buildPptFixture } from './test-fixtures.js';

const converterWasm = readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url));

beforeAll(async () => {
  const [docx, xlsx, pptx] = await Promise.all([
    readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)),
    readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)),
    readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)),
  ]);
  await Promise.all([
    initDocx({ module_or_path: docx }),
    initXlsx({ module_or_path: xlsx }),
    initPptx({ module_or_path: pptx }),
  ]);
});

describe('purpose-built legacy Office WASM converter', () => {
  it('converts a Word 97 Unicode main story into parser-readable DOCX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('doc', 'docx', buildDocFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'docx');
    expect(docx_to_markdown(bytes)).toContain('Hello 日本語');
    expect(result).toMatchObject({
      engine: 'silurus-legacy-office',
      engineVersion: '0.1.0',
      warnings: expect.arrayContaining(['legacy-doc:missing-formatting-tables-default-character-properties']),
    });
  });

  it('converts BIFF8 values into parser-readable XLSX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('xls', 'xlsx', buildXlsFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'xlsx');
    const markdown = xlsx_to_markdown(bytes);
    expect(markdown).toContain('表計算');
    expect(markdown).toContain('42.5');
    expect(markdown).toContain('日本語');
  });

  it('validates DOC settings packages and rejects zero automatic-tab intervals', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    for (const defaultTabTwips of [1, 360, 720, 2160, 65535]) {
      const result = await converter.convert(request('doc', 'docx', buildDocFixture({ defaultTabTwips })));
      await validateConvertedOoxml(asBytes(result.bytes), 'docx');
      expect(docx_to_markdown(asBytes(result.bytes))).toContain('Hello 日本語');
      expect(result.warnings ?? []).not.toContain('legacy-doc:missing-document-properties-default-tab-interval');
    }
    await expect(converter.convert(request('doc', 'docx', buildDocFixture({ defaultTabTwips: 0 }))))
      .rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(converter.convert({
      ...request('doc', 'docx', buildDocFixture({ defaultTabTwips: 360 })), maxOutputBytes: 128,
    })).rejects.toMatchObject({ reason: 'output-too-large' });
  });

  it('converts PowerPoint Unicode text atoms into parser-readable PPTX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('ppt', 'pptx', buildPptFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'pptx');
    expect(pptx_to_markdown(bytes)).toContain('Legacy 日本語 slide');
  });

  it('maps unsupported and bounded-output failures to stable typed reasons', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    await expect(converter.convert(request('doc', 'docx', new Uint8Array([1, 2, 3]))))
      .rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(converter.convert({
      ...request('doc', 'docx', buildDocFixture()),
      maxOutputBytes: 128,
    })).rejects.toMatchObject({ reason: 'output-too-large' });
    await expect(converter.convert({
      ...request('doc', 'docx', buildDocFixture()),
      maxOutputBytes: -1,
    })).rejects.toMatchObject({ reason: 'failed' });
  });
});

function request(
  from: 'doc' | 'xls' | 'ppt',
  to: 'docx' | 'xlsx' | 'pptx',
  bytes: Uint8Array,
) {
  return {
    bytes,
    from,
    to,
    maxOutputBytes: 1024 * 1024,
    signal: new AbortController().signal,
  } as const;
}

function asBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
}
