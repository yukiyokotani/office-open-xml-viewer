import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

it.each([false, true])('preserves no-border sentinels before parsing ordinary border flags, old=%s', async old => {
  const initial = concat(little16(0x6426), new Uint8Array([8, 1, 2, 0]));
  const clear = old ? concat(little16(0x6426), new Uint8Array(4).fill(255))
    : concat(little16(0xc650), new Uint8Array([8, 0x12, 0x34, 0x56, 0, 255, 255, 255, 255]));
  const result = await converter.convert({ bytes: buildDocFixture({ text: 'Body\r', paragraphProperties: concat(initial, clear) }),
    from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const xml = new TextDecoder().decode(archive.extract_image('word/document.xml'));
    expect(xml).toContain('<w:pBdr><w:bottom w:val="nil"/></w:pBdr>');
    expect(xml).not.toContain('w:color="0000FF"');
    expect(() => archive.parse()).not.toThrow();
  } finally { archive.free(); }
});

it('rejects malformed BrcOperand lengths without consuming the next property', async () => {
  const bytes = buildDocFixture({ text: 'Body\r', paragraphProperties: concat(
    little16(0xc650), new Uint8Array([7, 0, 0, 0, 0, 8, 1, 0]), little16(0x2406), new Uint8Array([1]),
  ) });
  await expect(converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 }))
    .rejects.toMatchObject({ reason: 'unsupported-input', from: 'doc', to: 'docx' });
});
