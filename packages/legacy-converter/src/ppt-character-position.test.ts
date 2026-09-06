import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptFixture, concat, little16, little32 } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const record = (kind: number, bytes: Uint8Array, options = 0) => concat(little16(options), little16(kind), little32(bytes.length), bytes);

function request(position: number, outline: boolean) {
  const body = concat(
    record(3999, little32(4)), record(4008, new TextEncoder().encode('X')),
    record(4001, concat(little32(2), little16(0), little32(0), little32(2), little32(0x80000), little16(position & 65535))),
  );
  const textbox = record(0xf00d, outline ? record(3998, little32(0)) : body, 15);
  const shape = record(0xf004, concat(
    record(0xf00a, concat(little32(42), little32(0xa00)), (202 << 4) | 2),
    record(0xf010, concat(...[0, 0, 5760, 4320].map(little32))), textbox,
  ), 15);
  return converter.convert({
    bytes: buildPptFixture(record(1036, record(0xf002, shape, 15), 15), outline ? body : undefined),
    from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
  });
}

it.each([false, true])('rejects invalid signed character positions through the public converter (outline=%s)', async outline => {
  for (const position of [-32768, -101, 101, 32767]) {
    await expect(request(position, outline)).rejects.toMatchObject({ reason: 'unsupported-input' });
  }
});

it.each([false, true])('retains supported text without inventing a font-relative baseline (outline=%s)', async outline => {
  for (const position of [-100, -1, 0, 1, 100]) {
    const result = await request(position, outline);
    const archive = new PptxArchive(new Uint8Array(result.bytes));
    try {
      const xml = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
      expect(xml).toContain('<a:t>X</a:t>');
      // Input validation is not baseline-projection support (MS-PPT 2.9.14).
      expect(xml).not.toContain('baseline=');
      const model = JSON.parse(new TextDecoder().decode(archive.parse()));
      expect(model.slides[0].elements[0].textBody.paragraphs[0].runs[0].text).toBe('X');
    } finally { archive.free(); }
  }
});
