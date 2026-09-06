import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptFixture, concat, little16, little32, utf16le } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({
  wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)),
});
const record = (options: number, kind: number, payload: Uint8Array) => concat(
  little16(options), little16(kind), little32(payload.length), payload,
);
const info = (flags: number) => record(0, 0x03f9, concat(new Uint8Array(10), little16(flags), new Uint8Array(4)));
const text = record(0, 4000, utf16le('Retained hidden slide content'));
const request = (payload: Uint8Array, master?: Uint8Array) => ({
  bytes: buildPptFixture(concat(payload, text), new Uint8Array(), master),
  from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
} as const);

it.each([0, 1, 2, 4, 8, 0x400, 0xfffb, 0xffff])(
  'preserves only SlideShowSlideInfoAtom fHidden in the PPTX parser: %s', async flags => {
    // MS-PPT 2.6.6 fHidden -> ECMA-376 CT_Slide/@show (default true).
    const converted = await converter.convert(request(info(flags)));
    const archive = new PptxArchive(new Uint8Array(converted.bytes));
    try {
      const model = JSON.parse(new TextDecoder().decode(archive.parse()));
      expect(model.slides).toHaveLength(1);
      expect(model.slides[0].hidden).toBe((flags & 4) !== 0 ? true : undefined);
      const xml = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
      expect(xml).toContain('Retained hidden slide content');
      expect(xml.includes('show="0"')).toBe((flags & 4) !== 0);
      expect(xml).not.toContain('<p:transition');
      expect(xml).not.toContain('<p:timing');
    } finally { archive.free(); }
  },
);

it('does not inherit master or nested visibility when the slide has no info atom', async () => {
  const nested = record(15, 5000, info(4));
  const converted = await converter.convert(request(nested, info(4)));
  const archive = new PptxArchive(new Uint8Array(converted.bytes));
  try {
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    expect(model.slides[0].hidden).toBeUndefined();
  } finally { archive.free(); }
});

it.each([
  concat(info(0), info(4)),
  record(1, 0x03f9, new Uint8Array(16)),
  record(16, 0x03f9, new Uint8Array(16)),
  record(0, 0x03f9, new Uint8Array(15)),
  record(0, 0x03f9, new Uint8Array(17)),
])('rejects ambiguous or malformed slide visibility metadata', async payload => {
  await expect(converter.convert(request(payload)).then(() => 'unexpected-success'))
    .rejects.toMatchObject({ reason: 'unsupported-input' });
});
