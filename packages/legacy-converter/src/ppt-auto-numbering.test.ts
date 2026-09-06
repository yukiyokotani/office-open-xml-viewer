import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptFixture, concat, little16, little32 } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const record = (kind: number, bytes: Uint8Array, options = 0) => concat(little16(options), little16(kind), little32(bytes.length), bytes);
const utf16 = (text: string) => concat(...Array.from(text, c => little16(c.charCodeAt(0))));
const extension = (scheme = 3, start = 1, enabled = 1) => concat(
  little32(0x03800000), little16(65535), little16(enabled), little16(scheme), little16(start), little32(0), little32(0),
);
function input(ext: Uint8Array, groups = [[4, 0]], tag = '___PPT9', flags = 1, owner: 'inline' | 'outline' | 'no-header' = 'inline'): Uint8Array {
  const body = concat(owner === 'no-header' ? new Uint8Array() : record(3999, little32(4)), record(4008, new TextEncoder().encode('A\rB')),
    record(4001, concat(little32(4), little16(0), little32(0x81), little16(flags), little16(0x2022),
      ...groups.map(([count, id]) => concat(little32(count), little32(0x400), little16(id << 10))))));
  const tags = record(5000, record(5002, concat(record(4026, utf16(tag)), record(5003, record(4012, ext))), 15), 15);
  const shape = record(0xf004, concat(
    record(0xf00a, concat(little32(42), little32(0xa00)), (202 << 4) | 2),
    record(0xf010, concat(...[0, 0, 5760, 4320].map(little32))), record(0xf00d, owner === 'outline' ? record(3998, little32(0)) : body, 15), record(0xf011, tags, 15),
  ), 15);
  return buildPptFixture(record(1036, record(0xf002, shape, 15), 15), owner === 'outline' ? body : undefined);
}
async function convert(bytes: Uint8Array) {
  const result = await converter.convert({ bytes, from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = new PptxArchive(new Uint8Array(result.bytes));
  try {
    return { xml: new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml')),
      paragraphs: JSON.parse(new TextDecoder().decode(archive.parse())).slides[0].elements[0].textBody.paragraphs };
  } finally { archive.free(); }
}

it('passes explicit local PPT numbering to the ordinary PPTX parser', async () => {
  const { xml, paragraphs } = await convert(input(extension(3, 6)));
  expect(xml.match(/<a:buAutoNum type="arabicPeriod" startAt="6"\/>/g)).toHaveLength(2);
  expect(xml).not.toContain('<a:buChar');
  expect(paragraphs.map((p: { bullet: unknown }) => p.bullet)).toEqual([
    expect.objectContaining({ type: 'autoNum', numType: 'arabicPeriod', startAt: 6 }),
    expect.objectContaining({ type: 'autoNum', numType: 'arabicPeriod', startAt: 6 }),
  ]);
});

it('binds run groups, skips nonmatching entries and retains explicit starts', async () => {
  const { paragraphs } = await convert(input(concat(extension(0, 9), extension(3, 2), extension(7, 4)), [[2, 1], [2, 2]]));
  expect(paragraphs[0].bullet).toMatchObject({ numType: 'arabicPeriod', startAt: 2 });
  expect(paragraphs[1].bullet).toMatchObject({ numType: 'romanUcPeriod', startAt: 4 });
});

it('does not treat unknown tags, missing run matches or disabled bullets as numbering', async () => {
  for (const bytes of [input(extension(), [[4, 0]], 'notPPT9'), input(extension(), [[4, 1]]), input(extension(3, 1, 0)), input(extension(), [[4, 0]], '___PPT9', 0)]) {
    expect((await convert(bytes)).xml).not.toContain('<a:buAutoNum');
  }
});

it.each(['outline', 'no-header'] as const)('does not attach shape-local PP9 data to %s text', async owner => {
  const { xml, paragraphs } = await convert(input(extension(), [[4, 0]], '___PPT9', 1, owner));
  expect(xml).not.toContain('<a:buAutoNum');
  expect(paragraphs[0].bullet.type).toBe('char');
});

it('keeps conflicting numbering within one paragraph out of the supported subset', async () => {
  const { paragraphs } = await convert(input(concat(extension(3, 1), extension(3, 2)), [[1, 0], [3, 1]]));
  expect(paragraphs[0].bullet.type).toBe('char');
  expect(paragraphs[1].bullet).toMatchObject({ type: 'autoNum', startAt: 2 });
});

it('preserves the character fallback when an explicit scheme or enabling flag is absent', async () => {
  for (const ext of [concat(little32(0x02000000), little16(1), little32(0), little32(0)),
    concat(little32(0x01000000), little16(3), little16(1), little32(0), little32(0))]) {
    expect((await convert(input(ext))).paragraphs[0].bullet.type).toBe('char');
  }
});

it('rejects malformed local numbering through the public error contract', async () => {
  for (const ext of [extension(41), extension(3, 0), extension(3, 32768), extension(3, 1, 2), extension().slice(0, -1)]) {
    await expect(convert(input(ext))).rejects.toMatchObject({ reason: 'unsupported-input' });
  }
});

it.each([
  'alphaLcPeriod', 'alphaUcPeriod', 'arabicParenR', 'arabicPeriod',
  'romanLcParenBoth', 'romanLcParenR', 'romanLcPeriod', 'romanUcPeriod',
  'alphaLcParenBoth', 'alphaLcParenR', 'alphaUcParenBoth', 'alphaUcParenR',
  'arabicParenBoth', 'arabicPlain', 'romanUcParenBoth', 'romanUcParenR',
  'ea1ChsPlain', 'ea1ChsPeriod', 'circleNumDbPlain', 'circleNumWdWhitePlain',
  'circleNumWdBlackPlain', 'ea1ChtPlain', 'ea1ChtPeriod', 'arabic1Minus',
  'arabic2Minus', 'hebrew2Minus', 'ea1JpnKorPlain', 'ea1JpnKorPeriod',
  'arabicDbPlain', 'arabicDbPeriod', 'thaiAlphaPeriod', 'thaiAlphaParenR',
  'thaiAlphaParenBoth', 'thaiNumPeriod', 'thaiNumParenR', 'thaiNumParenBoth',
  'hindiAlphaPeriod', 'hindiNumPeriod', 'ea1JpnChsDbPeriod', 'hindiNumParenR', 'hindiAlpha1Period',
].map((type, id) => ({ type, id })))('preserves numbering enumeration $id as $type', async ({ type, id }) => {
  expect((await convert(input(extension(id, 32767)))).paragraphs[0].bullet).toMatchObject({ type: 'autoNum', numType: type, startAt: 32767 });
});
