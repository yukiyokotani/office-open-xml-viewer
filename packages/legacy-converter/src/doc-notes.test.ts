import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16, little32 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const convert = (bytes: Uint8Array) => converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });

it('retains formatted footnotes and endnotes with independent IDs and part relationships', async () => {
  const result = await convert(buildDocFixture({ text: 'A\u0002B\u0002\r',
    footnotes: [{ cp: 1, text: '\u0002Footnote 😀\rSecond paragraph\r' }],
    endnotes: [{ cp: 3, text: '\u0002Endnote\r' }], comments: 'IGNORED COMMENT\r',
    headers: ['', 'Header\r', '', '', '', ''],
    characterProperties: new Uint8Array([0x55, 0x08, 1, 0x35, 0x08, 1]),
  }));
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  const xml = (part: string) => new TextDecoder().decode(archive.extract_image(part));
  try {
    expect(xml('word/document.xml')).toContain('<w:footnoteReference w:id="1"/>');
    expect(xml('word/document.xml')).toContain('<w:endnoteReference w:id="1"/>');
    const footnotes = xml('word/footnotes.xml');
    expect(footnotes).toContain('<w:footnote w:id="1">');
    expect(footnotes).toContain('<w:footnoteRef/>');
    expect(footnotes).toContain('Footnote 😀');
    expect(footnotes).toContain('Second paragraph');
    expect(footnotes).toContain('<w:b w:val="1"/>');
    expect(footnotes).not.toContain('Header');
    expect(xml('word/endnotes.xml')).toContain('Endnote');
    expect(xml('word/endnotes.xml')).not.toContain('IGNORED COMMENT');
    expect(xml('word/_rels/document.xml.rels')).toContain('Target="footnotes.xml"');
    expect(xml('[Content_Types].xml')).toContain('wordprocessingml.endnotes+xml');
    const model = JSON.parse(new TextDecoder().decode(archive.parse())) as { footnotes: unknown[]; endnotes: unknown[] };
    expect(model.footnotes).toHaveLength(1); expect(model.endnotes).toHaveLength(1);
  } finally { archive.free(); }
});

it('keeps custom marks as text with customMarkFollows and never revives field instructions', async () => {
  const result = await convert(buildDocFixture({ text: 'A* B\u0013DDE hidden\u0014cached\u0015\r',
    footnotes: [{ cp: 1, text: '* Note\u0013DDE hidden\u0014safe cache\u0015\r', automatic: false }],
  }));
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const body = new TextDecoder().decode(archive.extract_image('word/document.xml'));
    const note = new TextDecoder().decode(archive.extract_image('word/footnotes.xml'));
    expect(body).toContain('<w:footnoteReference w:id="1" w:customMarkFollows="1"/>');
    expect(body).toContain('*'); expect(note).toContain('* Note'); expect(note).toContain('safe cache');
    expect(note).not.toContain('<w:footnoteRef'); expect(body + note).not.toContain('DDE');
  } finally { archive.free(); }
});

it.each([
  { text: 'Ax\r', characterProperties: new Uint8Array([0x55, 0x08, 1]) },
  { text: 'A\u0002\r', characterProperties: new Uint8Array() },
  { text: '😀\r', characterProperties: new Uint8Array([0x55, 0x08, 1]) },
])('rejects invalid automatic note anchors %#', async properties => {
  await expect(convert(buildDocFixture({ ...properties, footnotes: [{ cp: 1, text: '\u0002Note\r' }] })))
    .rejects.toMatchObject({ reason: 'unsupported-input' });
});

it('gives note pictures part-local relationships without losing shared body/header media', async () => {
  // MS-DOC PICFAndOfficeArtData and MS-ODRAW OfficeArtBlipPNG: one passive
  // inline picture shared by all story kinds, never an external resource.
  const record = (kind: number, options: number, data: Uint8Array) => concat(little16(options), little16(kind), little32(data.length), data);
  const png = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=', 'base64'));
  const picf = new Uint8Array(68);
  const view = new DataView(picf.buffer);
  for (const [at, value] of [[4, 68], [6, 100], [28, 1440], [30, 720], [32, 1000], [34, 1000]]) view.setUint16(at, value, true);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1), little32(0x800))),
    record(0xf00b, 0x13, concat(little16(0x0104), little32(1))),
  ));
  const data = concat(picf, shape, record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), png)));
  new DataView(data.buffer).setUint32(0, data.length, true);
  const result = await convert(buildDocFixture({
    text: 'Body\u0001\u0002\u0002\r', data,
    headers: ['', 'Header\u0001\r', '', '', '', ''],
    footnotes: [{ cp: 5, text: '\u0002Footnote\u0001\r' }],
    endnotes: [{ cp: 6, text: '\u0002Endnote\u0001\r' }],
    characterProperties: concat(little16(0x0855), new Uint8Array([1]), little16(0x6a03), little32(0)),
  }));
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    expect(archive.extract_image('word/media/image0.png')).toEqual(png);
    const ids: string[] = [];
    for (const part of ['document', 'header2', 'footnotes', 'endnotes']) {
      const xml = new TextDecoder().decode(archive.extract_image(`word/${part}.xml`));
      const rels = new TextDecoder().decode(archive.extract_image(`word/_rels/${part}.xml.rels`));
      expect(xml).toContain('r:embed="rImg0"');
      expect(rels.match(/Target="media\/image0.png"/g)).toHaveLength(1);
      ids.push(...Array.from(xml.matchAll(/wp:docPr id="(\d+)"/g), m => m[1]));
    }
    expect(new Set(ids).size).toBe(4);
  } finally { archive.free(); }
});

it('does not resurrect an automatic reference hidden inside field instructions', async () => {
  const result = await convert(buildDocFixture({
    text: 'A\u0013\u0002HIDDEN\u0014cached\u0015\r',
    footnotes: [{ cp: 2, text: '\u0002Note\r' }],
    characterProperties: concat(little16(0x0855), new Uint8Array([1])),
  }));
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const xml = new TextDecoder().decode(archive.extract_image('word/document.xml'));
    expect(xml).toContain('cached');
    expect(xml).not.toContain('HIDDEN');
    expect(xml).not.toContain('footnoteReference');
  } finally { archive.free(); }
});
