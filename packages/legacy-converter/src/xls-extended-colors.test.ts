import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { buildCfbWithStreams, buildStoredZip } from '@silurus/ooxml-core/testing';
import initXlsx, { XlsxArchive } from '../../xlsx/src/wasm/xlsx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { concat, little16, little32 } from './test-fixtures.js';

await initXlsx({ module_or_path: await readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const rec = (kind: number, bytes: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(bytes.length), bytes);
const bof = (kind: number) => rec(0x809, concat(little16(0x600), little16(kind), new Uint8Array(12)));
const frt = (kind: number) => concat(little16(kind), new Uint8Array(10));
const prop = (kind: number, data: Uint8Array) => concat(little16(kind), little16(data.length + 4), data);
const color = (rgba = [0x12, 0x34, 0x56, 0x78], type = 2, tint = 0) =>
  concat(little16(type), little16(tint & 0xffff), new Uint8Array(rgba), new Uint8Array(8));

// Independently expressed byte-wise MS-OSHARED polynomial division for fixtures.
function checksum(bytes: Uint8Array) {
  let crc = 0;
  for (const byte of bytes) {
    crc ^= byte << 24;
    for (let bit = 0; bit < 8; bit++) crc = (crc << 1) ^ (crc < 0 ? 0xaf : 0);
  }
  return crc >>> 0;
}

interface Options {
  stale?: boolean;
  missingChecksum?: boolean;
  wrongCount?: boolean;
  duplicateChecksum?: boolean;
  duplicateXf?: boolean;
  sharedColorXf?: number;
  index?: number;
  propertyCount?: number;
  tail?: Uint8Array;
  cellHasExtension?: boolean;
  styleXf?: boolean;
  checksumType?: number;
  extensionType?: number;
  themeRecords?: Uint8Array[];
}

function fixture(properties: Uint8Array[], options: Options = {}) {
  const name = new TextEncoder().encode('Arial');
  const font = new Uint8Array(16 + name.length);
  font.set(little16(220)); font.set(little16(0x7fff), 4); font.set(little16(400), 6);
  font[14] = name.length; font.set(name, 16);
  const xf = new Uint8Array(20);
  // Solid fill and all five thin borders; both diagonal direction flags.
  xf.set(little32(0xc0001111), 10);
  xf.set(little32((1 << 26) | (1 << 25) | (1 << 21)), 14);
  if (options.cellHasExtension === false) xf[17] &= ~2;
  if (options.styleXf) { xf[4] = 4; xf[17] &= ~2; }
  const xfs = Array.from({ length: 16 }, () => xf);
  const check = rec(0x87c, concat(frt(options.checksumType ?? 0x87c), little16(0),
    little16(options.wrongCount ? 15 : xfs.length), little32(checksum(concat(...xfs)) ^ Number(Boolean(options.stale)))));
  const ext = rec(0x87d, concat(frt(options.extensionType ?? 0x87d), little16(0),
    little16(options.index ?? 1), little16(0), little16(options.propertyCount ?? properties.length),
    ...properties, options.tail ?? new Uint8Array()));
  const otherExt = ext.slice();
  otherExt.set(little16(options.sharedColorXf ?? 1), 18);
  const globals = concat(bof(5), rec(0x31, font), ...xfs.map(x => rec(0xe0, x)),
    ...(options.missingChecksum ? [] : [check]), ...(options.duplicateChecksum ? [check] : []),
    ext, ...(options.duplicateXf ? [ext] : []), ...(options.sharedColorXf === undefined ? [] : [otherExt]));
  const number = new Uint8Array(14); number[4] = 1;
  new DataView(number.buffer).setFloat64(6, 42, true);
  const bound = (offset: number) => rec(0x85, concat(little32(offset), new Uint8Array([0, 0, 1, 0, 65])));
  const themeRecords = options.themeRecords ?? [];
  const size = concat(globals, bound(0), ...themeRecords, rec(10)).length;
  const stream = concat(globals, bound(size), ...themeRecords, rec(10), bof(16), rec(0x203, number), rec(10));
  return new Uint8Array(buildCfbWithStreams([{ name: 'Workbook', data: stream }]));
}

async function convert(bytes: Uint8Array) {
  const output = await converter.convert({ bytes, from: 'xls', to: 'xlsx',
    signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = new XlsxArchive(new Uint8Array(output.bytes));
  try {
    return { xml: new TextDecoder().decode(archive.extract_image('xl/styles.xml')),
      model: JSON.parse(new TextDecoder().decode(archive.parse())).styles };
  } finally { archive.free(); }
}

it.each([[4, 'fgColor'], [5, 'bgColor'], [7, 'top'], [8, 'bottom'], [9, 'left'], [10, 'right'], [11, 'diagonal'], [13, 'font']] as const)
('preserves RGBA extension %s as ARGB in the corresponding OOXML style', async (kind, tag) => {
  const { xml, model } = await convert(fixture([prop(kind, color())]));
  if (kind === 13) {
    expect(xml).toContain('<color rgb="78123456"/>');
    expect(model.cellXfs[1].fontId).not.toBe(model.cellXfs[0].fontId);
    expect(model.fonts[model.cellXfs[1].fontId].color).toBe('#123456');
    expect(xml.match(/<font>/g)).toHaveLength(2);
  } else if (kind === 4 || kind === 5) {
    expect(xml).toContain(`<${tag} rgb="78123456"/>`);
    expect(model.cellXfs[1].fillId).not.toBe(model.cellXfs[0].fillId);
  } else {
    expect(xml).toContain(`<${tag} style="thin"><color rgb="78123456"/></${tag}>`);
    expect(model.cellXfs[1].borderId).not.toBe(model.cellXfs[0].borderId);
  }
  expect(xml).toContain('<cellXfs count="16">');
});

it('shares appended font variants without mutating the original font', async () => {
  const { xml, model } = await convert(fixture([prop(13, color([1, 2, 3, 255]))], { sharedColorXf: 2 }));
  expect(model.cellXfs.filter((xf: { fontId: number }) => xf.fontId === model.cellXfs[1].fontId)).toHaveLength(2);
  expect(xml.match(/<font>/g)).toHaveLength(2);
  expect(xml).toContain('<color auto="1"/>');
});

it.each([{ stale: true }, { missingChecksum: true }, { wrongCount: true }, { cellHasExtension: false }])
('does not apply an extension without a current, owned XF binding: %j', async options => {
  const { xml } = await convert(fixture([prop(13, color())], options));
  expect(xml).not.toContain('78123456');
});

it('does not interpret the reserved StyleXF bit as CellXF.fHasXFExt', async () => {
  expect((await convert(fixture([prop(13, color())], { styleXf: true }))).xml).toContain('78123456');
});

it.each([color([4, 0, 0, 0], 3), color([1, 2, 3, 255], 2, 8191), color([0, 0, 0, 0], 0)])
('retains palette fallback for color modes that need a separate resolver', async value => {
  const { xml } = await convert(fixture([prop(13, value)]));
  expect(xml.match(/<font>/g)).toHaveLength(1);
});

it.each([
  { duplicateChecksum: true }, { duplicateXf: true }, { index: 16 },
  { propertyCount: 2 }, { propertyCount: 1025 }, { tail: new Uint8Array([0]) },
  { checksumType: 0 }, { extensionType: 0 },
])('rejects malformed or ambiguous bound extensions: %j', async options => {
  await expect(convert(fixture([prop(4, color())], options))).rejects.toMatchObject({ reason: 'unsupported-input' });
});

it.each([
  [prop(4, color()), prop(4, color())], [prop(4, new Uint8Array(15))],
  [prop(4, color([0, 0, 0, 0], 5))], [concat(little16(4), little16(3))],
].map(properties => [properties]))('rejects malformed color payloads and duplicate properties', async properties => {
  await expect(convert(fixture(properties))).rejects.toMatchObject({ reason: 'unsupported-input' });
});

it('ignores reserved FullColorExt bytes instead of using them as color data', async () => {
  const value = color(); value.fill(0xab, 8);
  expect((await convert(fixture([prop(4, value)]))).xml).toContain('78123456');
});

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const themeNames = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const themeColor = (index: number) => [index + 1, index + 17, index + 33];
const hex = (bytes: number[]) => bytes.map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
function themeXml() {
  return `<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="Test">${themeNames.map((name, i) =>
    `<a:${name}>${i === 0 ? `<a:sysClr val="windowText" lastClr="${hex(themeColor(i))}"/>` : `<a:srgbClr val="${hex(themeColor(i))}"/>`}</a:${name}>`).join('')}</a:clrScheme></a:themeElements></a:theme>`;
}
function themeZip(xml = themeXml(), targetMode = '') {
  return buildStoredZip({
    // An unreferenced competing theme must never win by path or enumeration.
    'theme/theme1.xml': themeXml().replace(hex(themeColor(0)), 'FFFFFF'),
    '_rels/.rels': `<Relationships xmlns="${R}"><Relationship Id="main" Type="${REL}/officeDocument" Target="settings/manager.xml"/></Relationships>`,
    'settings/manager.xml': `<a:themeManager xmlns:a="${A}"/>`,
    'settings/_rels/manager.xml.rels': `<Relationships xmlns="${R}"><Relationship Id="colors" Type="${REL}/theme" Target="../owned/colors.xml" ${targetMode}/></Relationships>`,
    'owned/colors.xml': xml,
  });
}
const themeRecords = (zip: Uint8Array) => [rec(0x896, concat(frt(0x896), little32(0), zip))];

it.each(themeNames.map((name, index) => [name, index] as const))('resolves owned theme slot %s through WASM and the ordinary XLSX parser', async (_name, index) => {
  const { xml, model } = await convert(fixture([prop(13, color([index, 0, 0, 0], 3))], { themeRecords: themeRecords(themeZip()) }));
  expect(xml).toContain(`<color rgb="FF${hex(themeColor(index))}"/>`);
  expect(model.fonts[model.cellXfs[1].fontId].color).toBe(`#${hex(themeColor(index))}`);
});

it.each([[4, 'fgColor'], [5, 'bgColor'], [7, 'top'], [8, 'bottom'], [9, 'left'], [10, 'right'], [11, 'diagonal']] as const)
('resolves theme color for fill or border property %s', async (kind, name) => {
  const { xml } = await convert(fixture([prop(kind, color([4, 0, 0, 0], 3))], { themeRecords: themeRecords(themeZip()) }));
  const colorXml = `rgb="FF${hex(themeColor(4))}"`;
  expect(xml).toContain(kind <= 5 ? `<${name} ${colorXml}/>` : `<${name} style="thin"><color ${colorXml}/></${name}>`);
});

it.each([{ stale: true }, { missingChecksum: true }, { wrongCount: true }])
('does not read theme content without a current XF binding: %j', async options => {
  const { xml } = await convert(fixture([prop(13, color([4, 0, 0, 0], 3))], {
    ...options, themeRecords: [rec(0x896, new Uint8Array([1]))],
  }));
  expect(xml.match(/<font>/g)).toHaveLength(1);
});

it('assembles a large theme across ContinueFrt12 records without exposing package content', async () => {
  const zip = themeZip(themeXml().replace('</a:theme>', `<!--${'padding'.repeat(2000)}--></a:theme>`));
  expect(zip.length).toBeGreaterThan(8224);
  const records = [rec(0x896, concat(frt(0x896), little32(0), zip.subarray(0, 8208)))];
  for (let i = 8208; i < zip.length; i += 8212) records.push(rec(0x87f, concat(frt(0x87f), zip.subarray(i, i + 8212))));
  expect((await convert(fixture([prop(4, color([4, 0, 0, 0], 3))], { themeRecords: records }))).xml)
    .toContain(`<fgColor rgb="FF${hex(themeColor(4))}"/>`);
});

it('does not inflate or interpret theme data when only literal colors are requested', async () => {
  expect((await convert(fixture([prop(13, color())], { themeRecords: [rec(0x896, new Uint8Array([1]))] }))).xml)
    .toContain('78123456');
});

it('does not read a theme for an unowned CellXF extension', async () => {
  const { xml } = await convert(fixture([prop(13, color([4, 0, 0, 0], 3))], {
    cellHasExtension: false, themeRecords: [rec(0x896, new Uint8Array([1]))],
  }));
  expect(xml.match(/<font>/g)).toHaveLength(1);
});

it.each([124226, 123820])('retains base palette rather than inventing version-only default theme %s', async version => {
  const { xml } = await convert(fixture([prop(13, color([4, 0, 0, 0], 3))], {
    themeRecords: [rec(0x896, concat(frt(0x896), little32(version)))],
  }));
  expect(xml.match(/<font>/g)).toHaveLength(1);
});

it.each([
  themeZip(themeXml(), 'TargetMode="External"'),
  themeZip(`<!DOCTYPE theme [<!ENTITY ex SYSTEM "https://example.invalid/data">]>${themeXml()}`),
  themeZip(themeXml().replace('name="Test"', 'name="Test" name="Again"')),
  themeZip(themeXml().replace('lastClr=', 'x:spoof=')),
])('rejects external or malformed embedded themes without evaluating their content', async zip => {
  await expect(convert(fixture([prop(13, color([0, 0, 0, 0], 3))], { themeRecords: themeRecords(zip) })))
    .rejects.toMatchObject({ reason: 'unsupported-input' });
});

it('does not erase a theme transform or a FullColorExt tint to force a color match', async () => {
  const transformed = themeXml().replace(`<a:srgbClr val="${hex(themeColor(4))}"/>`, `<a:srgbClr val="${hex(themeColor(4))}"><a:tint val="50000"/></a:srgbClr>`);
  for (const [zip, tint] of [[themeZip(transformed), 0], [themeZip(), 8191]] as const) {
    const { xml } = await convert(fixture([prop(13, color([4, 0, 0, 0], 3, tint))], { themeRecords: themeRecords(zip) }));
    expect(xml.match(/<font>/g)).toHaveLength(1);
  }
});
