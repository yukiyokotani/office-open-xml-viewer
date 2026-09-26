// BIFF8 XFExt (MS-XLS 2.4.355) extended colors and indentation through the
// direct XLS reader into the ordinary XLSX style model. Theme ZIP parsing,
// version-only default themes, theme XML hardening, continuation reassembly
// and the XFExt style-model rules are unit-tested in the Rust reader
// (xls/theme.rs, xls/theme/xml.rs, xls/styles.rs); these cases prove the
// checksum binding, framing rejections and resolved colors reach the host.
import { describe, expect, it } from 'vitest';
import { concat, little16, little32 } from '../test-fixtures.js';
import { workbookCfb } from '../xls-workbook-fixture.js';
import { buildStoredZip } from '../zip-fixture.js';
import { testXlsSource } from '../test-sources.js';
import { materializeXlsxWorkbook } from './node-facade.js';

const rec = (kind: number, bytes: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(bytes.length), bytes);
const bof = (kind: number) => rec(0x809, concat(little16(0x600), little16(kind), new Uint8Array(12)));
const frt = (kind: number) => concat(little16(kind), new Uint8Array(10));
const prop = (kind: number, data: Uint8Array) => concat(little16(kind), little16(data.length + 4), data);
/** FullColorExt: xclrType, nTintShade, xclrValue, 8 reserved bytes. */
const color = (value = [0x12, 0x34, 0x56, 0x78], type = 2, tint = 0, reserved = 0) =>
  concat(little16(type), little16(tint & 0xffff), new Uint8Array(value), new Uint8Array(8).fill(reserved));
const FILL_FG = 4, FILL_BG = 5, TOP = 7, BOTTOM = 8, LEFT = 9, RIGHT = 10, DIAGONAL = 11, FONT = 13;
const INDENT = 0x000f;

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
  baseIndent?: number;
  readingOrder?: number;
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
  xf[8] = (options.baseIndent ?? 0) | ((options.readingOrder ?? 0) << 6);
  if (options.cellHasExtension === false) xf[17] &= ~2;
  const xfs = Array.from({ length: 16 }, () => xf.slice());
  // A BIFF8 workbook begins with StyleXFs. Keep index zero as the Normal style
  // while the default extension target remains CellXF index one.
  xfs[0][4] = 4; xfs[0][17] &= ~2;
  if (options.styleXf) {
    const index = options.index ?? 1;
    xfs[index][4] = 4; xfs[index][17] &= ~2;
  }
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
  return workbookCfb(concat(globals, bound(size), ...themeRecords, rec(10), bof(16), rec(0x203, number), rec(10)));
}

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const themeNames = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const themeColor = (index: number) => [index + 1, index + 17, index + 33];
const hex = (bytes: number[]) => `#${bytes.map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
const themeXml = () => `<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="Test">${themeNames.map((name, i) =>
  `<a:${name}>${i === 0 ? `<a:sysClr val="windowText" lastClr="${hex(themeColor(i)).slice(1)}"/>` : `<a:srgbClr val="${hex(themeColor(i)).slice(1)}"/>`}</a:${name}>`).join('')}</a:clrScheme></a:themeElements></a:theme>`;
const themeRecords = () => [rec(0x896, concat(frt(0x896), little32(0), buildStoredZip({
  // An unreferenced competing theme must never win by path or enumeration.
  'theme/theme1.xml': themeXml().replace(hex(themeColor(0)).slice(1), 'FFFFFF'),
  '_rels/.rels': `<Relationships xmlns="${R}"><Relationship Id="main" Type="${REL}/officeDocument" Target="settings/manager.xml"/></Relationships>`,
  'settings/manager.xml': `<a:themeManager xmlns:a="${A}"/>`,
  'settings/_rels/manager.xml.rels': `<Relationships xmlns="${R}"><Relationship Id="colors" Type="${REL}/theme" Target="../owned/colors.xml"/></Relationships>`,
  'owned/colors.xml': themeXml(),
})))];
const themed = (index: number, tint = 0) => color([index, 0, 0, 0], 3, tint);
const malformedTheme = [rec(0x896, new Uint8Array([1]))];

async function styles(bytes: Uint8Array) {
  return (await materializeXlsxWorkbook(bytes, { modelSources: [testXlsSource()] })).workbookIndex.styles;
}

async function resolved(bytes: Uint8Array, index = 1) {
  const model = await styles(bytes);
  const xf = model.cellXfs[index]!;
  const border = model.borders[xf.borderId]!;
  return {
    model, xf,
    font: model.fonts[xf.fontId]!.color,
    fill: model.fills[xf.fillId]!,
    border: [border.top, border.bottom, border.left, border.right, border.diagonalUp, border.diagonalDown].map(edge => edge?.color),
  };
}

describe('XLS extended colors', () => {
  it('applies owned RGB extensions to the CellXF fill, every border and a new font variant, ignoring reserved bytes', async () => {
    const properties = [FILL_FG, FILL_BG, TOP, BOTTOM, LEFT, RIGHT, DIAGONAL, FONT].map(kind => prop(kind, color(undefined, 2, 0, 0xab)));
    const { model, xf, font, fill, border } = await resolved(fixture(properties));
    expect(fill).toMatchObject({ patternType: 'solid', fgColor: '#123456', bgColor: '#123456' });
    expect(border).toEqual(Array(6).fill('#123456'));
    expect(font).toBe('#123456');
    // Unextended XFs keep the palette formats and the original font.
    const plain = model.cellXfs[2]!;
    expect([xf.fontId, xf.fillId, xf.borderId]).not.toEqual([plain.fontId, plain.fillId, plain.borderId]);
    expect(model.fonts).toHaveLength(2);
    expect(model.fonts[plain.fontId]!.color).toBeNull();
  });

  it('shares one appended font variant between XFs with the same extension color', async () => {
    const { model, xf } = await resolved(fixture([prop(FONT, color([1, 2, 3, 255]))], { sharedColorXf: 2 }));
    expect(model.cellXfs[2]!.fontId).toBe(xf.fontId);
    expect(model.fonts.map(font => font.color)).toEqual([null, '#010203']);
  });

  it.each([
    [{ stale: true }, false], [{ missingChecksum: true }, false], [{ wrongCount: true }, false],
    [{ cellHasExtension: false }, false],
    // StyleXF bit 1 of byte 17 is reserved, not CellXF.fHasXFExt.
    [{ styleXf: true }, true],
  ] as const)('applies an extension only with a current, owned XF binding: %j', async (options, applied) => {
    const { model } = await resolved(fixture([prop(FONT, color())], options));
    expect(model.fonts.map(font => font.color)).toEqual(applied ? [null, '#123456'] : [null]);
  });

  it.each([
    ['theme without a theme record', themed(4)],
    ['automatic', color([0, 0, 0, 0], 0)],
  ])('keeps the palette format for an unresolvable %s color', async (_label, value) => {
    const { model } = await resolved(fixture([prop(FONT, value)]));
    expect(model.fonts).toHaveLength(1);
  });

  // FullColorExt nTintShade is n/32767; the ECMA-376 18.8.19 tint lightens
  // 010203 by 8191/32767 to 214162 (Rust covers a fill tint independently).
  it('applies a FullColorExt tint to an RGB color', async () => {
    expect((await resolved(fixture([prop(FONT, color([1, 2, 3, 255], 2, 8191))]))).font).toBe('#214162');
  });

  it.each([
    { duplicateChecksum: true }, { duplicateXf: true }, { index: 16 },
    { propertyCount: 2 }, { propertyCount: 1025 }, { tail: new Uint8Array([0]) },
    { checksumType: 0 }, { extensionType: 0 },
  ])('rejects malformed or ambiguous bound extensions: %j', async options => {
    await expect(styles(fixture([prop(FILL_FG, color())], options))).rejects.toThrow(/^UNSUPPORTED:.*BIFF/);
  });

  it.each([
    [prop(FILL_FG, color()), prop(FILL_FG, color())], [prop(FILL_FG, new Uint8Array(15))],
    [prop(FILL_FG, color([0, 0, 0, 0], 5))], [concat(little16(FILL_FG), little16(3))],
  ].map(properties => [properties]))('rejects malformed color payloads and duplicate properties', async properties => {
    await expect(styles(fixture(properties))).rejects.toThrow(/^UNSUPPORTED:.*BIFF/);
  });
});

describe('XLS extended indentation', () => {
  it('applies extended indentation to the owned CellXF and keeps base indent and reading order elsewhere', async () => {
    const { model } = await resolved(fixture([prop(INDENT, little16(250))], { baseIndent: 7, readingOrder: 2 }));
    expect(model.cellXfs[1]).toMatchObject({ indent: 250, readingOrder: 2 });
    expect(model.cellXfs[2]).toMatchObject({ indent: 7, readingOrder: 2 });
  });
});

describe('XLS extended theme colors', () => {
  it('resolves accent and hyperlink slots through the owned embedded theme', async () => {
    const kinds = [FILL_FG, FILL_BG, TOP, BOTTOM, LEFT, RIGHT, DIAGONAL, FONT];
    const { fill, border, font } = await resolved(fixture(kinds.map((kind, i) => prop(kind, themed(i + 4))), { themeRecords: themeRecords() }));
    expect([fill.fgColor, fill.bgColor]).toEqual([hex(themeColor(4)), hex(themeColor(5))]);
    expect(border.slice(0, 4)).toEqual([6, 7, 8, 9].map(i => hex(themeColor(i))));
    expect(border.slice(4)).toEqual([hex(themeColor(10)), hex(themeColor(10))]);
    expect(font).toBe(hex(themeColor(11)));
  });

  // Theme indices 0-3 follow the SpreadsheetML light/dark order (0 = lt1,
  // 1 = dk1, 2 = lt2, 3 = dk2), not MS-XLS 2.5.49's listing.
  it('resolves light/dark indices in SpreadsheetML order and tints a theme color', async () => {
    const { fill, border, font } = await resolved(fixture([
      prop(FILL_FG, themed(0)), prop(FILL_BG, themed(1)), prop(TOP, themed(2)), prop(BOTTOM, themed(3)),
      // 051525 lightened by 8191/32767 (ECMA-376 18.8.19).
      prop(FONT, themed(4, 8191)),
    ], { themeRecords: themeRecords() }));
    expect([fill.fgColor, fill.bgColor, border[0], border[1]]).toEqual([1, 0, 3, 2].map(i => hex(themeColor(i))));
    expect(font).toBe('#134F8C');
  });

  it('rejects a malformed theme that a theme color needs', async () => {
    await expect(styles(fixture([prop(FONT, themed(1))], { themeRecords: malformedTheme })))
      .rejects.toThrow(/^UNSUPPORTED:.*theme/);
  });

  it.each([
    [{ stale: true }, themed(4), [null]], [{ missingChecksum: true }, themed(4), [null]],
    [{ wrongCount: true }, themed(4), [null]], [{ cellHasExtension: false }, themed(4), [null]],
    // An owned literal RGB color never needs the theme.
    [{}, color(), [null, '#123456']],
  ] as const)('does not read theme content unless an owned extension needs it: %j', async (options, value, fonts) => {
    const { model } = await resolved(fixture([prop(FONT, value)], { ...options, themeRecords: malformedTheme }));
    expect(model.fonts.map(font => font.color)).toEqual(fonts);
  });
});
