// BIFF8 OfficeArt pictures (MS-XLS 2.4.180 MsoDrawing, MS-ODRAW
// OfficeArtClientAnchorSheet) through the direct XLS reader. Their anchors are
// stored in the Normal font's maximum digit width, which the XLSX host
// measures with the session canvas (host_layout_request /
// configure_host_layout). Store traversal, crop, sharing and geometry budgets
// are unit-tested in xls/pictures.rs, xls/drawing_media.rs and
// xls/drawing_anchors.
import { describe, expect, it } from 'vitest';
import { buildXlsPicturesFixture, picturePng, pictureWmf } from '../xls-pictures-fixture.js';
import { TEST_SOURCE_URLS, testXlsSource } from '../test-sources.js';
import { openModelSource } from '../legacy-xls-source-module.js';
import { materializeXlsxWorkbook } from './node-facade.js';
import { digitWidthFactory } from './xls-host-layout-fixture.js';

const EMU_PER_PIXEL = 9525;
type Width = number | (() => number);

async function sheet(bytes: Uint8Array, width?: Width) {
  const factory = width === undefined ? undefined : digitWidthFactory(width);
  const workbook = await materializeXlsxWorkbook(bytes, {
    modelSources: [testXlsSource()], ...(factory ? { factory } : {}),
  });
  const [worksheet] = workbook.worksheets;
  // Pictures never displace cells.
  expect(worksheet!.rows.flatMap(row => row.cells).map(cell => cell.value)).toEqual([{ type: 'number', number: 42.5 }]);
  return { worksheet: worksheet!, layout: workbook.workbookIndex.layoutMetrics, fonts: factory?.fonts ?? [] };
}

/** The picture bytes a host reads through the source archive. */
async function media(bytes: Uint8Array, path: string) {
  const opened = await openModelSource(bytes, { wasmUrl: TEST_SOURCE_URLS.xls.wasmUrl });
  try {
    opened.archive.configure_host_layout(7);
    opened.archive.parse();
    return opened.archive.extract_image(path);
  } finally { opened.close(); }
}

describe('XLS pictures', () => {
  it.each([[0, 'twoCell'], [2, 'oneCell'], [3, 'absolute']] as const)
  ('anchors a PNG picture with movement flags %i from the measured Normal font', async (behavior, editAs) => {
    const bytes = buildXlsPicturesFixture({ behavior });
    const { worksheet, layout, fonts } = await sheet(bytes, 7);
    // The host measured the authored Normal font: Arial 11pt regular.
    expect(fonts).toEqual([expect.stringMatching(/^14\.6+7?px "Arial",/)]);
    expect(layout).toEqual({ maximumDigitWidth: 7 });
    // Authored STANDARDWIDTH is 10 characters; the anchor starts 512/1024 into
    // column A and the rows keep their explicit BIFF twips (635 EMU/twip).
    expect(worksheet.defaultColWidth).toBe(10);
    expect(worksheet.images).toEqual([expect.objectContaining({
      editAs, mimeType: 'image/png', fromCol: 0, fromColOff: 333375, fromRow: 0, fromRowOff: 95250,
      nativeExtCx: 1166813, nativeExtCy: 523875,
    })]);
    expect(await media(bytes, worksheet.images![0]!.imagePath)).toEqual(picturePng);
  });

  it('keeps WMF media and its MIME type', async () => {
    const bytes = buildXlsPicturesFixture({ imageFormat: 'wmf' });
    const { worksheet } = await sheet(bytes, 7);
    expect(worksheet.images).toEqual([expect.objectContaining({ mimeType: 'image/wmf' })]);
    expect(await media(bytes, worksheet.images![0]!.imagePath)).toEqual(pictureWmf);
  });

  it('keeps authored rotation, flips and signed cell fractions', async () => {
    const { worksheet } = await sheet(buildXlsPicturesFixture({
      rotation: -90 * 65536, flipH: true, flipV: true, dx: -512,
    }), 7);
    expect(worksheet.images?.[0]).toMatchObject({ rotation: -90, flipH: true, flipV: true, fromColOff: -333375 });
  });

  // Ten measured digits per column; MS-XLS cell fractions 512/1024 and
  // 256/1024. Two columns minus the two corner offsets span 17.5 digits. Only
  // the columns depend on the measured width; rows stay explicit twips.
  it.each([
    [undefined, 1, /^14\.6+7?px "Arial",/],
    [undefined, 4096, /^14\.6+7?px "Arial",/],
    [{ family: '検証書体', sizePoints: 12, bold: true, italic: true }, 9, /^italic 700 16px "検証書体",/],
  ] as const)('measures the Normal font %j of each workbook and scales column geometry by width %i', async (normalFont, width, font) => {
    const { worksheet, layout, fonts } = await sheet(buildXlsPicturesFixture(normalFont ? { normalFont } : {}), width);
    expect(fonts).toEqual([expect.stringMatching(font)]);
    expect(layout).toEqual({ maximumDigitWidth: width });
    expect(worksheet.images?.[0]).toMatchObject({
      fromColOff: 5 * width * EMU_PER_PIXEL,
      nativeExtCx: Math.round(17.5 * width * EMU_PER_PIXEL),
      fromRowOff: 95250, nativeExtCy: 523875,
    });
  });

  it.each([
    ['no canvas factory', undefined],
    ['a width above the host bound', 4097],
    ['a non-finite width', Infinity],
  ] as const)('omits pictures without inventing metrics for %s', async (_label, width) => {
    const { worksheet, layout } = await sheet(buildXlsPicturesFixture(), width);
    expect(layout).toBeUndefined();
    expect(worksheet.images ?? []).toEqual([]);
  });

  it('asks for no measurement and omits pictures when the Normal font is unknown', async () => {
    const { worksheet, layout, fonts } = await sheet(buildXlsPicturesFixture({ unknownFont: true }), 7);
    expect(fonts).toEqual([]);
    expect(layout).toBeUndefined();
    expect(worksheet.images ?? []).toEqual([]);
  });

  it('rejects the workbook when the host measurement fails, then opens the next one', async () => {
    await expect(sheet(buildXlsPicturesFixture(), () => { throw new Error('host failure'); })).rejects.toThrow('host failure');
    expect((await sheet(buildXlsPicturesFixture(), 7)).worksheet.images).toHaveLength(1);
  });

  // The converter omitted these pictures with a warning; the direct reader
  // does not drop drawn content and rejects the workbook instead.
  it.each([
    [{ imageFormat: 'wmf', wmfTrailer: new Uint8Array(8).fill(0xa5) }, /BLIP is not a supported image/],
    [{ imageFormat: 'wmf', wmfMalformedBeforeEof: true }, /WMF/],
    [{ malformedImage: true }, /BLIP is not a supported image/],
    [{ hiddenRoot: true }, /grouped or transformed drawing shapes/],
    [{ nested: true }, /drawing group without its anchor/],
  ] as const)('rejects an unsupported referenced picture: %j', async (options, reason) => {
    await expect(sheet(buildXlsPicturesFixture(options), 7)).rejects.toThrow(reason);
  });
});
