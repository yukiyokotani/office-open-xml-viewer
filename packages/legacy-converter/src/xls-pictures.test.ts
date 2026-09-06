import { readFile } from 'node:fs/promises';
import { describe, expect, it, vi } from 'vitest';
import initXlsx, { XlsxArchive } from '../../xlsx/src/wasm/xlsx_parser.js';
import { validateConvertedOoxml } from '@silurus/ooxml-core';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildXlsPicturesFixture, picturePng } from './xls-pictures-fixture.js';
import { buildDocFixture } from './test-fixtures.js';

const wasm = await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url));
await initXlsx({ module_or_path: await readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)) });
const request = (bytes = buildXlsPicturesFixture(), signal = new AbortController().signal) => ({ bytes, from: 'xls' as const, to: 'xlsx' as const, signal, maxOutputBytes: 1024 * 1024 });
const strFromU8 = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const parts = (bytes: Uint8Array | ArrayBuffer) => {
  const archive = new XlsxArchive(new Uint8Array(bytes));
  const result: Record<string, Uint8Array> = {};
  try {
    for (const path of ['xl/media/image1.png', 'xl/drawings/drawing1.xml', 'xl/worksheets/sheet1.xml']) {
      try { result[path] = archive.extract_image(path); } catch { /* absent part */ }
    }
    return result;
  } finally { archive.free(); }
};

describe('measured XLS picture conversion', () => {
  it.each([0, 2, 3] as const)('emits validated picture parts with movement flags %i', async (behavior) => {
    const measure = vi.fn(() => 7);
    const result = await createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: measure })
      .convert(request(buildXlsPicturesFixture({ behavior })));
    expect(measure).toHaveBeenCalledWith({ family: 'Arial', sizePoints: 11, bold: false, italic: false }, expect.any(AbortSignal));
    const zip = parts(result.bytes);
    expect(zip['xl/media/image1.png']).toEqual(picturePng);
    const xml = strFromU8(zip['xl/drawings/drawing1.xml']);
    expect(xml).toContain(`editAs="${behavior === 0 ? 'twoCell' : behavior === 2 ? 'oneCell' : 'absolute'}"`);
    expect(xml).toContain('<xdr:colOff>333375</xdr:colOff>');
    expect(xml).toContain('<xdr:rowOff>95250</xdr:rowOff>');
    expect(xml).toContain('<a:ext cx="1166813" cy="523875"/>');
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'])).toContain('defaultColWidth="10"');
    const archive = new XlsxArchive(new Uint8Array(result.bytes));
    try {
      archive.open_sheet_cursor(0, 'S');
      let worksheet;
      for (let pull = 0; pull < 10; pull++) {
        const value = JSON.parse(strFromU8(archive.pull_sheet_cursor(100)));
        if (archive.sheet_cursor_pull_finished()) {
          worksheet = value.worksheet; archive.acknowledge_sheet_cursor_terminal(); break;
        }
      }
      expect(worksheet.images).toHaveLength(1);
      expect(worksheet.images[0]).toMatchObject({ imagePath: 'xl/media/image1.png', fromColOff: 333375, fromRowOff: 95250, nativeExtCx: 1166813, nativeExtCy: 523875 });
      archive.close_sheet_cursor();
    } finally { archive.free(); }
    await validateConvertedOoxml(result.bytes instanceof Uint8Array ? result.bytes : new Uint8Array(result.bytes), 'xlsx');
  });

  it('does not load metrics or emit pictures unless explicitly configured', async () => {
    const result = await createLegacyOfficeWasmConverter({ wasm }).convert(request());
    expect(Object.keys(parts(result.bytes)).some((p) => p.includes('/drawings/'))).toBe(false);
  });

  it.each([{ hiddenRoot: true }, { nested: true }, { malformedImage: true }, { unknownFont: true }])('omits unsupported pictures without invoking fonts: %j', async (options) => {
    const measure = vi.fn(() => 7);
    const result = await createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: measure }).convert(request(buildXlsPicturesFixture(options)));
    expect(measure).not.toHaveBeenCalled();
    expect(Object.keys(parts(result.bytes)).some((p) => p.includes('/media/'))).toBe(false);
    expect(strFromU8(parts(result.bytes)['xl/worksheets/sheet1.xml'])).toContain('<v>42.5</v>');
    expect(result.warnings?.length).toBeGreaterThan(0);
  });

  it('omits unmeasurable fonts with an explicit warning', async () => {
    const result = await createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: () => undefined }).convert(request());
    expect(result.warnings).toContain('legacy-xls:unmeasured-pictures-omitted');
    expect(parts(result.bytes)['xl/media/image1.png']).toBeUndefined();
  });

  it.each([0, -1, 7.5, Infinity, NaN, 4097])('rejects invalid metric %s', async (width) => {
    const converter = createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: () => width });
    await expect(converter.convert(request())).rejects.toMatchObject({ reason: 'failed' });
  });

  it('releases pending conversion on abort, handles late resolution, and can run again', async () => {
    let resolve: ((width: number) => void) | undefined;
    const measure = vi.fn().mockImplementationOnce(() => new Promise<number>((done) => { resolve = done; })).mockReturnValue(7);
    const converter = createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: measure });
    const controller = new AbortController();
    const pending = converter.convert(request(undefined, controller.signal));
    const failed = expect(pending).rejects.toMatchObject({ reason: 'aborted' });
    await vi.waitFor(() => expect(measure).toHaveBeenCalledOnce());
    await expect(converter.convert(request())).rejects.toMatchObject({ reason: 'capacity-exceeded' });
    controller.abort(); await failed;
    resolve?.(7);
    expect(parts((await converter.convert(request())).bytes)['xl/media/image1.png']).toEqual(picturePng);
  });

  it('does not invoke the XLS measurement hook for DOC', async () => {
    const measure = vi.fn(() => 7);
    await createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: measure }).convert({ ...request(buildDocFixture()), from: 'doc', to: 'docx' });
    expect(measure).not.toHaveBeenCalled();
  });

  it('keeps signed cell fractions instead of clamping them', async () => {
    const result = await createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: () => 7 })
      .convert(request(buildXlsPicturesFixture({ dx: -512 })));
    expect(strFromU8(parts(result.bytes)['xl/drawings/drawing1.xml'])).toContain('<xdr:colOff>-333375</xdr:colOff>');
  });

  it('releases the model after callback and output-budget failures', async () => {
    const measure = vi.fn().mockRejectedValueOnce(new Error('host failure')).mockReturnValue(7);
    const converter = createLegacyOfficeWasmConverter({ wasm, measureXlsNormalFont: measure });
    await expect(converter.convert(request())).rejects.toMatchObject({ reason: 'failed' });
    await expect(converter.convert({ ...request(), maxOutputBytes: 128 })).rejects.toMatchObject({ reason: 'output-too-large' });
    expect(parts((await converter.convert(request())).bytes)['xl/media/image1.png']).toEqual(picturePng);
  });
});
