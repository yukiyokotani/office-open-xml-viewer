import { expect, test } from '@playwright/test';
import { deflateSync } from 'node:zlib';
import { storeZip } from '../../src/conformance/generate.js';

const LARGE_RASTER_WIDTH = 4_096;
const LARGE_RASTER_HEIGHT = 4_096;
const LARGE_RASTER_COUNT = 3;
const LARGE_RASTER_NATURAL_RGBA_BYTES =
  LARGE_RASTER_WIDTH * LARGE_RASTER_HEIGHT * 4 * LARGE_RASTER_COUNT;

type Rgba = readonly [number, number, number, number];

interface SyntheticImageCase {
  readonly id: string;
  readonly filename: string;
  readonly bytes: Uint8Array;
  /** Expected pixel after the image is composited onto the white DOCX page. */
  readonly expected: Rgba;
}

interface BrowserPageReport {
  readonly id: string;
  readonly index: number;
  readonly pixel: Rgba;
  readonly maxChannelDelta: number;
  readonly canvasWidth: number;
  readonly canvasHeight: number;
  readonly mountedPages: readonly number[];
}

interface BrowserModeReport {
  readonly mode: 'main' | 'worker';
  readonly pageCount: number;
  readonly layoutComplete: boolean;
  readonly errors: readonly {
    readonly name: string;
    readonly message: string;
    readonly code?: string;
    readonly metric?: string;
  }[];
  readonly pages: readonly BrowserPageReport[];
}

interface BrowserAcceptanceReport {
  readonly main: BrowserModeReport;
  readonly worker: BrowserModeReport;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, payload: Uint8Array): Buffer {
  const typeBytes = Buffer.from(type, 'ascii');
  const output = Buffer.alloc(12 + payload.byteLength);
  output.writeUInt32BE(payload.byteLength, 0);
  typeBytes.copy(output, 4);
  Buffer.from(payload).copy(output, 8);
  output.writeUInt32BE(
    crc32(output.subarray(4, 8 + payload.byteLength)),
    8 + payload.byteLength,
  );
  return output;
}

/** Valid, highly compressible one-bit indexed PNG with a single visible color. */
function indexedPng(width: number, height: number, color: readonly [number, number, number]): Uint8Array {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 1;
  ihdr[9] = 3;
  const scanlines = Buffer.alloc((Math.ceil(width / 8) + 1) * height);
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('PLTE', Uint8Array.from(color)),
    pngChunk('IDAT', deflateSync(scanlines, { level: 9 })),
    pngChunk('IEND', new Uint8Array()),
  ]);
}

interface UncompressedTiffOptions {
  readonly width: number;
  readonly height: number;
  readonly bitsPerSample: readonly number[];
  readonly photometric: 0 | 1 | 2;
  readonly samplesPerPixel: number;
  readonly pixels: Uint8Array;
  readonly extraSample?: 1 | 2;
}

interface TiffIfdEntry {
  readonly tag: number;
  readonly type: 3 | 4;
  readonly values: readonly number[];
}

/** Minimal classic-TIFF v42, one uncompressed chunky strip, little-endian. */
function uncompressedTiff(options: UncompressedTiffOptions): Uint8Array {
  const entries: TiffIfdEntry[] = [
    { tag: 256, type: 3, values: [options.width] },
    { tag: 257, type: 3, values: [options.height] },
    { tag: 258, type: 3, values: options.bitsPerSample },
    { tag: 259, type: 3, values: [1] },
    { tag: 262, type: 3, values: [options.photometric] },
    { tag: 273, type: 4, values: [0] },
    { tag: 274, type: 3, values: [1] },
    { tag: 277, type: 3, values: [options.samplesPerPixel] },
    { tag: 278, type: 4, values: [options.height] },
    { tag: 279, type: 4, values: [options.pixels.byteLength] },
    { tag: 284, type: 3, values: [1] },
    ...(options.extraSample === undefined
      ? []
      : [{ tag: 338, type: 3 as const, values: [options.extraSample] }]),
  ];
  entries.sort((left, right) => left.tag - right.tag);

  const ifdOffset = 8;
  const ifdByteLength = 2 + entries.length * 12 + 4;
  let trailingOffset = ifdOffset + ifdByteLength;
  const indirectOffsets = new Map<number, number>();
  for (const entry of entries) {
    const byteLength = entry.values.length * (entry.type === 3 ? 2 : 4);
    if (byteLength <= 4) continue;
    indirectOffsets.set(entry.tag, trailingOffset);
    trailingOffset += byteLength;
  }
  if (trailingOffset % 2 !== 0) trailingOffset += 1;
  const pixelOffset = trailingOffset;
  const stripOffset = entries.find(({ tag }) => tag === 273);
  if (!stripOffset) throw new Error('synthetic TIFF is missing StripOffsets');
  (stripOffset.values as number[])[0] = pixelOffset;

  const output = Buffer.alloc(pixelOffset + options.pixels.byteLength);
  output.write('II', 0, 'ascii');
  output.writeUInt16LE(42, 2);
  output.writeUInt32LE(ifdOffset, 4);
  output.writeUInt16LE(entries.length, ifdOffset);

  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    const offset = ifdOffset + 2 + index * 12;
    output.writeUInt16LE(entry.tag, offset);
    output.writeUInt16LE(entry.type, offset + 2);
    output.writeUInt32LE(entry.values.length, offset + 4);
    const indirectOffset = indirectOffsets.get(entry.tag);
    if (indirectOffset !== undefined) {
      output.writeUInt32LE(indirectOffset, offset + 8);
      entry.values.forEach((value, valueIndex) => {
        if (entry.type === 3) output.writeUInt16LE(value, indirectOffset + valueIndex * 2);
        else output.writeUInt32LE(value, indirectOffset + valueIndex * 4);
      });
    } else if (entry.type === 3) {
      entry.values.forEach((value, valueIndex) => {
        output.writeUInt16LE(value, offset + 8 + valueIndex * 2);
      });
    } else {
      output.writeUInt32LE(entry.values[0], offset + 8);
    }
  }
  output.writeUInt32LE(0, ifdOffset + 2 + entries.length * 12);
  Buffer.from(options.pixels).copy(output, pixelOffset);
  return output;
}

function solidPixels(width: number, height: number, color: readonly number[]): Uint8Array {
  const output = new Uint8Array(width * height * color.length);
  for (let offset = 0; offset < output.length; offset += color.length) output.set(color, offset);
  return output;
}

function syntheticImageCases(): readonly SyntheticImageCase[] {
  const smallWidth = 8;
  const smallHeight = 8;
  return [
    {
      id: 'bilevel-tiff',
      filename: 'bilevel.tif',
      bytes: uncompressedTiff({
        width: smallWidth,
        height: smallHeight,
        bitsPerSample: [1],
        photometric: 1,
        samplesPerPixel: 1,
        pixels: new Uint8Array(smallHeight),
      }),
      expected: [0, 0, 0, 255],
    },
    {
      id: 'grayscale-tiff',
      filename: 'grayscale.tiff',
      bytes: uncompressedTiff({
        width: smallWidth,
        height: smallHeight,
        bitsPerSample: [8],
        photometric: 1,
        samplesPerPixel: 1,
        pixels: solidPixels(smallWidth, smallHeight, [96]),
      }),
      expected: [96, 96, 96, 255],
    },
    {
      id: 'rgb-tiff',
      filename: 'rgb.tif',
      bytes: uncompressedTiff({
        width: smallWidth,
        height: smallHeight,
        bitsPerSample: [8, 8, 8],
        photometric: 2,
        samplesPerPixel: 3,
        pixels: solidPixels(smallWidth, smallHeight, [220, 30, 40]),
      }),
      expected: [220, 30, 40, 255],
    },
    {
      id: 'rgba-tiff',
      filename: 'rgba.tiff',
      bytes: uncompressedTiff({
        width: smallWidth,
        height: smallHeight,
        bitsPerSample: [8, 8, 8, 8],
        photometric: 2,
        samplesPerPixel: 4,
        pixels: solidPixels(smallWidth, smallHeight, [20, 180, 70, 128]),
        extraSample: 2,
      }),
      expected: [137, 217, 162, 255],
    },
    {
      id: 'signature-sniffed-octet-stream-tiff',
      filename: 'signature.bin',
      bytes: uncompressedTiff({
        width: smallWidth,
        height: smallHeight,
        bitsPerSample: [8, 8, 8],
        photometric: 2,
        samplesPerPixel: 3,
        pixels: solidPixels(smallWidth, smallHeight, [30, 70, 220]),
      }),
      expected: [30, 70, 220, 255],
    },
    {
      id: 'large-indexed-png-1',
      filename: 'large-1.png',
      bytes: indexedPng(LARGE_RASTER_WIDTH, LARGE_RASTER_HEIGHT, [17, 34, 51]),
      expected: [17, 34, 51, 255],
    },
    {
      id: 'large-indexed-png-2',
      filename: 'large-2.png',
      bytes: indexedPng(LARGE_RASTER_WIDTH, LARGE_RASTER_HEIGHT, [61, 122, 183]),
      expected: [61, 122, 183, 255],
    },
    {
      id: 'large-indexed-png-3',
      filename: 'large-3.png',
      bytes: indexedPng(LARGE_RASTER_WIDTH, LARGE_RASTER_HEIGHT, [200, 80, 20]),
      expected: [200, 80, 20, 255],
    },
  ];
}

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WORD_DRAWING_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PICTURE_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const REL_IMAGE = `${OFFICE_REL_NS}/image`;

const encoder = new TextEncoder();

function xml(value: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${value}`);
}

function inlinePicture(relationshipId: string, index: number, name: string): string {
  const extent = 2_286_000; // 180 pt square: a small display target for every source raster.
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="${extent}" cy="${extent}"/>
    <wp:docPr id="${index + 1}" name="${name}"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic><a:graphicData uri="${PICTURE_NS}">
      <pic:pic><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="${name}"/>
        <pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="${relationshipId}"/>
          <a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extent}" cy="${extent}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r></w:p>`;
}

function syntheticImageDocx(images: readonly SyntheticImageCase[]): Uint8Array {
  const relationships = images.map((image, index) =>
    `<Relationship Id="rIdImage${index + 1}" Type="${REL_IMAGE}" Target="media/${image.filename}"/>`,
  ).join('');
  const body = images.map((image, index) =>
    `${inlinePicture(`rIdImage${index + 1}`, index, image.id)}${index === images.length - 1
      ? ''
      : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'}`,
  ).join('');

  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="tif" ContentType="image/tiff"/>
      <Default Extension="tiff" ContentType="image/tiff"/>
      <Default Extension="bin" ContentType="application/octet-stream"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`)],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rId1" Type="${OFFICE_REL_NS}/officeDocument" Target="word/document.xml"/>
    </Relationships>`)],
    ['word/_rels/document.xml.rels', xml(`<Relationships xmlns="${PACKAGE_REL_NS}">
      ${relationships}
    </Relationships>`)],
    ['word/document.xml', xml(`<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"
        xmlns:wp="${WORD_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:pic="${PICTURE_NS}">
      <w:body>${body}<w:sectPr>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"
          w:header="0" w:footer="0" w:gutter="0"/>
      </w:sectPr></w:body>
    </w:document>`)],
  ]);
  for (const image of images) parts.set(`word/media/${image.filename}`, image.bytes);
  return storeZip(parts);
}

test('synthetic TIFF classes and aggregate-large rasters survive DocxScrollViewer main/worker traversal', async ({ page }) => {
  test.setTimeout(180_000);
  const images = syntheticImageCases();
  const docx = syntheticImageDocx(images);

  expect(LARGE_RASTER_WIDTH * LARGE_RASTER_HEIGHT).toBeLessThan(32_000_000);
  expect(LARGE_RASTER_NATURAL_RGBA_BYTES).toBeGreaterThan(128 * 1024 * 1024);
  expect(docx.byteLength).toBeLessThan(250_000);

  await page.goto('/tests/visual/image-decode-scroll-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready');
  const report = await page.evaluate(async ({ base64, expectations }) =>
    (window as unknown as {
      runSyntheticImageAcceptance: (
        encodedDocx: string,
        expected: readonly { id: string; rgba: Rgba }[],
      ) => Promise<BrowserAcceptanceReport>;
    }).runSyntheticImageAcceptance(base64, expectations), {
    base64: Buffer.from(docx).toString('base64'),
    expectations: images.map(({ id, expected }) => ({ id, rgba: expected })),
  });

  expect(report.main.mode).toBe('main');
  expect(report.worker.mode).toBe('worker');

  for (const mode of [report.main, report.worker]) {
    expect(mode.pageCount).toBe(images.length);
    expect(mode.layoutComplete).toBe(true);
    expect(mode.errors).toEqual([]);
    expect(mode.pages.map(({ id, index }) => ({ id, index }))).toEqual(
      images.map(({ id }, index) => ({ id, index })),
    );
    for (const rendered of mode.pages) {
      expect(rendered.mountedPages).toContain(rendered.index);
      expect(rendered.maxChannelDelta).toBeLessThanOrEqual(8);
      expect(rendered.canvasWidth).toBeGreaterThan(0);
      expect(rendered.canvasWidth).toBeLessThanOrEqual(700);
      expect(rendered.canvasHeight).toBeGreaterThan(0);
      expect(rendered.canvasHeight).toBeLessThanOrEqual(900);
    }
  }

  expect(report.worker.pages.map(({ canvasWidth, canvasHeight }) => ({ canvasWidth, canvasHeight })))
    .toEqual(report.main.pages.map(({ canvasWidth, canvasHeight }) => ({ canvasWidth, canvasHeight })));
  for (let index = 0; index < images.length; index += 1) {
    const mainPixel = report.main.pages[index].pixel;
    const workerPixel = report.worker.pages[index].pixel;
    for (let channel = 0; channel < 4; channel += 1) {
      expect(Math.abs(workerPixel[channel] - mainPixel[channel])).toBeLessThanOrEqual(1);
    }
  }
});
