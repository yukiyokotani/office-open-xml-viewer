// Windows metafile pictures in OpenDocument text and spreadsheets reach the
// shared renderer under the media types it keys metafile handling on
// (core `isMetafileMime`): a clipped metafile is cropped against its frame,
// and its text records make the font demand cover every scalar.
import { describe, expect, it } from 'vitest';
import { metafileRasterSize } from '@silurus/ooxml-core';
import { buildOdtFixture } from '../odt-fixture.js';
import { buildOdsFixture } from '../ods-fixture.js';
import { testOdsSource, testOdtSource } from '../test-sources.js';
import { materializeDocxDocument } from './node-facade.js';
import { materializeXlsxWorkbook } from './ods-facade.js';
import { docxFontDemandScalars } from '../../../docx/src/font-source-load.ts';
import { XlsxFontDemand } from '../../../xlsx/src/font-source-load.ts';
import type { DocxDocumentModel } from '../../../docx/src/types.ts';
import type { ParsedWorkbook, Worksheet } from '../../../xlsx/src/types.ts';

/** A minimal EMF (header and EOF record) whose frame (`rclFrame`, 0.01 mm)
 * is 25.60 × 12.80 mm, as a string of code units below 0x80 (the fixture
 * zip stores UTF-8). */
function emf(): string {
  const words = [
    1, 88, 0, 0, 49, 24, // EMR_HEADER, size, rclBounds (device pixels)
    0, 0, 2560, 1280, // rclFrame
    0x464d4520, 0x10000, 108, 2, 0, 0, 0, 0, // " EMF", version, bytes, records, handles, ...
    100, 100, 64, 64, // szlDevice, szlMillimeters
    14, 20, 0, 16, 20, // EMR_EOF
  ];
  const bytes = new Uint8Array(new Uint32Array(words).buffer);
  expect(bytes.every((b) => b < 0x80)).toBe(true);
  return String.fromCharCode(...bytes);
}

// The picture has no file extension: only its media type marks it a metafile.
const PICTURE = 'Pictures/1000000000000';
// A quarter of the frame's 25.6 mm width clipped from the right.
const CLIP = 'rect(0mm, 6.4mm, 0mm, 0mm)';
const DRAW = 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"';

type Picture = { mimeType?: string; imagePath?: string; srcRect?: { l: number; t: number; r: number; b: number } };

function pictures(value: unknown, found: Picture[] = []): Picture[] {
  if (value && typeof value === 'object') {
    const record = value as Picture;
    if (record.imagePath === PICTURE) found.push(record);
    for (const child of Object.values(value)) pictures(child, found);
  }
  return found;
}

describe('OpenDocument metafile pictures', () => {
  it('ODT: crops against the metafile frame and demands every scalar', async () => {
    const model = await materializeDocxDocument(buildOdtFixture({
      automaticStyles: `<style:style style:name="fr1" style:family="graphic"><style:graphic-properties fo:clip="${CLIP}"/></style:style>`,
      body: `<text:p><draw:frame ${DRAW} draw:style-name="fr1" text:anchor-type="as-char" svg:width="3in" svg:height="1in"><draw:image xlink:href="${PICTURE}"/></draw:frame>x</text:p>`,
      extraFiles: [[PICTURE, emf()]],
    }), { odt: testOdtSource() });
    const [picture] = pictures(model);
    expect(picture?.mimeType).toBe('image/emf');
    expect(picture?.srcRect?.r).toBeCloseTo(0.25, 9);
    // The shared painter rasterizes the whole frame and crops it.
    const size = metafileRasterSize(picture!.mimeType!, picture!.srcRect as never, 216, 72)!;
    expect(size.widthPt).toBeCloseTo(288, 6);
    expect(docxFontDemandScalars(model as unknown as DocxDocumentModel)).toBe('all');
  });

  it('ODS: crops against the metafile frame and demands every scalar', async () => {
    const workbook = await materializeXlsxWorkbook(buildOdsFixture({
      automaticStyles: `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties fo:clip="${CLIP}"/></style:style>`,
      body: '<table:table table:name="S"><table:table-row><table:table-cell>'
        + `<draw:frame draw:style-name="gr1" svg:width="3in" svg:height="1in" svg:x="0cm" svg:y="0cm"><draw:image xlink:href="${PICTURE}"/></draw:frame>`
        + '</table:table-cell></table:table-row></table:table>',
      extraFiles: [[PICTURE, emf()]],
    }), { ods: testOdsSource() });
    const sheet = workbook.worksheets[0] as unknown as Worksheet;
    const [picture] = pictures(sheet);
    expect(picture?.mimeType).toBe('image/emf');
    expect(picture?.srcRect?.r).toBeCloseTo(0.25, 9);
    const demand = new XlsxFontDemand(workbook.workbookIndex as unknown as ParsedWorkbook, 'jp', []);
    demand.take();
    demand.worksheetParts(sheet);
    expect(demand.take()?.scalars).toBe('all');
  });
});
