// Synthetic binary PowerPoint record builders for the direct-reader Node
// tests (MS-PPT record headers, MS-ODRAW OfficeArt containers). No private data.
import { buildPptFixture, concat, little16, little32, utf16le } from '../test-fixtures.js';
import { testPptSource } from '../test-sources.js';
import { materializePptxPresentation, openPptxPresentation } from './node-facade.js';

export { buildPptFixture, concat, little16, little32, utf16le };

/** One record: version/instance word, type, length, payload. */
export const record = (options: number, kind: number, payload: Uint8Array): Uint8Array =>
  concat(little16(options), little16(kind), little32(payload.length), payload);

/** OfficeArtFOPT with simple (id, value) properties and an optional complex tail. */
export const properties = (entries: readonly (readonly [number, number])[], tail: Uint8Array = new Uint8Array()): Uint8Array =>
  record((entries.length << 4) | 3, 0xf00b, concat(...entries.map(([id, value]) => concat(little16(id), little32(value))), tail));

/** OfficeArtFSP: shape type, id and flags. */
export const shapeAtom = (kind: number, id: number, flags: number): Uint8Array =>
  record((kind << 4) | 2, 0xf00a, concat(little32(id), little32(flags)));

/** ClientAnchor in master units (1/576 inch), in PPT top, left, right, bottom order. */
export const anchor = (top: number, left: number, right: number, bottom: number): Uint8Array =>
  record(0, 0xf010, concat(little32(top), little32(left), little32(right), little32(bottom)));

/**
 * TextRulerAtom giving level 0 explicit zero margin and indent origins, so a
 * text box without document or master paragraph styles has its paragraph
 * geometry authored (the direct reader never invents it).
 */
export const zeroOriginRuler = (): Uint8Array => record(0, 4006, concat(little32(8 | 256), little32(0)));

export const spContainer = (...parts: Uint8Array[]): Uint8Array => record(15, 0xf004, concat(...parts));

/** A PPDrawing (1036) holding one OfficeArtDgContainer with the given shapes. */
export const drawing = (...shapes: Uint8Array[]): Uint8Array => record(15, 1036, record(15, 0xf002, concat(...shapes)));

/** SlideAtom (1007) with a master persist id and the fFollowMaster* flags. */
export const slideAtom = (master: number, flags: number): Uint8Array =>
  record(2, 1007, concat(new Uint8Array(12), little32(master), little32(0), little16(flags), little16(0)));

export const pptOptions = () => ({ modelSources: [testPptSource()] });

export const materialize = (bytes: Uint8Array) => materializePptxPresentation(bytes, pptOptions());

export const openSession = (bytes: Uint8Array) => openPptxPresentation(bytes, pptOptions());
