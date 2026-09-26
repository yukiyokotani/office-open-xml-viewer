import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

/**
 * SC18 contract: the wasm-bindgen glue for `{Pptx,Docx,Xlsx}Archive.extract_*`
 * returns an INDEPENDENT `Uint8Array` — the glue does
 * `getArrayU8FromWasm0(ptr, len).slice()` then `__wbindgen_free`s the Rust Vec,
 * so the returned array (a) no longer aliases WASM linear memory and (b) is
 * full-span over its own `ArrayBuffer` (byteOffset 0, byteLength === buffer
 * length). Those two facts are exactly what let the three `worker.ts` files
 * transfer `bytes.buffer` DIRECTLY instead of re-copying via
 * `new Uint8Array(bytes).slice().buffer`. This test pins the contract so a
 * future wasm-bindgen upgrade that returned a memory VIEW (which would make a
 * direct transfer unsafe / throw) fails here loudly.
 */

const MAIN_PART: Record<string, string> = {
  ppt: 'ppt/presentation.xml',
  word: 'word/document.xml',
  xl: 'xl/workbook.xml',
};

// Minimal STORED OPC package: the archive constructors admit only a ZIP with
// `[Content_Types].xml` and the format's main part, so those two placeholder
// entries precede the entry under test.
function makeZipWithEntry(name: string, data: Uint8Array): Uint8Array {
  const enc = new TextEncoder();
  const entries: Array<[string, Uint8Array]> = [
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')],
    [MAIN_PART[name.split('/')[0]], enc.encode('<root/>')],
    [name, data],
  ];
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const [entryName, body] of entries) {
    const nameBytes = enc.encode(entryName);
    const crc = crc32(Buffer.from(body)) >>> 0;
    const size = body.length;

    const local = new Uint8Array(30 + nameBytes.length + size);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(8, 0, true); // stored
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    local.set(body, 30 + nameBytes.length);

    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);

    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }
  const centralLength = centrals.reduce((sum, part) => sum + part.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralLength, true);
  ev.setUint32(16, offset, true);

  const out = new Uint8Array(offset + centralLength + eocd.length);
  let cursor = 0;
  for (const part of [...locals, ...centrals, eocd]) {
    out.set(part, cursor);
    cursor += part.length;
  }
  return out;
}

interface ArchiveHandle {
  extract_image(path: string): Uint8Array;
  free(): void;
}
interface PptxHandle extends ArchiveHandle {
  extract_media(path: string): Uint8Array;
}

// The WASM parser bindings are git-ignored build output. CI runs `pnpm build:wasm`
// before `pnpm test`, so they are present there; but a local run before a wasm
// build (or a stripped environment) may lack them. Probe once via a top-level-await
// IIFE and gate the whole suite with `describe.skipIf` so it reports as SKIPPED
// (not a silent zero-assertion pass) when the wasm is unavailable — mirroring the
// sibling `source-buffer-image.test.ts`.
const wasmReady = await (async () => {
  try {
    loadWasmModule(
      pptxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../pptx/src/wasm/pptx_parser_bg.wasm'),
    );
    loadWasmModule(
      docxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../docx/src/wasm/docx_parser_bg.wasm'),
    );
    loadWasmModule(
      xlsxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../xlsx/src/wasm/xlsx_parser_bg.wasm'),
    );
    return true;
  } catch {
    return false;
  }
})();

/** Assert the returned array is byte-correct AND a transfer-safe, full-span,
 *  non-WASM-aliasing buffer (the SC18 precondition for a direct transfer). */
function assertIndependentFullSpan(bytes: Uint8Array, expected: Uint8Array) {
  // Byte-correct: the copy the glue made preserves the entry contents.
  expect(Array.from(bytes)).toEqual(Array.from(expected));
  // Full-span: a bare `.buffer` transfer carries exactly these bytes, nothing
  // else (no shared arena, no leading/trailing slop from a subarray view).
  expect(bytes.byteOffset).toBe(0);
  expect(bytes.byteLength).toBe(bytes.buffer.byteLength);
  // Independent of WASM linear memory: an ArrayBuffer is transferable only if it
  // is not WASM-backed. structuredClone with transfer succeeds here and detaches
  // the buffer — it would throw for a WASM-memory view.
  const buf = bytes.buffer;
  expect(() => structuredClone(buf, { transfer: [buf] })).not.toThrow();
  expect(buf.byteLength).toBe(0); // detached by the transfer → was standalone
}

describe.skipIf(!wasmReady)('SC18 extract_* returns a transfer-safe buffer', () => {
  it('pptx: extract_image + extract_media are independent full-span copies', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5, 6]);
    const mp4 = new Uint8Array([0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70]);
    // Each path lives in its own package (the builder carries one entry under test).
    const zip = makeZipWithEntry('ppt/media/image1.png', png);
    const Handle = (pptxWasm as unknown as { PptxArchive: new (b: Uint8Array) => PptxHandle })
      .PptxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('ppt/media/image1.png'), png);
    } finally {
      ar.free();
    }

    const zip2 = makeZipWithEntry('ppt/media/media2.mp4', mp4);
    const ar2 = new Handle(zip2);
    try {
      assertIndependentFullSpan(ar2.extract_media('ppt/media/media2.mp4'), mp4);
    } finally {
      ar2.free();
    }
  });

  it('docx: extract_image is an independent full-span copy', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]);
    const zip = makeZipWithEntry('word/media/image1.png', png);
    const Handle = (docxWasm as unknown as { DocxArchive: new (b: Uint8Array) => ArchiveHandle })
      .DocxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('word/media/image1.png'), png);
    } finally {
      ar.free();
    }
  });

  it('xlsx: extract_image is an independent full-span copy', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 42, 43]);
    const zip = makeZipWithEntry('xl/media/image1.png', png);
    const Handle = (xlsxWasm as unknown as { XlsxArchive: new (b: Uint8Array) => ArchiveHandle })
      .XlsxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('xl/media/image1.png'), png);
    } finally {
      ar.free();
    }
  });
});
