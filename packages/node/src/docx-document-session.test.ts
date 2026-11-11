import { readFile } from 'node:fs/promises';
import { Canvas, loadImage } from 'skia-canvas';
import { beforeAll, describe, expect, it, vi } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
import { createLayoutServices } from '../../docx/src/layout-runtime.ts';
import { layoutDocument } from '../../docx/src/document-layout.ts';
import {
  materializeDocxDocument,
  openDocxDocument,
} from './docx.ts';
import type { NodeCanvasFactory } from './render.ts';
import { OoxmlError } from '@silurus/ooxml-core';
import { nonOoxmlInputs } from '@silurus/ooxml-core/testing';

const factory: NodeCanvasFactory = {
  createCanvas: (width, height) =>
    new Canvas(width, height) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buffer: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buffer as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

let bytes: Buffer;

beforeAll(async () => {
  bytes = await readFile(new URL('../../docx/public/demo/sample-1.docx', import.meta.url));
});

describe('Node bounded DOCX document session', () => {
  it('materializes the compatibility model through the bounded document producer', async () => {
    const parse = vi.spyOn(docxWasm, 'parse_docx');
    try {
      const document = await materializeDocxDocument(bytes);
      expect(document.body.length).toBeGreaterThan(0);
      expect(document.section).toBeDefined();
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it.each(nonOoxmlInputs('docx'))('rejects %s with OoxmlError not-ooxml', async (_name, input) => {
    for (const open of [
      () => materializeDocxDocument(input),
      () => openDocxDocument(input, { factory }),
    ]) {
      const error = await open().then(
        (document) => { throw new Error(`resolved ${JSON.stringify(document).slice(0, 80)}`); },
        (rejection: unknown) => rejection,
      );
      expect(error).toBeInstanceOf(OoxmlError);
      expect(error).toMatchObject({ code: 'not-ooxml' });
    }
  });

  it('materializes a package with a local-only ZIP timestamp mismatch', async () => {
    const mismatched = Buffer.from(bytes);
    const local = mismatched.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]));
    expect(local).toBeGreaterThanOrEqual(0);
    mismatched.writeUInt16LE(mismatched.readUInt16LE(local + 10) ^ 1, local + 10);
    mismatched.writeUInt16LE(mismatched.readUInt16LE(local + 12) ^ 1, local + 12);

    const document = await materializeDocxDocument(mismatched);
    expect(document.body.length).toBeGreaterThan(0);
  });

  it('matches compatibility pagination and renders one caller-owned canvas at a time', async () => {
    const expected = await materializeDocxDocument(bytes);
    const measure = factory.createCanvas(1, 1).getContext('2d');
    const expectedPages = layoutDocument(
      expected,
      createLayoutServices(expected, {
        measureContext: measure as unknown as CanvasRenderingContext2D,
      }),
      { currentDateMs: 0 },
    ).pages.length;

    const session = await openDocxDocument(bytes, { factory, currentDate: 0 });
    expect(session.pageCount).toBe(expectedPages);
    expect(session.resourceUsage?.operationInflatedBytes).toBeGreaterThan(0);

    let count = 0;
    for await (const page of session) {
      expect(page.pageIndex).toBe(count);
      expect(page.widthPt).toBeGreaterThan(0);
      expect(page.heightPt).toBeGreaterThan(0);
      expect(page.canvas.width).toBeGreaterThan(0);
      expect(page.canvas.height).toBeGreaterThan(0);
      count += 1;
    }
    expect(count).toBe(expectedPages);
    // This compares two parser/layout paths and paints every public demo page
    // through native Skia. Keep the same bounded integration timeout as the
    // multi-page lifetime test below; the assertions check pagination and
    // ownership, rather than a machine-dependent five-second performance SLO.
  }, 15_000);

  it('does not route the bounded session through the materializing parse export', async () => {
    const parse = vi.spyOn(docxWasm, 'parse_docx');
    try {
      const session = await openDocxDocument(bytes, { factory, currentDate: 0 });
      await session.close();
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it('extracts model pictures through the session-owned bounded cache and rejects after close', async () => {
    const extract = vi.spyOn(archivePrototype(), 'extract_image');
    try {
      const session = await openDocxDocument(bytes, { factory, currentDate: 0 });
      const first = await session.getImage('word/media/image1.jpeg', 'image/jpeg');
      const second = await session.getImage('word/media/image1.jpeg', 'image/jpeg');
      expect(first).toBe(second);
      expect(first.type).toBe('image/jpeg');
      // The demo package stores image1.jpeg as a 202,076-byte JPEG (SOI FF D8).
      expect(first.size).toBe(202_076);
      expect([...new Uint8Array(await first.slice(0, 2).arrayBuffer())]).toEqual([0xff, 0xd8]);
      expect((await session.getImage('word/media/image1.jpeg')).type).toBe('application/octet-stream');
      expect(extract).toHaveBeenCalledTimes(1);
      expect(session.resourceUsage?.distinctInflatedBytes).toBeGreaterThanOrEqual(202_076);

      await session.close();
      await expect(session.getImage('word/media/image1.jpeg', 'image/jpeg')).rejects.toThrow(/closed/);
      expect(extract).toHaveBeenCalledTimes(1);
    } finally {
      extract.mockRestore();
    }
  });

  it('fails the session when a picture read exceeds the archive entry budget', async () => {
    // document.xml (68,517 bytes) and every other non-media part fit; the
    // 202,076-byte image1.jpeg does not.
    const session = await openDocxDocument(bytes, {
      factory,
      currentDate: 0,
      resourceLimits: { maxArchiveEntryBytes: 150_000 },
    });
    try {
      const failure = await session.getImage('word/media/image1.jpeg', 'image/jpeg').then(
        () => { throw new Error('resolved an over-budget picture'); },
        (error: unknown) => error,
      );
      expect(failure).toMatchObject({ name: 'OoxmlResourceLimitError', code: 'ooxml-resource-limit' });
      // A resource-limit failure fails the session, as in the PPTX session.
      await expect(session.getImage('word/media/image3.jpeg', 'image/jpeg')).rejects.toBe(failure);
      await expect(session.renderPage(0)).rejects.toBe(failure);
    } finally {
      await session.close();
    }
  });

  it('fans a render-time trap out to sibling sessions and recovers on the next open', async () => {
    const extract = vi.spyOn(archivePrototype(), 'extract_image')
      .mockImplementationOnce(() => { throw new RangeError('synthetic trap'); });
    const first = await openDocxDocument(bytes, { factory, currentDate: 0 });
    const sibling = await openDocxDocument(bytes, { factory, currentDate: 0 });
    try {
      await expect(first.renderPage(0)).rejects.toMatchObject({ code: 'parser-crashed' });
      await expect(sibling.renderPage(0)).rejects.toMatchObject({ code: 'parser-crashed' });
      await expect(first.close()).resolves.toBeUndefined();
      await expect(sibling.close()).resolves.toBeUndefined();
    } finally {
      extract.mockRestore();
    }

    const recovered = await openDocxDocument(bytes, { factory, currentDate: 0 });
    await expect(recovered.renderPage(0)).resolves.toMatchObject({ width: expect.any(Number) });
    await recovered.close();
  }, 15_000);

  it('frees exactly once after completion, early return, and explicit close', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const completed = await openDocxDocument(bytes, { factory, currentDate: 0 });
      for await (const _page of completed.pages({ dpr: 1 })) { /* drain */ }
      await completed.close();
      expect(free).toHaveBeenCalledTimes(1);

      const early = await openDocxDocument(bytes, { factory, currentDate: 0 });
      for await (const _page of early.pages({ dpr: 1 })) break;
      await early.close();
      expect(free).toHaveBeenCalledTimes(2);

      const unopened = await openDocxDocument(bytes, { factory, currentDate: 0 });
      await unopened.close();
      await unopened.close();
      expect(free).toHaveBeenCalledTimes(3);
    } finally {
      free.mockRestore();
    }
  }, 15_000);

  it('reports metrics for the owned session when it closes', async () => {
    const onResourceMetrics = vi.fn();
    const session = await openDocxDocument(bytes, {
      factory,
      currentDate: 0,
      onResourceMetrics,
    });
    expect(onResourceMetrics).not.toHaveBeenCalled();
    await session.close();
    expect(onResourceMetrics).toHaveBeenCalledOnce();
    expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
      format: 'docx',
      scope: 'session',
      status: 'ok',
    }));
  });

  it('normalizes package limits and aborts between page pulls', async () => {
    await expect(openDocxDocument(bytes, {
      factory,
      resourceLimits: { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: null },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
    });
    await expect(openDocxDocument(bytes, {
      factory,
      resourceLimits: { maxArchiveEntries: 1 },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
      details: expect.objectContaining({
        violation: expect.objectContaining({
          metric: 'entry-count',
          configurable: true,
          limit: 1,
        }),
      }),
    });

    const before = new AbortController();
    before.abort();
    await expect(openDocxDocument(bytes, { factory, signal: before.signal }))
      .rejects.toMatchObject({ name: 'AbortError' });

    const after = new AbortController();
    const session = await openDocxDocument(bytes, {
      factory,
      currentDate: 0,
      signal: after.signal,
    });
    const iterator = session.pages({ dpr: 1 });
    expect((await iterator.next()).value?.pageIndex).toBe(0);
    after.abort();
    await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  });
});

type ArchivePrototype = {
  free(): void;
  extract_image(path: string): Uint8Array;
};

function archivePrototype(): ArchivePrototype {
  return (docxWasm as unknown as { DocxArchive: { prototype: ArchivePrototype } })
    .DocxArchive.prototype;
}
