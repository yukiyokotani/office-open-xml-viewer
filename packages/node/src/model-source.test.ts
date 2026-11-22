import { readFile } from 'node:fs/promises';
import { Canvas, loadImage } from 'skia-canvas';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { ModelSource, ModelSourceConfig, ModelSourceTarget } from '@silurus/ooxml-core';
import { computeMdw } from '../../xlsx/src/renderer.ts';
import { openDocxDocument } from './docx.ts';
import { openPptxPresentation } from './pptx.ts';
import { isParsedPptxSlide } from './pptx-parsed-slides.ts';
import type { NodeCanvasFactory } from './render.ts';
import { minimalDocx } from './test-ooxml-package.ts';
import { openXlsxWorkbook } from './xlsx.ts';

// End-to-end through the Node facade: an application-supplied ModelSource whose
// module (test-fixtures/*-model-source.mjs) opens the input with the REAL parser
// archive, so the generic contract is exercised against real cursors, bootstrap
// and pagination rather than mocks.

const factory: NodeCanvasFactory = {
  createCanvas: (width, height) =>
    new Canvas(width, height) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buffer: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buffer as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const fixtureUrl = (name: string) => new URL(`./test-fixtures/${name}`, import.meta.url).href;

/**
 * A calling-realm source that claims by an explicit flag (never by content) and
 * hands its module a probe buffer: Float64Array [close calls, configure calls, last width].
 */
function fakeSource<T extends ModelSourceTarget>(
  target: T,
  config: ModelSourceConfig = {},
  claim: (bytes: Uint8Array) => boolean = () => true,
) {
  const probe = new ArrayBuffer(24);
  const release = vi.fn();
  const source: ModelSource<T> = {
    target,
    claim: vi.fn(claim),
    beginLoad: () => ({
      module: {
        protocol: 'ooxml-model-source-module/v2',
        target,
        moduleUrl: fixtureUrl(`${target}-model-source.mjs`),
        config,
      },
      transfer: [probe],
      release,
    }),
  };
  const counters = new Float64Array(probe);
  return {
    source,
    release,
    closes: () => counters[0],
    configures: () => counters[1],
    configuredWidth: () => counters[2],
  };
}

/** Final view: one short paragraph. Markup view: a long deletion spanning pages. */
const trackedDocx = minimalDocx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Kept text.</w:t></w:r>
      <w:ins w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:r><w:t xml:space="preserve"> Inserted.</w:t></w:r></w:ins>
      <w:del w:id="2" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> ${'Deleted words fill several pages in the markup view. '.repeat(900)}</w:delText></w:r></w:del>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`);

let docxSample: Buffer;
let xlsxSample: Buffer;
let pptxSample: Buffer;

beforeAll(async () => {
  [docxSample, xlsxSample, pptxSample] = await Promise.all([
    readFile(new URL('../../docx/public/demo/sample-1.docx', import.meta.url)),
    readFile(new URL('../../xlsx/public/demo/sample-1.xlsx', import.meta.url)),
    readFile(new URL('../../pptx/public/demo/sample-1.pptx', import.meta.url)),
  ]);
});

describe('Node sessions with a model source', () => {
  it.each(['docx', 'xlsx', 'pptx'] as const)('closes an opened %s archive when release throws', async (target) => {
    const probe = fakeSource(target);
    const failure = new Error('release failed');
    probe.release.mockImplementation(() => { throw failure; });
    const load = target === 'docx'
      ? openDocxDocument(docxSample, { factory, legacyDoc: probe.source as ModelSource<'docx'> })
      : target === 'xlsx'
        ? openXlsxWorkbook(xlsxSample, { legacyXls: probe.source as ModelSource<'xlsx'> })
        : openPptxPresentation(pptxSample, { legacyPpt: probe.source as ModelSource<'pptx'> });
    await expect(load).rejects.toBe(failure);
    expect(probe.closes()).toBe(1);
    expect(probe.release).toHaveBeenCalledOnce();
  });
  it('paginates the markup view when the DOCX source reports it as its view default', async () => {
    const ooxml = await openDocxDocument(trackedDocx, { factory, currentDate: 0 });
    const finalPages = ooxml.pageCount;
    await ooxml.close();

    const markup = fakeSource('docx', { showTrackedChanges: true });
    const followed = await openDocxDocument(trackedDocx, {
      factory, currentDate: 0, legacyDoc: markup.source,
    });
    expect(markup.release).toHaveBeenCalledOnce();
    expect(followed.pageCount).toBeGreaterThan(finalPages);
    await expect(followed.renderPage(followed.pageCount - 1)).resolves.toBeDefined();
    await followed.close();
    expect(markup.closes()).toBe(1);

    const plain = fakeSource('docx', { showTrackedChanges: null });
    const unchosen = await openDocxDocument(trackedDocx, {
      factory, currentDate: 0, legacyDoc: plain.source,
    });
    expect(unchosen.pageCount).toBe(finalPages);
    await unchosen.close();
  }, 30_000);

  it('fails closed and closes the archive on an unsupported DOCX view default', async () => {
    const invalid = fakeSource('docx', { showTrackedChanges: true, extraViewDefault: 'showComments' });
    await expect(openDocxDocument(docxSample, { factory, legacyDoc: invalid.source }))
      .rejects.toThrow('unsupported DOCX model source view default: showComments');
    expect(invalid.closes()).toBe(1);
    expect(invalid.release).toHaveBeenCalledOnce();
  });

  it('streams and renders a DOCX source with only the required archive methods', async () => {
    const minimal = fakeSource('docx', { minimal: true });
    const session = await openDocxDocument(docxSample, { factory, currentDate: 0, legacyDoc: minimal.source });
    expect(session.resourceUsage).toBeUndefined();
    expect(session.pageCount).toBeGreaterThan(0);
    for await (const page of session.pages({ dpr: 1 })) {
      expect(page.canvas.width).toBeGreaterThan(0);
      break;
    }
    await session.close();
    expect(minimal.closes()).toBe(1);
  }, 30_000);

  it('configures XLSX host layout with the renderer measurement through the factory canvas', async () => {
    const measured = fakeSource('xlsx', { hostLayoutFamily: 'Calibri', hostLayoutSizePt: 11 });
    const session = await openXlsxWorkbook(xlsxSample, { factory, legacyXls: measured.source });
    const context = factory.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D;
    const expected = computeMdw('Calibri', 11, undefined, undefined, 400, 'normal', context);
    expect(expected).toBeGreaterThan(0);
    expect(measured.configures()).toBe(1);
    expect(measured.configuredWidth()).toBe(expected);
    expect(session.workbookIndex.layoutMetrics?.maximumDigitWidth).toBe(expected);
    expect(session.resourceUsage).toBeUndefined();
    let finished = false;
    for await (const chunk of session.worksheetRows(0)) {
      if (chunk.kind === 'finished') finished = true;
    }
    expect(finished).toBe(true);
    await session.close();
    expect(measured.closes()).toBe(1);

    const unmeasured = fakeSource('xlsx', { hostLayoutFamily: 'Calibri', hostLayoutSizePt: 11 });
    const withoutFactory = await openXlsxWorkbook(xlsxSample, { legacyXls: unmeasured.source });
    expect(unmeasured.configures()).toBe(1);
    expect(unmeasured.configuredWidth()).toBeNaN();
    expect(withoutFactory.workbookIndex.layoutMetrics).toBeUndefined();
    await withoutFactory.close();
  }, 30_000);

  it('streams PPTX slides without optional capabilities and rejects media reads', async () => {
    const minimal = fakeSource('pptx', { minimal: true });
    const session = await openPptxPresentation(pptxSample, { legacyPpt: minimal.source });
    expect(session.slideCount).toBeGreaterThan(0);
    expect(session.resourceUsage).toBeUndefined();
    await expect(session.getMedia('ppt/media/media1.mp4'))
      .rejects.toThrow('media extraction is unsupported for this source');
    let slides = 0;
    for await (const _slide of session.slides()) slides += 1;
    expect(slides).toBe(session.slideCount);
    await session.close();
    expect(minimal.closes()).toBe(1);
  }, 30_000);

  it('marks only built-in parser slides as a parsed text-colour chain', async () => {
    // Issue #1989's hard black default needs the parser's resolved colour
    // chain; model-source slides keep the 0.89.7 fallback.
    const marks = async (session: Awaited<ReturnType<typeof openPptxPresentation>>) => {
      const result: boolean[] = [];
      for await (const slide of session.slides()) result.push(isParsedPptxSlide(slide));
      return result;
    };
    const parsed = await marks(await openPptxPresentation(pptxSample));
    expect(parsed.length).toBeGreaterThan(0);
    expect(parsed.every(Boolean)).toBe(true);
    const sourced = await marks(await openPptxPresentation(pptxSample, { legacyPpt: fakeSource('pptx').source }));
    expect(sourced.length).toBeGreaterThan(0);
    expect(sourced.some(Boolean)).toBe(false);
    // A source that declines the input leaves it on the built-in parser.
    const declined = fakeSource('pptx', {}, () => false);
    const fallback = await marks(await openPptxPresentation(pptxSample, { legacyPpt: declined.source }));
    expect(declined.source.claim).toHaveBeenCalled();
    expect(fallback.length).toBeGreaterThan(0);
    expect(fallback.every(Boolean)).toBe(true);
  }, 30_000);

  it('keeps unclaimed input on the OOXML path and fails closed on invalid sources', async () => {
    const unclaimed = fakeSource('pptx', {}, () => false);
    const ooxml = await openPptxPresentation(pptxSample, { legacyPpt: unclaimed.source });
    expect(unclaimed.source.claim).toHaveBeenCalledOnce();
    expect(unclaimed.release).not.toHaveBeenCalled();
    expect(ooxml.resourceUsage).toBeDefined();
    await ooxml.close();

    const mismatched = fakeSource('docx');
    // A source for another format is a runtime TypeError, not only a type error.
    await expect(openXlsxWorkbook(xlsxSample, { legacyXls: mismatched.source as unknown as ModelSource<'xlsx'> }))
      .rejects.toThrow(TypeError);
    expect(mismatched.source.claim).not.toHaveBeenCalled();

    const failure = new Error('claim failed');
    const throwing = fakeSource('xlsx', {}, () => { throw failure; });
    await expect(openXlsxWorkbook(xlsxSample, { legacyXls: throwing.source })).rejects.toBe(failure);
    expect(throwing.release).not.toHaveBeenCalled();
  });
});
