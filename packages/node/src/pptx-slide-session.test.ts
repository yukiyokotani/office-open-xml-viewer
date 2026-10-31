import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
import {
  materializePptxPresentation,
  openPptxPresentation,
} from './pptx.ts';
import { OoxmlDecodedImageLimitError, OoxmlError } from '@silurus/ooxml-core';
import { nonOoxmlInputs } from '@silurus/ooxml-core/testing';
import type { NodeCanvasFactory, NodeCanvasLike } from './render.ts';
import { storedZip } from './test-ooxml-package.ts';

let bytes: Buffer;

beforeAll(async () => {
  bytes = await readFile(new URL('../../pptx/public/demo/sample-1.pptx', import.meta.url));
});

describe('Node bounded PPTX presentation session', () => {
  it('materializes a complete presentation through the acknowledged slide producer', async () => {
    const parse = vi.spyOn(pptxWasm, 'parse_pptx');
    try {
      const presentation = await materializePptxPresentation(bytes);
      expect(presentation.slides.length).toBeGreaterThan(0);
      expect(presentation.slideWidth).toBeGreaterThan(0);
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it.each(nonOoxmlInputs('pptx'))('rejects %s with OoxmlError not-ooxml', async (_name, input) => {
    for (const open of [() => openPptxPresentation(input), () => materializePptxPresentation(input)]) {
      const error = await open().then(
        () => { throw new Error('resolved a non-OOXML presentation'); },
        (rejection: unknown) => rejection,
      );
      expect(error).toBeInstanceOf(OoxmlError);
      expect(error).toMatchObject({ code: 'not-ooxml' });
    }
  });

  it('matches the materializing compatibility model one canonical slide at a time', async () => {
    const expected = await materializePptxPresentation(bytes);
    const session = await openPptxPresentation(bytes);
    expect(session.slideCount).toBe(expected.slides.length);
    expect(session.slideWidth).toBe(expected.slideWidth);
    expect(session.slideHeight).toBe(expected.slideHeight);

    let index = 0;
    for await (const slide of session) {
      expect(slide).toEqual(expected.slides[index]);
      expect(session.resourceUsage?.operationInflatedBytes).toBeGreaterThan(0);
      index += 1;
    }
    expect(index).toBe(expected.slides.length);
  });

  it('does not route a bounded drain through the materializing parse export', async () => {
    const parse = vi.spyOn(pptxWasm, 'parse_pptx');
    try {
      const session = await openPptxPresentation(bytes);
      for await (const _slide of session) { /* drain */ }
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it('extracts lazy parts through the session-owned bounded cache', async () => {
    const extract = vi.spyOn(archivePrototype(), 'extract_image');
    try {
      const session = await openPptxPresentation(bytes);
      const first = await session.getImage('ppt/media/image1.jpeg', 'image/jpeg');
      const second = await session.getImage('ppt/media/image1.jpeg', 'image/jpeg');

      expect(first).toBe(second);
      expect(first.size).toBeGreaterThan(0);
      expect(extract).toHaveBeenCalledTimes(1);
      expect(session.resourceUsage?.distinctInflatedBytes).toBeGreaterThan(0);
      await session.close();
      await expect(session.getImage('ppt/media/image1.jpeg', 'image/jpeg'))
        .rejects.toThrow(/closed/);
    } finally {
      extract.mockRestore();
    }
  });

  it('keeps raw-part extraction independent while a slide cursor awaits acknowledgement', async () => {
    const prototype = archivePrototype();
    const originalPullSlide = prototype.pull_slide;
    let extractedBytes = 0;
    const pullSlide = vi.spyOn(prototype, 'pull_slide').mockImplementation(function (
      this: ArchivePrototype,
      ...args: Parameters<ArchivePrototype['pull_slide']>
    ) {
      const payload = originalPullSlide.apply(this, args);
      extractedBytes = this.extract_image('ppt/media/image1.jpeg').byteLength;
      return payload;
    });
    try {
      const session = await openPptxPresentation(bytes);
      const iterator = session.slides();
      await expect(iterator.next()).resolves.toMatchObject({ done: false });
      expect(extractedBytes).toBeGreaterThan(0);
      await iterator.return();
    } finally {
      pullSlide.mockRestore();
    }
  });

  it('invalidates sibling sessions after a trap and recovers on one fresh generation', async () => {
    const extract = vi.spyOn(archivePrototype(), 'extract_image')
      .mockImplementationOnce(() => { throw new RangeError('synthetic trap'); });
    const first = await openPptxPresentation(bytes);
    const sibling = await openPptxPresentation(bytes);
    try {
      await expect(first.getImage('ppt/media/image1.jpeg', 'image/jpeg'))
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
      await expect(sibling.getImage('ppt/media/image1.jpeg', 'image/jpeg'))
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
      await expect(first.close()).rejects.toMatchObject({ code: 'parser-crashed' });
      await expect(sibling.close()).rejects.toMatchObject({ code: 'parser-crashed' });
    } finally {
      extract.mockRestore();
    }

    const recovered = await openPptxPresentation(bytes);
    await expect(recovered.getImage('ppt/media/image1.jpeg', 'image/jpeg'))
      .resolves.toMatchObject({ size: expect.any(Number) });
    await recovered.close();
  });

  it('reuses one render byte source and lets an accepted render finish before close', async () => {
    const slide = (await materializePptxPresentation(bytes)).slides[0]!;
    const renderModule = await import('./render.ts');
    const started = deferred<void>();
    const resume = deferred<void>();
    const fetchers: unknown[] = [];
    const render = vi.spyOn(renderModule, 'renderParsedSlideNode').mockImplementation(
      async (_canvas, _presentation, _slideIndex, options) => {
        fetchers.push(options?.fetchImage);
        started.resolve();
        await resume.promise;
        await options?.fetchImage?.('ppt/media/image1.jpeg', 'image/jpeg');
      },
    );
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const session = await openPptxPresentation(bytes);
      const rendering = session.renderSlide(fakeCanvas(), slide, { factory: fakeFactory() });
      await started.promise;
      let closeSettled = false;
      const closing = session.close().finally(() => { closeSettled = true; });
      await Promise.resolve();
      expect(closeSettled).toBe(false);
      expect(free).not.toHaveBeenCalled();

      resume.resolve();
      await rendering;
      await closing;
      expect(fetchers).toHaveLength(1);
      expect(free).toHaveBeenCalledTimes(1);
    } finally {
      render.mockRestore();
      free.mockRestore();
    }
  });

  it('reports render failures as terminal session metrics', async () => {
    const renderModule = await import('./render.ts');
    const failure = new OoxmlDecodedImageLimitError('image-pixels', 10, 11);
    const render = vi.spyOn(renderModule, 'renderParsedSlideNode').mockRejectedValue(failure);
    const onResourceMetrics = vi.fn();
    try {
      const session = await openPptxPresentation(bytes, { onResourceMetrics });
      const slide = (await materializePptxPresentation(bytes)).slides[0]!;
      await expect(session.renderSlide(fakeCanvas(), slide, { factory: fakeFactory() }))
        .rejects.toBe(failure);
      await session.close();
      expect(onResourceMetrics).toHaveBeenCalledOnce();
      expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
        format: 'pptx',
        scope: 'session',
        status: 'error',
      }));
    } finally {
      render.mockRestore();
    }
  });

  it('closes and frees exactly once after completion, early return, and explicit close', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const completed = await openPptxPresentation(bytes);
      for await (const _slide of completed) { /* drain */ }
      await completed.close();
      expect(free).toHaveBeenCalledTimes(1);

      const early = await openPptxPresentation(bytes);
      for await (const _slide of early) break;
      await early.close();
      expect(free).toHaveBeenCalledTimes(2);

      const unopened = await openPptxPresentation(bytes);
      await unopened.close();
      await unopened.close();
      expect(free).toHaveBeenCalledTimes(3);
    } finally {
      free.mockRestore();
    }
  });

  it('reports a cleanup failure exactly once when archive release rejects close', async () => {
    const prototype = archivePrototype();
    const originalFree = prototype.free;
    const cleanupError = new Error('archive cleanup failed');
    const free = vi.spyOn(prototype, 'free').mockImplementationOnce(function (this: ArchivePrototype) {
      originalFree.call(this);
      throw cleanupError;
    });
    const onResourceMetrics = vi.fn();
    try {
      const session = await openPptxPresentation(bytes, { onResourceMetrics });
      await expect(session.close()).rejects.toBe(cleanupError);
      expect(onResourceMetrics).toHaveBeenCalledOnce();
      expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
        format: 'pptx',
        scope: 'session',
        status: 'error',
      }));
    } finally {
      free.mockRestore();
    }
  });

  it('is one-pass and rejects iteration after ownership has closed', async () => {
    const session = await openPptxPresentation(bytes);
    const iterator = session.slides();
    expect((await iterator.next()).value?.index).toBe(0);
    await iterator.return();
    await expect(session.slides().next()).rejects.toThrow(/closed|one-pass/);
  });

  it('normalizes package limits and aborts deterministically', async () => {
    await expect(openPptxPresentation(bytes, {
      resourceLimits: { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: null },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
    });
    await expect(openPptxPresentation(bytes, {
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
    await expect(openPptxPresentation(bytes, { signal: before.signal }))
      .rejects.toMatchObject({ name: 'AbortError' });

    const after = new AbortController();
    const session = await openPptxPresentation(bytes, { signal: after.signal });
    const iterator = session.slides();
    expect((await iterator.next()).value?.index).toBe(0);
    after.abort();
    await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  });
});

type ArchivePrototype = {
  free(): void;
  extract_image(path: string): Uint8Array;
  pull_slide(
    slideIndex: number,
    operationId: number,
    generation: number,
    byteCredit: number,
  ): Uint8Array;
};

function archivePrototype(): ArchivePrototype {
  return (pptxWasm as unknown as { PptxArchive: { prototype: ArchivePrototype } })
    .PptxArchive.prototype;
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  const promise = new Promise<T>((res) => { resolve = res; });
  return { promise, resolve };
}

/** A deck whose presentation theme sets dk1 but whose master names a missing
 * theme part: the parsed slide has no theme and its title text has no colour
 * of its own. */
function themeLessMasterDeck(): Uint8Array {
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"`;
  const rels = (items: Array<[string, string, string]>): string =>
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + items.map(([id, type, target]) =>
      `<Relationship Id="${id}" Type="${r}/${type}" Target="${target}"/>`).join('')
    + '</Relationships>';
  const clrMap = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
    + 'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
    + 'hlink="hlink" folHlink="folHlink"';
  const scheme = Object.entries({ dk1: '10263F', lt1: 'F9F9F9', dk2: '111111', lt2: 'EEEEEE' })
    .map(([name, value]) => `<a:${name}><a:srgbClr val="${value}"/></a:${name}>`).join('');
  const type = (part: string, kind: string): string =>
    `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${kind}+xml"/>`;
  return storedZip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + type('ppt/presentation.xml', 'presentationml.presentation.main')
      + type('ppt/slideMasters/slideMaster1.xml', 'presentationml.slideMaster')
      + type('ppt/slideLayouts/slideLayout1.xml', 'presentationml.slideLayout')
      + type('ppt/slides/slide1.xml', 'presentationml.slide')
      + type('ppt/theme/theme1.xml', 'theme')
      + '</Types>',
    '_rels/.rels': rels([['pres', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `<p:presentation ${ns}><p:sldMasterIdLst>`
      + '<p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst>'
      + '<p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst>'
      + '<p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels([
      ['master', 'slideMaster', 'slideMasters/slideMaster1.xml'],
      ['slide', 'slide', 'slides/slide1.xml'],
      ['theme', 'theme', 'theme/theme1.xml'],
    ]),
    'ppt/theme/theme1.xml': `<a:theme xmlns:a="${a}" name="T"><a:themeElements>`
      + `<a:clrScheme name="C">${scheme}</a:clrScheme>`
      + '<a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>'
      + '<a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>'
      + '<a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/>'
      + '<a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>',
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${ns}><p:cSld><p:spTree/></p:cSld>`
      + `<p:clrMap ${clrMap}/></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([['theme', 'theme', '../theme/missing.xml']]),
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${ns}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([
      ['master', 'slideMaster', '../slideMasters/slideMaster1.xml'],
    ]),
    'ppt/slides/slide1.xml': `<p:sld ${ns}><p:cSld><p:spTree><p:sp><p:nvSpPr>`
      + '<p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
      + '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
      + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="1"/><a:r><a:rPr/>'
      + '<a:t>title-1</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    'ppt/slides/_rels/slide1.xml.rels': rels([
      ['layout', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
    ]),
  });
}

/** Canvas whose context records the fill style of every painted text run. */
function textRecordingCanvas(): { canvas: NodeCanvasLike; fills: Map<string, unknown> } {
  const fills = new Map<string, unknown>();
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1, letterSpacing: '0px',
  };
  const canvas = { width: 1, height: 1 } as NodeCanvasLike;
  const ctx = new Proxy(state, {
    get(_target, prop: string) {
      if (prop === 'canvas') return canvas;
      if (prop in state) return state[prop];
      if (prop === 'measureText') {
        return (text: string) => ({
          width: text.length * 7, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        });
      }
      if (prop === 'fillText') return (text: string) => { fills.set(text, state.fillStyle); };
      if (prop === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      return () => undefined;
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  });
  (canvas as { getContext: unknown }).getContext = () => ctx;
  return { canvas, fills };
}

function fakeCanvas(): NodeCanvasLike {
  return { width: 1, height: 1, getContext: vi.fn() as unknown as NodeCanvasLike['getContext'] };
}

function fakeFactory(): NodeCanvasFactory {
  return {
    createCanvas: vi.fn(() => fakeCanvas()),
    loadImage: vi.fn(async () => ({ width: 1, height: 1 })),
  };
}

describe('Node PPTX session text colour', () => {
  it('paints uncoloured text on a theme-less parsed slide black, not the presentation dk1', async () => {
    const session = await openPptxPresentation(themeLessMasterDeck());
    try {
      expect(session.slideCount).toBe(1);
      for await (const slide of session) {
        expect(slide.theme).toBeUndefined();
        const { canvas, fills } = textRecordingCanvas();
        await session.renderSlide(canvas, slide, { width: 960, dpr: 1, factory: fakeFactory() });
        expect(fills.get('title-1')).toBe('#000000');
      }
    } finally {
      await session.close();
    }
  });
});
