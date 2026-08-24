import init, { PptxArchive, reinit } from './wasm/pptx_parser.js';
import { renderSlide, type PptxTextRunInfo } from './renderer';
import { PPTX_GOOGLE_FONTS } from './google-fonts';
import {
  findPreflightMimeType,
  PresentationPreflightBuilder,
  type PresentationPreflight,
} from './presentation-preflight';
import { PptxSlideRepository } from './slide-repository';
import { loadPptxSlideFromCursor, readPptxSlideCursorUsage } from './slide-cursor-operation';
import { SlidePullWorker } from './slide-pull-worker';
import {
  preloadGoogleFonts,
  decodeDataUrl,
  WasmParserHost,
  dropDecodedBitmapCache,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  HARD_MAX_PPTX_CACHED_SLIDES,
  HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  resourcePolicyForWasm,
  serializeWorkerError,
  loadWorkerRenderers,
  type LoadedWorkerRenderers,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import type {
  PresentationBootstrap,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import { hitTestPptxSlideContext } from './element-selection';

const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (archive) => archive.free(),
  reinit,
});

const executeArchive = <T>(operation: (archive: PptxArchive) => T): T => {
  const archive = host.archive;
  if (!archive) throw new Error('Presentation not loaded');
  return host.run(() => operation(archive));
};

const slidePull = new SlidePullWorker(
  () => host.archive,
  undefined,
  (operation) => executeArchive(operation),
);

let preflight: PresentationPreflight | null = null;
let preflightBuilder: PresentationPreflightBuilder | null = null;
let slides: PptxSlideRepository | null = null;
let generation = 0;
let nextOperationId = 1;
type PresentationLifecycleState = 'empty' | 'opening' | 'ready' | 'failed';
let presentationState: PresentationLifecycleState = 'empty';
let fontsLoaded: Promise<unknown> = Promise.resolve();
let resourceUsage: OoxmlResourceUsageSnapshot | undefined;
let renderers: LoadedWorkerRenderers = {};
const rawParts = new BoundedRawPartCache({
  maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
  maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
});

function reservePresentationParse(): void {
  if (presentationState !== 'empty') {
    const error = new Error('this PPTX render worker already owns a presentation parse');
    error.name = 'PptxWorkerStateError';
    throw Object.assign(error, { code: 'ooxml-pptx-parse-already-started' });
  }
  presentationState = 'opening';
}

const post = (message: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (value: unknown, transfer?: Transferable[]) => void)(message, transfer);

function requirePreflight(): PresentationPreflight {
  if (!preflight) throw new Error('No pptx loaded');
  return preflight;
}

function requireSlides(): PptxSlideRepository {
  if (!slides) throw new Error('No pptx loaded');
  return slides;
}

function loadSlide(slideIndex: number) {
  const operationId = nextOperationId++;
  const currentGeneration = generation;
  return slidePull.run(() => loadPptxSlideFromCursor(
    (operation) => executeArchive(operation),
    slideIndex,
    { operationId, generation: currentGeneration },
    preflightBuilder
      ? (index, slide, usage) => {
        resourceUsage = usage;
        return preflightBuilder?.prepareSlide(slide, usage);
      }
      : undefined,
  ));
}

function getMedia(path: string): Promise<Blob> {
  const mimeType = findPreflightMimeType(requirePreflight(), path);
  return rawParts.get(path, mimeType, () => slidePull.run(() => {
    const bytes = executeArchive((archive) => archive.extract_media(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  }));
}

function getImage(path: string, mimeType: string): Promise<Blob> {
  return rawParts.get(path, mimeType, () => slidePull.run(() => {
    const bytes = executeArchive((archive) => archive.extract_image(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  }));
}

async function openPresentation(request: Extract<RenderWorkerRequest, { kind: 'parse' }>) {
  await slidePull.reset();
  slides?.clear();
  slides = null;
  preflight = null;
  preflightBuilder = null;
  generation += 1;
  nextOperationId = 1;
  rawParts.clear();
  dropDecodedBitmapCache(getImage);
  dropSvgImageCache(getImage);
  fontsLoaded = Promise.resolve();
  resourceUsage = undefined;
  renderers = await loadWorkerRenderers(request.renderers);

  const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(request.resourcePolicy);
  const bootstrap = await slidePull.run(() => executeArchiveFromNew(
    request.buffer,
    maxEntry,
    maxTotal,
    maxEntries,
  ));
  preflightBuilder = new PresentationPreflightBuilder(bootstrap);
  slides = new PptxSlideRepository({
    slideCount: bootstrap.slideCount,
    maxCachedSlides: HARD_MAX_PPTX_CACHED_SLIDES,
    maxCachedStructuralBytes: HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
    loadSlide,
  });
  for (let index = 0; index < bootstrap.slideCount; index += 1) {
    await slides.withSlide(index, () => undefined);
  }
  preflight = preflightBuilder.finish();
  preflightBuilder = null;
  if (request.useGoogleFonts) {
    fontsLoaded = preloadGoogleFonts(preflight.fontPreloadNames, PPTX_GOOGLE_FONTS);
  }
  return preflight;
}

function executeArchiveFromNew(
  buffer: ArrayBuffer,
  maxEntry: bigint | null | undefined,
  maxTotal: bigint | null | undefined,
  maxEntries: bigint | null | undefined,
): PresentationBootstrap {
  return host.run(() => {
    const archive = new PptxArchive(
      new Uint8Array(buffer),
      maxEntry,
      maxTotal,
      maxEntries,
    );
    host.setArchive(archive);
    return JSON.parse(
      new TextDecoder().decode(archive.presentation_bootstrap()),
    ) as PresentationBootstrap;
  });
}

self.onmessage = async (event: MessageEvent<RenderWorkerRequest>) => {
  const request = event.data;
  if (request.kind === 'init') {
    host.setWasmInput(decodeDataUrl(request.wasmUrl) ?? request.wasmUrl);
    return;
  }

  let ownsParseReservation = false;
  try {
    if (request.kind === 'parse') {
      reservePresentationParse();
      ownsParseReservation = true;
    }
    await host.ensureReady();

    if (request.kind === 'parse') {
      const compact = await openPresentation(request);
      post({
        kind: 'presentationReady',
        id: request.id,
        preflight: compact,
        usage: resourceUsage,
      });
      presentationState = 'ready';
      return;
    }

    const compact = requirePreflight();
    await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));

    if (request.kind === 'renderSlide') {
      const { bitmap, runs } = await requireSlides().withSlide(request.slideIndex, async (slide) => {
        await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));
        await fontsLoaded;
        const canvas = new OffscreenCanvas(1, 1);
        const runs: PptxTextRunInfo[] = [];
        await renderSlide(canvas, slide, compact.slideWidth, compact.slideHeight, {
          width: request.width,
          dpr: request.dpr,
          defaultTextColor: compact.defaultTextColor,
          majorFont: compact.majorFont,
          minorFont: compact.minorFont,
          hlinkColor: compact.hlinkColor,
          fetchMedia: getMedia,
          fetchImage: getImage,
          skipMediaControls: request.skipMediaControls,
          dim: request.dim,
          math: renderers.math,
          threeD: renderers.threeD,
          regionMap: renderers.regionMap,
          chartEx: renderers.chartEx,
        }, (run) => runs.push(run));
        return { bitmap: canvas.transferToImageBitmap(), runs };
      });
      post({ kind: 'slideRendered', id: request.id, bitmap, runs }, [bitmap]);
      return;
    }

    if (request.kind === 'collectRuns') {
      const runs = await requireSlides().withSlide(request.slideIndex, async (slide) => {
        await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));
        await fontsLoaded;
        const canvas = new OffscreenCanvas(1, 1);
        const runs: PptxTextRunInfo[] = [];
        await renderSlide(canvas, slide, compact.slideWidth, compact.slideHeight, {
          width: request.width,
          defaultTextColor: compact.defaultTextColor,
          majorFont: compact.majorFont,
          minorFont: compact.minorFont,
          hlinkColor: compact.hlinkColor,
          fetchMedia: getMedia,
          fetchImage: getImage,
          math: renderers.math,
          threeD: renderers.threeD,
          regionMap: renderers.regionMap,
          chartEx: renderers.chartEx,
        }, (run) => runs.push(run));
        return runs;
      });
      post({ kind: 'runsCollected', id: request.id, runs });
      return;
    }

    if (request.kind === 'hitTestElement') {
      const context = await requireSlides().withSlide(request.slideIndex, (slide) =>
        hitTestPptxSlideContext(request.slideIndex, slide, request.point, request.options));
      post({ kind: 'elementHit', id: request.id, context });
      return;
    }

    if (request.kind === 'extractMedia') {
      const bytes = await (await getMedia(request.path)).arrayBuffer();
      post({ kind: 'mediaExtracted', id: request.id, bytes }, [bytes]);
      return;
    }

    if (request.kind === 'extractImage') {
      const mimeType = findPreflightMimeType(compact, request.path);
      const bytes = await (await getImage(request.path, mimeType)).arrayBuffer();
      post({ kind: 'imageExtracted', id: request.id, bytes }, [bytes]);
      return;
    }
    if (request.kind === 'resourceUsage') {
      const usage = decodeOoxmlResourceUsage(executeArchive(
        (archive) => archive.resource_usage(),
      ));
      post({ kind: 'resourceUsage', id: request.id, usage });
      return;
    }

    if (request.kind === 'toMarkdown') {
      const markdown = await slidePull.run(() =>
        executeArchive((archive) => archive.to_markdown()));
      post({ kind: 'markdownRendered', id: request.id, markdown });
    }
  } catch (error) {
    if (ownsParseReservation) presentationState = 'failed';
    if (request.kind === 'parse') {
      slides?.clear();
      slides = null;
      preflight = null;
      preflightBuilder = null;
    }
    try {
      post({ kind: 'error', id: request.id, ...serializeWorkerError(error) });
    } catch {
      // The worker response channel is unavailable; local cleanup already ran.
    }
  }
};
