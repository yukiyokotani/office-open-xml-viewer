import init, { PptxArchive, reinit } from './wasm/pptx_parser.js';
import type { PptxTextRunInfo } from './renderer';
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
  isWorkerSvgDecodeResponse,
  postOwnedImageBitmap,
  WorkerSvgDecodeClient,
  type LoadedWorkerRenderers,
  type WorkerSvgDecodeResponse,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import type {
  PresentationBootstrap,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import { findPptxElementBoundsByIds, hitTestPptxSlideContext } from './element-selection';
import { excludeEmbeddedFontFamilies, loadEmbeddedFonts } from './embedded-fonts';
import { ProgressivePreflightGate } from './progressive-preflight-gate';

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
let availableSlideCount = 0;
const slideAvailabilityWaiters = new Set<() => void>();
const progressivePreflightGate = new ProgressivePreflightGate();
let generation = 0;
let nextOperationId = 1;
type PresentationLifecycleState = 'empty' | 'opening' | 'ready' | 'failed';
let presentationState: PresentationLifecycleState = 'empty';
let fontsLoaded: Promise<unknown> = Promise.resolve();
let embeddedFontAliases: ReadonlyMap<string, string> = new Map();
let embeddedFontAuthoredFamilies: ReadonlyMap<string, string> = new Map();
let resourceUsage: OoxmlResourceUsageSnapshot | undefined;
let renderers: LoadedWorkerRenderers = {};
const rawParts = new BoundedRawPartCache({
  maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
  maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
});
// Keep the renderer behind an explicit module boundary. The production worker
// is flattened into one self-contained asset, and a static function import can
// otherwise be hoisted past the initializers of shared DrawingML dependencies
// that are also reached by optional renderers. Awaiting the module preserves ESM
// initialization order before any text metrics use those shared unit constants.
const rendererModule = import('./renderer');

function reservePresentationParse(): void {
  if (presentationState !== 'empty') {
    const error = new Error('this PPTX render worker already owns a presentation parse');
    error.name = 'PptxWorkerStateError';
    throw Object.assign(error, { code: 'ooxml-pptx-parse-already-started' });
  }
  presentationState = 'opening';
}

const rawPost = (message: unknown, transfer?: Transferable[]) =>
  (self.postMessage as (value: unknown, transfer?: Transferable[]) => void)(message, transfer);
const post = (message: RenderWorkerResponse, transfer?: Transferable[]) => rawPost(message, transfer);
const svgDecodeClient = new WorkerSvgDecodeClient(rawPost);

function requirePreflight(): PresentationPreflight {
  if (preflight) return preflight;
  if (preflightBuilder?.acceptedSlideCount) return preflightBuilder.snapshot();
  throw new Error('No pptx loaded');
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

function getFontBytes(path: string): Promise<Uint8Array> {
  return slidePull.run(() => {
    const bytes = executeArchive((archive) => archive.extract_font(path));
    return new Uint8Array(bytes as Uint8Array);
  });
}

async function openPresentation(request: Extract<RenderWorkerRequest, { kind: 'parse' }>) {
  progressivePreflightGate.reset();
  await slidePull.reset();
  slides?.clear();
  slides = null;
  preflight = null;
  preflightBuilder = null;
  availableSlideCount = 0;
  wakeSlideAvailabilityWaiters();
  generation += 1;
  nextOperationId = 1;
  rawParts.clear();
  dropDecodedBitmapCache(getImage);
  dropSvgImageCache(getImage);
  fontsLoaded = Promise.resolve();
  embeddedFontAliases = new Map();
  embeddedFontAuthoredFamilies = new Map();
  resourceUsage = undefined;
  renderers = await loadWorkerRenderers(request.renderers);

  const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(request.resourcePolicy);
  const bootstrap = await slidePull.run(() => executeArchiveFromNew(
    request.buffer,
    maxEntry,
    maxTotal,
    maxEntries,
  ));
  // The retained archive exposes font reads as independent operations, so font
  // decoding can overlap the sequential slide preflight without sharing cursor
  // ownership. First paint still waits for `fontsLoaded` below.
  const embeddedFontsLoaded = loadEmbeddedFonts(bootstrap.embeddedFonts, getFontBytes);
  preflightBuilder = new PresentationPreflightBuilder(bootstrap);
  slides = new PptxSlideRepository({
    slideCount: bootstrap.slideCount,
    maxCachedSlides: HARD_MAX_PPTX_CACHED_SLIDES,
    maxCachedStructuralBytes: HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
    loadSlide,
  });
  if (request.progressiveLayout) {
    const loadedGoogleFonts = new Set<string>();
    const ensureFonts = async (): Promise<void> => {
      const embedded = await embeddedFontsLoaded;
      embeddedFontAliases = embedded.aliases;
      embeddedFontAuthoredFamilies = embedded.authoredFamilies;
      if (!request.useGoogleFonts) return;
      const requested = excludeEmbeddedFontFamilies(
        preflightBuilder!.currentFontPreloadNames,
        embedded.aliases,
      ).filter((name): name is string => !!name && !loadedGoogleFonts.has(name));
      for (const name of requested) loadedGoogleFonts.add(name);
      if (requested.length) await preloadGoogleFonts(requested, PPTX_GOOGLE_FONTS);
    };
    for (let index = 0; index < bootstrap.slideCount; index += 1) {
      await slides.withSlide(index, () => undefined);
      await ensureFonts();
      availableSlideCount = index + 1;
      const slide = preflightBuilder.latestSlide;
      if (!slide) throw new Error(`PPTX progressive preflight lost slide ${index}`);
      wakeSlideAvailabilityWaiters();
      // Register before publishing so even an immediate host acknowledgement
      // cannot race past the checkpoint.
      const hostAcknowledgement = progressivePreflightGate.wait(request.id, availableSlideCount);
      post({
        kind: 'presentationLayoutPartial',
        forId: request.id,
        ...(index === 0 ? { bootstrap } : {}),
        availableSlides: availableSlideCount,
        slide,
        fontPreloadNames: preflightBuilder.currentFontPreloadNames,
        usage: resourceUsage,
      });
      await hostAcknowledgement;
    }
    preflight = preflightBuilder.finish();
    preflightBuilder = null;
    fontsLoaded = Promise.resolve();
  } else {
    for (let index = 0; index < bootstrap.slideCount; index += 1) {
      await slides.withSlide(index, () => undefined);
    }
    preflight = preflightBuilder.finish();
    preflightBuilder = null;
    availableSlideCount = bootstrap.slideCount;
    wakeSlideAvailabilityWaiters();
    fontsLoaded = (async () => {
      const embedded = await embeddedFontsLoaded;
      embeddedFontAliases = embedded.aliases;
      embeddedFontAuthoredFamilies = embedded.authoredFamilies;
      if (!request.useGoogleFonts) return embedded.faces;
      const substitutes = await preloadGoogleFonts(
        excludeEmbeddedFontFamilies(preflight.fontPreloadNames, embedded.aliases),
        PPTX_GOOGLE_FONTS,
      );
      return [...embedded.faces, ...substitutes];
    })();
  }
  return preflight;
}

function wakeSlideAvailabilityWaiters(): void {
  for (const resolve of slideAvailabilityWaiters) resolve();
  slideAvailabilityWaiters.clear();
}

async function waitForSlideAvailability(slideIndex: number): Promise<void> {
  if (presentationState === 'empty') throw new Error('No pptx loaded');
  while (slideIndex >= availableSlideCount && presentationState === 'opening') {
    await new Promise<void>((resolve) => slideAvailabilityWaiters.add(resolve));
  }
  if (presentationState === 'failed') throw new Error('PPTX progressive preflight failed');
  if (slideIndex >= availableSlideCount) {
    throw new Error(`Slide index ${slideIndex} out of range (count: ${availableSlideCount})`);
  }
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

self.onmessage = async (event: MessageEvent<RenderWorkerRequest | WorkerSvgDecodeResponse>) => {
  const request = event.data;
  if (isWorkerSvgDecodeResponse(request)) {
    svgDecodeClient.accept(request);
    return;
  }
  if (request.kind === 'init') {
    host.setWasmInput(decodeDataUrl(request.wasmUrl) ?? request.wasmUrl);
    return;
  }
  if (request.kind === 'continuePresentationPreflight') {
    progressivePreflightGate.continue(request.forId, request.availableSlides);
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

    if ('slideIndex' in request) await waitForSlideAvailability(request.slideIndex);

    const compact = requirePreflight();
    await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));

    if (request.kind === 'renderSlide') {
      const { bitmap, runs } = await requireSlides().withSlide(request.slideIndex, async (slide) => {
        await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));
        await fontsLoaded;
        const { renderSlideWithEmbeddedFonts } = await rendererModule;
        const canvas = new OffscreenCanvas(1, 1);
        const runs: PptxTextRunInfo[] = [];
        await renderSlideWithEmbeddedFonts(canvas, slide, compact.slideWidth, compact.slideHeight, {
          width: request.width,
          dpr: request.dpr,
          imageResources: request.imageResources,
          defaultTextColor: compact.defaultTextColor,
          majorFont: compact.majorFont,
          minorFont: compact.minorFont,
          hlinkColor: compact.hlinkColor,
          embeddedFontAliases,
          embeddedFontAuthoredFamilies,
          fetchMedia: getMedia,
          fetchImage: getImage,
          svgDecoder: svgDecodeClient.decode,
          skipMediaControls: request.skipMediaControls,
          dim: request.dim,
          math: renderers.math,
          threeD: renderers.threeD,
          regionMap: renderers.regionMap,
          chartEx: renderers.chartEx,
          tiff: renderers.tiff,
        }, (run) => runs.push(run));
        return { bitmap: canvas.transferToImageBitmap(), runs };
      });
      postOwnedImageBitmap(post, { kind: 'slideRendered', id: request.id, bitmap, runs });
      return;
    }

    if (request.kind === 'collectRuns') {
      const runs = await requireSlides().withSlide(request.slideIndex, async (slide) => {
        await slidePull.run(() => executeArchive((archive) => archive.assert_healthy()));
        await fontsLoaded;
        const { renderSlideWithEmbeddedFonts } = await rendererModule;
        const canvas = new OffscreenCanvas(1, 1);
        const runs: PptxTextRunInfo[] = [];
        await renderSlideWithEmbeddedFonts(canvas, slide, compact.slideWidth, compact.slideHeight, {
          width: request.width,
          defaultTextColor: compact.defaultTextColor,
          majorFont: compact.majorFont,
          minorFont: compact.minorFont,
          hlinkColor: compact.hlinkColor,
          embeddedFontAliases,
          embeddedFontAuthoredFamilies,
          fetchMedia: getMedia,
          fetchImage: getImage,
          svgDecoder: svgDecodeClient.decode,
          math: renderers.math,
          threeD: renderers.threeD,
          regionMap: renderers.regionMap,
          chartEx: renderers.chartEx,
          tiff: renderers.tiff,
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

    if (request.kind === 'resolveElementBounds') {
      const bounds = await requireSlides().withSlide(request.slideIndex, (slide) =>
        findPptxElementBoundsByIds(slide, request.elementIds));
      post({ kind: 'elementBoundsResolved', id: request.id, bounds });
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
    if (request.kind === 'extractFont') {
      const bytes = (await getFontBytes(request.path)).buffer as ArrayBuffer;
      post({ kind: 'fontExtracted', id: request.id, bytes }, [bytes]);
      return;
    }
    if (request.kind === 'resourceUsage') {
      const usage = decodeOoxmlResourceUsage(executeArchive(
        (archive) => archive.resource_usage(),
      ));
      post({ kind: 'resourceUsage', id: request.id, usage });
      return;
    }

  } catch (error) {
    if (ownsParseReservation) {
      presentationState = 'failed';
      wakeSlideAvailabilityWaiters();
    }
    if (request.kind === 'parse') {
      progressivePreflightGate.reset();
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
