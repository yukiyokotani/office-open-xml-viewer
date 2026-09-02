import type { DocxDocumentModel, BodyElement, DocxTextRunInfo } from './types';
import type { LayoutServices, MathRenderer } from './layout/types.js';
import type { ChartThreeDRenderer, ChartRegionMapRenderer, ChartExRenderer, ImageResourceOptions, TiffRenderer, SvgBlobDecoder } from '@silurus/ooxml-core';
export type { DocxTextRunInfo } from './types';
import { bodyMathOccurrences } from './layout/resources.js';
import { paintResourceRegistryOf, privateResourceLookupOf } from './layout/runtime-state.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { rasterPaintOccurrencesForPage } from './layout/text-index.js';
import { textRunsForPage } from './text-run-projection.js';
import { dropBrowserImageCache } from './paint/browser-images.js';
import { canvasPageScale, renderSelectedDocumentPage } from './paint/canvas-document.js';
import { ensureDocumentLayoutVariants } from './layout/document.js';
import { prepareMathResources } from './paint/math-resources.js';
import { createLayoutServices } from './layout-runtime.js';
import {
  type LayoutSourceStore,
} from './layout/layout-source-store.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';

/** True if any currently representable document story contains OMML. The body
 * array form remains supported for existing callers. */
export function documentHasMath(input: BodyElement[] | DocxDocumentModel): boolean {
  return (Array.isArray(input)
    ? bodyMathOccurrences(input)
    : layoutSourceStore(input).mathOccurrences).length > 0;
}

/** Convert equations before layout. Math resources use only normalized,
 * structural SourceRef/resourceKey facts; parser object identity is irrelevant. */
export async function prepareMathRuns(
  input: BodyElement[] | DocxDocumentModel,
  math: MathRenderer,
) {
  if (Array.isArray(input)) {
    throw new TypeError('prepareMathRuns requires a document model so every story has an explicit structural source');
  }
  return prepareMathResources(layoutSourceStore(input).mathOccurrences, math);
}

export interface RenderDocumentOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /**
   * Lazy image-byte loader: fetch the raw bytes for an embedded image by zip
   * path, wrapped in a Blob of the given MIME (twin of pptx's `fetchImage`).
   * Supplied by {@link DocxDocument} (routing to its `getImage`), so the
   * renderer decodes images on demand instead of from inlined base64. When
   * omitted, images are skipped (no byte source).
   */
  fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  /** Internal worker-to-Window SVG decoder. */
  svgDecoder?: SvgBlobDecoder;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — the "current" instant that a
   *  DATE/TIME field formats through its `\@` date picture (§17.16.4.1). Accepts a
   *  `Date` or epoch-ms number. Default = the real current time (`Date.now()` at
   *  render). Provide a fixed value to make DATE/TIME field output deterministic
   *  (e.g. in tests / reproducible exports). */
  currentDate?: Date | number;
  /** ECMA-376 §17.13.5 tracked-change view: `true` = markup view (revision
   *  decoration + change bars), absent/false = final view (deletions hidden).
   *  Selects the cached layout variant — see RenderPageOptions. */
  showTrackedChanges?: boolean;
  /** Internal per-document service snapshot. Public render options never expose it. */
  layoutServices?: LayoutServices;
  /** Internal load-time default captured once and mirrored into worker mode. */
  defaultCurrentDateMs?: number;
  /** Internal load-time optional 3-D renderer retained by DocxDocument. */
  threeD?: ChartThreeDRenderer;
  /** Internal load-time optional Region Map renderer retained by DocxDocument. */
  regionMap?: ChartRegionMapRenderer;
  /** Internal load-time optional ChartEx renderer retained by DocxDocument. */
  chartEx?: ChartExRenderer;
  /** Internal load-time optional TIFF codec retained by DocxDocument. */
  tiff?: TiffRenderer;
  /** Adaptive decoded-raster memory policy shared by every OOXML renderer. */
  imageResources?: ImageResourceOptions;
}

export function dropColorReplacedCache(
  fetchImage: (path: string, mime: string) => Promise<Blob>,
): void {
  dropBrowserImageCache(fetchImage);
}

function normalizeRenderOptions(
  source: LayoutSourceStore,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  options: RenderDocumentOptions,
) {
  const services = options.layoutServices ?? createLayoutServices(
    source,
    source.fatalParse === null ? {
      measureContext: canvas.getContext('2d') as
        | CanvasRenderingContext2D
        | OffscreenCanvasRenderingContext2D
        | null,
    } : {},
  );
  const retainedSource = layoutSourceStoreOf(services);
  if (retainedSource && retainedSource !== source) {
    throw new Error('Layout services belong to a different document source');
  }
  const defaultCurrentDateMs = options.defaultCurrentDateMs ?? Date.now();
  ensureDocumentLayoutVariants(
    services,
    defaultCurrentDateMs,
    () => source,
  );
  const selection = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate,
    defaultCurrentDateMs,
    showTrackedChanges: options.showTrackedChanges,
  }, pageIndex);
  const scale = canvasPageScale(selection.page, options.width);
  return {
    selection,
    paintOptions: {
      width: options.width,
      dpr: options.dpr,
      defaultTextColor: options.defaultTextColor,
      fetchImage: options.fetchImage,
      svgDecoder: options.svgDecoder,
      parseError: source.fatalParse !== null,
      registry: paintResourceRegistryOf(services),
      rasterPaintOccurrences: rasterPaintOccurrencesForPage(selection.layout, pageIndex),
      privateResources: privateResourceLookupOf<CanvasImageSource>(services),
      textRuns: options.onTextRun
        ? textRunsForPage(selection.layout, pageIndex, { scale })
        : [],
      onTextRun: options.onTextRun,
      threeD: options.threeD,
      regionMap: options.regionMap,
      chartEx: options.chartEx,
      tiff: options.tiff,
      imageResources: options.imageResources,
    },
  };
}

/** Internal production entry: paint from the same sealed source used by layout. */
export async function renderLayoutSourceToCanvas(
  source: LayoutSourceStore,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);
  return renderSelectedDocumentPage(
    normalized.selection.layout,
    normalized.selection.page,
    canvas,
    normalized.paintOptions,
  );
}

export async function renderDocumentToCanvas(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  return renderLayoutSourceToCanvas(layoutSourceStore(doc), canvas, pageIndex, opts);
}
