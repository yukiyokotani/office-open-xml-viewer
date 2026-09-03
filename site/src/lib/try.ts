// "Try yours" — render a user-supplied file entirely in the browser. The file
// is read with FileReader/arrayBuffer and parsed by the WASM engines; it never
// leaves the page (no upload, no server).
import {
  PptxPresentation,
  PptxScrollViewer,
  type PptxScrollViewerOptions,
} from '@silurus/ooxml-pptx';
import {
  DocxDocument,
  DocxScrollViewer,
  type DocxScrollViewerOptions,
} from '@silurus/ooxml-docx';
import { XlsxSheetViewer, XlsxViewer } from '@silurus/ooxml-xlsx';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';
import { tiff } from '../../../src/tiff';

// Opt-in OMML equation engine — enabled here so user-supplied docx/pptx with
// equations render. (In the published library this is `@silurus/ooxml/math`.)
const fullRenderers = { threeD, regionMap, chartEx, tiff };

const VIEWER_GAP = 26;
const MIN_SCALE = 0.5;

// Disposes the previous viewer and its parser/worker resources when a new file
// is loaded.
let activeCleanup: (() => void) | null = null;
let renderGeneration = 0;

export interface RenderResult {
  format: 'docx' | 'xlsx' | 'pptx';
  units: number; // pages / slides; 0 for xlsx (sheet-based)
  unitLabel: string;
  /** Resolves with the authoritative count when progressive layout finishes. */
  finalUnits?: Promise<number>;
}

export interface RenderProgress {
  format: 'docx' | 'pptx';
  units: number;
  unitLabel: 'page' | 'slide';
}

function scrollViewerHost(stage: HTMLElement): HTMLDivElement {
  const host = document.createElement('div');
  host.className = 'lv-scroll-viewer';
  stage.appendChild(host);
  return host;
}

class SupersededRenderError extends Error {
  override name = 'AbortError';
  constructor() {
    super('This file render was superseded by a newer selection.');
  }
}

function assertCurrentRender(generation: number): void {
  if (generation !== renderGeneration) throw new SupersededRenderError();
}

/** Tear down the current viewer when Try Yours leaves the page. */
export function disposeRenderedFile(): void {
  renderGeneration++;
  activeCleanup?.();
  activeCleanup = null;
}

export async function renderFile(
  stage: HTMLElement,
  file: File,
  onProgress?: (progress: Readonly<RenderProgress>) => void,
): Promise<RenderResult> {
  const generation = ++renderGeneration;
  // Any new selection — supported or not — supersedes and releases the current
  // viewer before validation. The page hides the panel on validation failure, so
  // retaining its media/worker resources would be both invisible and leaked.
  activeCleanup?.();
  activeCleanup = null;

  const ext = file.name.split('.').pop()?.toLowerCase();
  if (
    ext !== 'docx'
    && ext !== 'xlsx'
    && ext !== 'pptx'
    && ext !== 'csv'
    && ext !== 'tsv'
  ) {
    throw new Error('Unsupported file — choose a .docx, .xlsx or .pptx file.');
  }

  const buffer = await file.arrayBuffer();
  assertCurrentRender(generation);
  stage.innerHTML = '';

  if (ext === 'xlsx') {
    const host = document.createElement('div');
    host.className = 'lv-xlsx';
    stage.appendChild(host);
    const viewer = new XlsxViewer(host, {
      mode: 'main',
      useGoogleFonts: true,
      showZoomSlider: true,
      comments: false,
      math,
      ...fullRenderers,
    });
    try {
      await viewer.load(buffer);
    } catch (error) {
      viewer.destroy();
      throw error;
    }
    if (generation !== renderGeneration) {
      viewer.destroy();
      throw new SupersededRenderError();
    }
    activeCleanup = () => viewer.destroy();
    return { format: 'xlsx', units: 0, unitLabel: 'sheet' };
  }

  if (ext === 'csv' || ext === 'tsv') {
    const host = document.createElement('div');
    host.className = 'lv-xlsx';
    const canvas = document.createElement('canvas');
    canvas.style.width = '100%';
    canvas.style.height = '100%';
    host.appendChild(canvas);
    stage.appendChild(host);
    const viewer = new XlsxSheetViewer(canvas, {
      mode: 'main',
      useGoogleFonts: true,
      comments: false,
      math,
      ...fullRenderers,
    });
    try {
      await viewer.load(buffer, { format: ext });
    } catch (error) {
      viewer.destroy();
      throw error;
    }
    if (generation !== renderGeneration) {
      viewer.destroy();
      throw new SupersededRenderError();
    }
    activeCleanup = () => viewer.destroy();
    return { format: 'xlsx', units: 0, unitLabel: 'sheet' };
  }

  if (ext === 'pptx') {
    const host = scrollViewerHost(stage);
    let viewer!: PptxScrollViewer;
    const viewerOptions: PptxScrollViewerOptions = {
      gap: VIEWER_GAP,
      overscan: 0,
      enableTextSelection: true,
      enableMediaPlayback: true,
      mediaOverscan: 1,
      enableZoom: true,
      zoomMin: MIN_SCALE,
      pageShadow: false,
      useGoogleFonts: true,
      comments: false,
      math,
      // Keep progressive preflight and paint off the UI thread while the user
      // scrolls through slides that are still becoming available.
      mode: 'worker',
      progressiveLayout: true,
      onLayoutPartial: ({ availableUnits }) => {
        if (generation !== renderGeneration) return;
        // Try Yours intentionally keeps every already-paintable slide mounted:
        // the next slides continue loading without requiring reader scroll and
        // the complete text layer remains available to native Find.
        viewerOptions.overscan = availableUnits;
        viewer.relayout();
        onProgress?.({ format: 'pptx', units: availableUnits, unitLabel: 'slide' });
      },
      ...fullRenderers,
    };
    viewer = new PptxScrollViewer(host, viewerOptions);
    // Do not force an absolute scale here. ScrollViewer derives its initial
    // scale from the laid-out container width and keeps that fit on resize, so
    // a wide slide never opens with horizontal overflow in this workspace.
    try {
      await viewer.load(buffer);
    } catch (error) {
      viewer.destroy();
      throw error;
    }
    if (generation !== renderGeneration) {
      viewer.destroy();
      throw new SupersededRenderError();
    }
    activeCleanup = () => viewer.destroy();

    const mountAllSlides = (): number => {
      assertCurrentRender(generation);
      // Native Find needs every slide's text layer in the DOM. Keep the opening
      // progressive window virtualized, then expand only after every slide is
      // paintable so unfinished slides do not create a deck-wide waiter set.
      viewerOptions.overscan = viewer.slideCount;
      viewer.relayout();
      return viewer.slideCount;
    };

    if (viewer.layoutComplete) {
      return { format: 'pptx', units: mountAllSlides(), unitLabel: 'slide' };
    }

    return {
      format: 'pptx',
      units: viewer.availableSlideCount,
      unitLabel: 'slide',
      finalUnits: viewer.waitUntilLayoutComplete().then(mountAllSlides),
    };
  }

  const host = scrollViewerHost(stage);
  let viewer!: DocxScrollViewer;
  const viewerOptions: DocxScrollViewerOptions = {
    gap: VIEWER_GAP,
    overscan: 0,
    enableTextSelection: true,
    enableZoom: true,
    zoomMin: MIN_SCALE,
    pageShadow: false,
    useGoogleFonts: true,
    comments: false,
    math,
    // Keep progressive pagination and painting off the UI thread so scrolling
    // remains responsive while later pages are still being prepared.
    mode: 'worker',
    progressiveLayout: true,
    onLayoutPartial: ({ availableUnits }) => {
      if (generation !== renderGeneration) return;
      // Grow the preview and start painting each published page immediately.
      // Waiting for the final pagination result made an idle reader see one
      // page for seconds even though later pages were already paintable.
      viewerOptions.overscan = availableUnits;
      viewer.relayout();
      onProgress?.({ format: 'docx', units: availableUnits, unitLabel: 'page' });
    },
    ...fullRenderers,
  };
  viewer = new DocxScrollViewer(host, viewerOptions);
  // As with PPTX, the viewer-owned width fit is the initial zoom contract for
  // Try Yours. When it falls below zoomMin, that fit stays reachable as the
  // effective floor for subsequent user-driven zoom operations.
  try {
    await viewer.load(buffer);
  } catch (error) {
    viewer.destroy();
    throw error;
  }
  if (generation !== renderGeneration) {
    viewer.destroy();
    throw new SupersededRenderError();
  }
  activeCleanup = () => viewer.destroy();

  const mountAllPages = (): number => {
    assertCurrentRender(generation);
    // Native Find needs every page's text layer in the DOM. Wait for the
    // authoritative count before expanding overscan so progressive load can
    // paint its opening window without immediately mounting unfinished pages.
    viewerOptions.overscan = viewer.pageCount;
    viewer.relayout();
    return viewer.pageCount;
  };

  if (viewer.layoutComplete) {
    return { format: 'docx', units: mountAllPages(), unitLabel: 'page' };
  }

  return {
    format: 'docx',
    units: viewer.pageCount,
    unitLabel: 'page',
    finalUnits: viewer.waitUntilLayoutComplete().then(mountAllPages),
  };
}

// Hot standby: warm each WASM engine on an idle tick
// so the user's first real file parses without paying the cold cost of fetching
// + compiling the parser binaries. Each `renderFile` spawns a fresh inline
// worker that re-fetches its `*_parser_bg.wasm`; pre-loading the bundled demo of
// every format primes the browser's HTTP/code cache for that binary, then the
// throwaway engines are released. Fire-and-forget, errors swallowed — a failed
// warm-up just means the first real parse is as slow as before, never broken.
let warmed = false;
export function prewarmEngines(): void {
  if (warmed || typeof window === 'undefined') return;
  warmed = true;
  // Respect Data Saver / metered connections — don't spend bandwidth warming.
  const conn = (navigator as Navigator & { connection?: { saveData?: boolean } }).connection;
  if (conn?.saveData) return;

  const base = import.meta.env.BASE_URL;
  const sample = (f: string) => `${base}samples/${f}`.replace(/([^:])\/\/+/g, '$1/');

  const run = (): void => {
    // Fonts are intentionally disabled here. Prewarming should prime only the
    // parser binaries; downloading fonts used by bundled samples would compete
    // with (and may be irrelevant to) the user's own file.
    void PptxPresentation.load(sample('sample-1.pptx'), {
      useGoogleFonts: false,
      mode: 'main',
      ...fullRenderers,
    })
      .then((d) => d.destroy())
      .catch(() => {});
    void DocxDocument.load(sample('sample-1.docx'), {
      useGoogleFonts: false,
      mode: 'main',
      ...fullRenderers,
    })
      .then((d) => d.destroy())
      .catch(() => {});
    // XlsxViewer needs a container; mount into a detached node never added to the
    // DOM, then dispose. The parse + one render warms the xlsx WASM engine.
    const host = document.createElement('div');
    const v = new XlsxViewer(host, {
      useGoogleFonts: false,
      mode: 'main',
      ...fullRenderers,
    });
    void v.load(sample('sample-1.xlsx')).then(() => v.destroy()).catch(() => v.destroy());
  };

  const ric = (window as Window & { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => void })
    .requestIdleCallback;
  if (ric) ric(run, { timeout: 2500 });
  else setTimeout(run, 600);
}
