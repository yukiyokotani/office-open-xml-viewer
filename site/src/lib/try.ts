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
import { XlsxViewer } from '@silurus/ooxml-xlsx';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

// Opt-in OMML equation engine — enabled here so user-supplied docx/pptx with
// equations render. (In the published library this is `@silurus/ooxml/math`.)
const advancedChartRenderers = { threeD, regionMap, chartEx };

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

export async function renderFile(stage: HTMLElement, file: File): Promise<RenderResult> {
  const generation = ++renderGeneration;
  // Any new selection — supported or not — supersedes and releases the current
  // viewer before validation. The page hides the panel on validation failure, so
  // retaining its media/worker resources would be both invisible and leaked.
  activeCleanup?.();
  activeCleanup = null;

  const ext = file.name.split('.').pop()?.toLowerCase();
  if (ext !== 'docx' && ext !== 'xlsx' && ext !== 'pptx') {
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
      comments: true,
      math,
      ...advancedChartRenderers,
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

  if (ext === 'pptx') {
    const host = scrollViewerHost(stage);
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
      comments: true,
      math,
      // Keep progressive preflight and paint off the UI thread while the user
      // scrolls through slides that are still becoming available.
      mode: 'worker',
      progressiveLayout: true,
      ...advancedChartRenderers,
    };
    const viewer = new PptxScrollViewer(host, viewerOptions);
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
  const viewerOptions: DocxScrollViewerOptions = {
    gap: VIEWER_GAP,
    overscan: 0,
    enableTextSelection: true,
    enableZoom: true,
    zoomMin: MIN_SCALE,
    pageShadow: false,
    useGoogleFonts: true,
    comments: true,
    math,
    // Keep progressive pagination and painting off the UI thread so scrolling
    // remains responsive while later pages are still being prepared.
    mode: 'worker',
    progressiveLayout: true,
    ...advancedChartRenderers,
  };
  const viewer = new DocxScrollViewer(host, viewerOptions);
  // As with PPTX, the viewer-owned width fit is the initial zoom contract for
  // Try Yours. It can go below zoomMin when necessary to admit a wide page;
  // zoomMin remains the floor for subsequent user-driven zoom operations.
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
      ...advancedChartRenderers,
    })
      .then((d) => d.destroy())
      .catch(() => {});
    void DocxDocument.load(sample('sample-1.docx'), {
      useGoogleFonts: false,
      mode: 'main',
      ...advancedChartRenderers,
    })
      .then((d) => d.destroy())
      .catch(() => {});
    // XlsxViewer needs a container; mount into a detached node never added to the
    // DOM, then dispose. The parse + one render warms the xlsx WASM engine.
    const host = document.createElement('div');
    const v = new XlsxViewer(host, {
      useGoogleFonts: false,
      mode: 'main',
      ...advancedChartRenderers,
    });
    void v.load(sample('sample-1.xlsx')).then(() => v.destroy()).catch(() => v.destroy());
  };

  const ric = (window as Window & { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => void })
    .requestIdleCallback;
  if (ric) ric(run, { timeout: 2500 });
  else setTimeout(run, 600);
}
