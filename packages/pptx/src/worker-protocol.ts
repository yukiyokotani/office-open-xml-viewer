import type { DimOptions, MediaElement, WorkerResponse } from './types';
import type { PptxTextRunInfo } from './renderer';

/** Lightweight summary returned by the render worker's `parse` — everything
 *  the main-thread proxy needs for its synchronous getters. The full model
 *  stays in the worker. */
export interface PresentationMeta {
  slideCount: number;
  slideWidth: number;
  slideHeight: number;
  majorFont: string | null;
  minorFont: string | null;
  /** Speaker-notes text per slide (same contract as PptxPresentation.getNotes). */
  notes: (string | null)[];
  /** Media elements per slide (geometry + paths), for main-thread playback
   *  overlays. Small: a handful of plain objects per slide. */
  mediaElements: MediaElement[][];
  /** `Slide.hidden` per slide (`<p:sld show="0">`, §19.3.1.38). */
  hidden: boolean[];
  /** `Slide.partName` per slide (normalized OPC part name, `sldIdLst` order).
   *  Lets the main-thread proxy build the `partName → index` map that resolves
   *  an internal hyperlink slide jump in worker mode, mirroring `hidden`/`notes`.
   *  Entries are `undefined` only for a slide whose part path wasn't recorded. */
  partNames: (string | undefined)[];
}

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `kind`-based narrowing at use sites. The
// `init` / `extractMedia` arms are copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | { kind: 'extractMedia'; id: number; path: string }
  | { kind: 'extractImage'; id: number; path: string }
  | { kind: 'toMarkdown'; id: number }
  | { kind: 'parse'; id: number; buffer: ArrayBuffer; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number; useGoogleFonts?: boolean }
  | { kind: 'renderSlide'; id: number; slideIndex: number; width: number; dpr: number; skipMediaControls?: boolean; dim?: DimOptions }
  // IX6 — collect a slide's text-run geometry WITHOUT transferring a bitmap. The
  // find controller scans every slide for its runs; a bitmap per slide would be
  // wasted work + transfer for slides the user never looks at.
  | { kind: 'collectRuns'; id: number; slideIndex: number; width: number };

export type RenderWorkerResponse =
  | Exclude<WorkerResponse, { kind: 'parsed' }>
  | { kind: 'parsedMeta'; id: number; meta: PresentationMeta }
  // IX6 — the render worker collects each rendered slide's `onTextRun` geometry
  // (a plain, structured-clone-safe `PptxTextRunInfo[]`) and ships it beside the
  // bitmap, so the main thread can build the text-selection / find-highlight
  // overlay on the SAME code path as main mode (no second render).
  | { kind: 'slideRendered'; id: number; bitmap: ImageBitmap; runs: PptxTextRunInfo[] }
  | { kind: 'runsCollected'; id: number; runs: PptxTextRunInfo[] };
