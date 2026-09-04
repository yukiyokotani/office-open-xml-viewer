import type { DimOptions } from './types';
import type { PptxTextRunInfo } from './renderer';
import type { PresentationPreflight } from './presentation-preflight';
import type {
  NormalizedOoxmlResourcePolicy,
  PullSessionIdentity,
  WorkerErrorPayload,
  WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import type {
  PptxElementContextOptions,
  PptxElementContext,
  PptxElementBounds,
  PptxSlidePoint,
} from './element-selection';

/** Canonical compact payload emitted by Rust `presentation_bootstrap()`. */
export interface PresentationBootstrap {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly defaultTextColor: string | null;
  readonly majorFont: string | null;
  readonly minorFont: string | null;
  readonly hlinkColor: string | null;
  readonly folHlinkColor: string | null;
  readonly embeddedFonts: readonly PptxEmbeddedFontRef[];
  readonly slides: readonly PresentationBootstrapSlide[];
}

export type PptxEmbeddedFontStyle = 'regular' | 'bold' | 'italic' | 'boldItalic';

/** Compact ECMA-376 Part 1 §19.2.1.9 embedded-font reference. */
export interface PptxEmbeddedFontRef {
  readonly fontName: string;
  readonly style: PptxEmbeddedFontStyle;
  readonly partPath: string;
  readonly contentType: 'application/x-font-ttf' | 'application/x-fontdata';
}

export interface PresentationBootstrapSlide {
  readonly index: number;
  readonly partName?: string;
}

export type PptxWorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | {
      kind: 'parse';
      id: number;
      buffer: ArrayBuffer;
      resourcePolicy: NormalizedOoxmlResourcePolicy;
      progressiveLayout?: boolean;
    }
  | ({
      kind: 'openSlideSession';
      id: number;
      slideIndex: number;
    } & PullSessionIdentity<number>)
  | { kind: 'finishPresentationPreflight'; id: number }
  | { kind: 'extractMedia'; id: number; path: string }
  | { kind: 'extractImage'; id: number; path: string }
  | { kind: 'extractFont'; id: number; path: string }
  | { kind: 'resourceUsage'; id: number }
  | { kind: 'toMarkdown'; id: number };

export type PptxWorkerResponse =
  | { kind: 'presentationOpened'; id: number; bootstrap: PresentationBootstrap }
  | ({ kind: 'slideSessionOpened'; id: number } & PullSessionIdentity<number>)
  | { kind: 'presentationPreflightReady'; id: number; preflight: PresentationPreflight }
  | { kind: 'mediaExtracted'; id: number; bytes: ArrayBuffer }
  | { kind: 'imageExtracted'; id: number; bytes: ArrayBuffer }
  | { kind: 'fontExtracted'; id: number; bytes: ArrayBuffer }
  | { kind: 'resourceUsage'; id: number; usage: OoxmlResourceUsageSnapshot }
  | { kind: 'markdownRendered'; id: number; markdown: string }
  | ({ kind: 'error'; id: number } & WorkerErrorPayload);

// The render worker owns both the cursor and bounded slide repository. It
// returns the exact same compact preflight contract as the main-mode worker,
// while complete Slide models never cross into Window.
export type RenderWorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | { kind: 'continuePresentationPreflight'; forId: number; availableSlides: number }
  | {
      kind: 'parse';
      id: number;
      buffer: ArrayBuffer;
      resourcePolicy: NormalizedOoxmlResourcePolicy;
      useGoogleFonts?: boolean;
      renderers?: WorkerRendererDescriptors;
      progressiveLayout?: boolean;
    }
  | { kind: 'extractMedia'; id: number; path: string }
  | { kind: 'extractImage'; id: number; path: string }
  | { kind: 'extractFont'; id: number; path: string }
  | { kind: 'resourceUsage'; id: number }
  | { kind: 'toMarkdown'; id: number }
  | {
      kind: 'renderSlide';
      id: number;
      slideIndex: number;
      width: number;
      dpr: number;
      imageResources?: import('@silurus/ooxml-core').ImageResourceOptions;
      skipMediaControls?: boolean;
      dim?: DimOptions;
    }
  | { kind: 'collectRuns'; id: number; slideIndex: number; width: number }
  | {
      kind: 'hitTestElement';
      id: number;
      slideIndex: number;
      point: PptxSlidePoint;
      options: PptxElementContextOptions;
    }
  | {
      kind: 'resolveElementBounds';
      id: number;
      slideIndex: number;
      elementIds: readonly string[];
    };

export type RenderWorkerResponse =
  | Exclude<
      PptxWorkerResponse,
      { kind: 'presentationOpened' | 'slideSessionOpened' | 'presentationPreflightReady' }
    >
  | {
      kind: 'presentationReady';
      id: number;
      preflight: PresentationPreflight;
      usage?: OoxmlResourceUsageSnapshot;
    }
  | {
      kind: 'presentationLayoutPartial';
      forId: number;
      bootstrap?: PresentationBootstrap;
      availableSlides: number;
      slide: PresentationPreflight['slides'][number];
      fontPreloadNames: PresentationPreflight['fontPreloadNames'];
      usage?: OoxmlResourceUsageSnapshot;
    }
  | { kind: 'slideRendered'; id: number; bitmap: ImageBitmap; runs: PptxTextRunInfo[] }
  | { kind: 'runsCollected'; id: number; runs: PptxTextRunInfo[] }
  | { kind: 'elementHit'; id: number; context: PptxElementContext | null }
  | { kind: 'elementBoundsResolved'; id: number; bounds: readonly PptxElementBounds[] };
