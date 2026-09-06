import { withVertFeatureCanvasScope } from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import type { ResolvedFontMetric } from './layout/text.js';
import { snapshotFontMetrics } from './layout/text.js';
import type { MathLayoutResource } from './layout/resources.js';
import type { BodyLayoutKernel } from './layout/body-layout-kernel.js';
import type { LayoutServices } from './layout/types.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './layout/measurement-capabilities.js';
import {
  createProductionBodyLayoutRuntime,
} from './layout/production-body-layout.js';
import { createProductionLayoutServices } from './layout/production-services.js';
import {
  isLayoutSourceStore,
  type LayoutSourceStore,
} from './layout/layout-source-store.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import {
  planVerticalRunWithCapability,
  verticalRunInkExtraPx,
  verticalVertGlyphReachable,
} from './vertical-text.js';
import {
  attachBodyLayoutKernel,
  attachLayoutSourceStore,
} from './layout/runtime-state.js';
import {
  docxResolvedFontMetricCandidates,
  type DocxResolvedFontMetricCandidate,
} from './document-content.js';

function createConcreteBodyLayoutKernel(
  source: LayoutSourceStore,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>>,
): BodyLayoutKernel {
  return createProductionBodyLayoutRuntime(
    source,
    measureContext,
    resolvedLocalFonts,
  ).kernel;
}

export function createLayoutServices(
  input: DocxDocumentModel | LayoutSourceStore,
  options: {
    readonly localMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly useGoogleFonts?: boolean;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    readonly embeddedFaces?: readonly FontFace[];
    readonly googleFaces?: readonly FontFace[];
    readonly measureResolvedFontMetrics?: boolean;
    readonly resolvedFontMetricCandidates?: readonly DocxResolvedFontMetricCandidate[];
  } = {},
): LayoutServices {
  const source = isLayoutSourceStore(input) ? input : layoutSourceStore(input);
  const resolvedFontMetricCandidates = options.resolvedFontMetricCandidates
    ?? (isLayoutSourceStore(input)
      ? []
      : docxResolvedFontMetricCandidates(input, source.fontFamilyCharsets));
  // Main-thread layout must use an element-backed canvas when one is available:
  // OpenType `vert` is selected through the canvas element's CSS feature state,
  // and an OffscreenCanvas cannot prove or paint that feature route. Workers
  // have no `document`, so they retain the deterministic Offscreen fallback.
  const canvasContext = options.measureContext ?? (() => {
    if (typeof document !== 'undefined') {
      const mainThreadContext = document.createElement('canvas').getContext('2d');
      if (mainThreadContext !== null) return mainThreadContext;
    }
    return typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(1, 1).getContext('2d')
      : null;
  })();
  const context: MeasurementTextContext | null = canvasContext === null
    ? null
    : Object.freeze({
        get font() { return canvasContext.font; },
        set font(value: string) { canvasContext.font = value; },
        get letterSpacing() { return canvasContext.letterSpacing; },
        set letterSpacing(value: string) { canvasContext.letterSpacing = value; },
        get fontKerning() { return canvasContext.fontKerning; },
        set fontKerning(value: CanvasFontKerning) { canvasContext.fontKerning = value; },
        measureText(text: string) { return canvasContext.measureText(text); },
      });
  const canvasElement = canvasContext?.canvas as HTMLCanvasElement | undefined;
  const ownerCanvasConstructor =
    canvasElement?.ownerDocument?.defaultView?.HTMLCanvasElement;
  const hasDomVerticalProbe = canvasContext !== null
    && (
      (
        typeof ownerCanvasConstructor === 'function'
        && canvasElement instanceof ownerCanvasConstructor
      )
      || (
        typeof HTMLCanvasElement !== 'undefined'
        && canvasElement instanceof HTMLCanvasElement
      )
    );
  const verticalGlyphMeasurement: VerticalGlyphMeasurementService = Object.freeze({
    fingerprint: canvasContext === null
      ? 'vertical-glyph-measurement:deterministic-v1'
      : hasDomVerticalProbe
        ? 'vertical-glyph-measurement:dom-vert-probe-v2'
        : 'vertical-glyph-measurement:no-dom-vert-probe-v1',
    measureRunInkExtra(text: string): number {
      if (canvasContext === null) {
        throw new Error('Vertical glyph measurement requires a concrete text context');
      }
      return withVertFeatureCanvasScope(
        canvasContext,
        () => verticalRunInkExtraPx(canvasContext, text),
      );
    },
    planRun(input: Parameters<VerticalGlyphMeasurementService['planRun']>[0]) {
      if (canvasContext === null) {
        throw new Error('Vertical glyph planning requires a concrete text context');
      }
      return withVertFeatureCanvasScope(canvasContext, () => {
        const previousFont = canvasContext.font;
        const previousKerning = canvasContext.fontKerning;
        canvasContext.font = input.font;
        canvasContext.fontKerning = input.fontKerning;
        try {
          return planVerticalRunWithCapability(
            canvasContext,
            input.text,
            input.fontSizePt,
            input.letterSpacingPt,
            input.charScale,
            input.growTrRotateInk,
            (cp) => verticalVertGlyphReachable(canvasContext, cp),
            input.writingMode,
          );
        } finally {
          canvasContext.font = previousFont;
          canvasContext.fontKerning = previousKerning;
        }
      });
    },
  });
  const localMetrics = snapshotFontMetrics(options.localMetrics);
  const inputFontMetrics = snapshotFontMetrics({
    ...localMetrics,
    ...options.fontMetrics,
  });
  const services = createProductionLayoutServices(source, {
    ...options,
    resolvedFontMetricCandidates,
    localMetrics,
    fontMetrics: inputFontMetrics,
    measureContext: context,
    verticalGlyphMeasurement,
  });
  // The production service may add metrics proven from the concrete
  // Canvas-selected face. The body kernel (including empty paragraph-mark
  // measurement) must receive that same immutable snapshot rather than the
  // pre-probe caller input.
  const fontMetrics = services.text.fontMetrics ?? inputFontMetrics;
  attachLayoutSourceStore(services, source);
  attachBodyLayoutKernel(
    services,
    createConcreteBodyLayoutKernel(
      source,
      context,
      fontMetrics,
    ),
  );
  return services;
}
