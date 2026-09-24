import { classifyCjkFont, type CjkLang, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { withVertFeatureCanvasScope } from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import type { LoadedEmbeddedFontRoute } from './embedded-fonts.js';
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

function createConcreteBodyLayoutKernel(
  source: LayoutSourceStore,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>>,
  cjkFallback?: CjkLang,
): BodyLayoutKernel {
  return createProductionBodyLayoutRuntime(
    source,
    measureContext,
    resolvedLocalFonts,
    cjkFallback,
  ).kernel;
}

export function createLayoutServices(
  input: DocxDocumentModel | LayoutSourceStore,
  options: {
    readonly localMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly useGoogleFonts?: boolean;
    readonly cjkFallback?: CjkLang;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    readonly embeddedRoutes?: readonly LoadedEmbeddedFontRoute[];
    readonly officeRoutes?: readonly OfficeFontFallbackRoute[];
    readonly googleFaces?: readonly FontFace[];
  } = {},
): LayoutServices {
  const source = isLayoutSourceStore(input) ? input : layoutSourceStore(input);
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
  const cjkFallback = source.fonts.scriptCjkLanguage
    ?? classifyCjkFont(source.fonts.majorFamily) ?? classifyCjkFont(source.fonts.minorFamily) ?? options.cjkFallback;
  const services = createProductionLayoutServices(source, {
    ...options,
    cjkFallback,
    localMetrics,
    fontMetrics: inputFontMetrics,
    measureContext: context,
    verticalGlyphMeasurement,
  });
  // Body layout and text measurement share one immutable resource snapshot,
  // including caller-supplied or decoded embedded font metrics.
  const fontMetrics = services.text.fontMetrics ?? inputFontMetrics;
  attachLayoutSourceStore(services, source);
  attachBodyLayoutKernel(
    services,
    createConcreteBodyLayoutKernel(
      source,
      context,
      fontMetrics,
      cjkFallback,
    ),
  );
  return services;
}
