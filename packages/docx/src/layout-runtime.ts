import { withVertFeatureCanvasScope } from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import type { ResolvedLocalFontMetric } from './layout/text.js';
import type { FontFamilyRoutes } from '@silurus/ooxml-core';
import { snapshotLocalMetrics } from './layout/text.js';
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
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>,
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
    readonly localMetrics?: Readonly<Record<string, ResolvedLocalFontMetric>>;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    readonly embeddedFaces?: readonly FontFace[];
    readonly providerRoutes?: FontFamilyRoutes;
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
  const localMetrics = snapshotLocalMetrics(options.localMetrics);
  const services = createProductionLayoutServices(source, {
    ...options,
    localMetrics,
    measureContext: context,
    verticalGlyphMeasurement,
  });
  attachLayoutSourceStore(services, source);
  attachBodyLayoutKernel(
    services,
    createConcreteBodyLayoutKernel(
      source,
      context,
      localMetrics,
    ),
  );
  return services;
}
