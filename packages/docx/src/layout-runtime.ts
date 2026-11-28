import { restoreCanvasLanguage, selectCanvasLanguage } from '@silurus/ooxml-core/internal/canvas-language';
import { classifyCjkFont, type CjkLang, type NativeRouteLineRequest, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import {
  activeFontSet,
  createReferenceAdvanceContext,
  createTextAdvancePlanner,
  isHTMLCanvas,
  withReferenceAdvances,
  withVertFeatureCanvasScope,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import type { FontSubstitution } from '@silurus/ooxml-core/internal/font-sources';
import type { LoadedEmbeddedFontRoute } from './embedded-fonts.js';
import type { ResolvedFontMetric } from './layout/text.js';
import { snapshotFontMetrics } from './layout/text.js';
import type { NativeRouteLineMetric } from './layout/font-service.js';
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
import { createProductionLayoutServices, createProductionTextLayoutService } from './layout/production-services.js';
import {
  admitDocxNativeRouteLines,
  docxNativeRouteRequests,
  measureDocxNativeRouteLines,
} from './native-route-line-metrics.js';
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
  attachTextAdvancePlanner,
} from './layout/runtime-state.js';

/** Reference-advance probe clusters (core reference-advances.ts): Latin
 * letters and a digit of differing widths, plus the kana and ideograph probes
 * layout already measures. All are in line-layout.ts LAYOUT_SYNTHESIZED_SCALARS, so
 * finite font demand covers them and probing loads no extra font subset. */
export const TEXT_ADVANCE_PROBE_CLUSTERS = Object.freeze(['a', 'e', 'm', 'W', '0', '\u3042', '\u4E00']);

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

function layoutCjkFallback(source: LayoutSourceStore, requested: CjkLang | undefined): CjkLang | undefined {
  return source.fonts.scriptCjkLanguage
    ?? classifyCjkFont(source.fonts.majorFamily) ?? classifyCjkFont(source.fonts.minorFamily) ?? requested;
}

/** Every native route family list that {@link createLayoutServices} can
 * request for this document (native-route-line-metrics.ts), resolved through
 * the production text service itself. A native route's family list depends
 * only on the authored family, the document's family facts, language, script
 * slot, generic family, CJK fallback and the optional-substitute policy, never
 * on which resources load: a loaded resource only turns a tuple into a
 * registered route. So the service is built without any resource inventory,
 * which makes every tuple native, and the result is the complete set of
 * native lists before any font is loaded. Lists that later resolve to a
 * registered or installed route are measured but never looked up. */
export function nativeRouteLineRequests(
  input: DocxDocumentModel | LayoutSourceStore,
  model: DocxDocumentModel,
  options: Readonly<{ fontSubstitutes?: readonly FontSubstitution[]; cjkFallback?: CjkLang }> = {},
): NativeRouteLineRequest[] {
  const source = isLayoutSourceStore(input) ? input : layoutSourceStore(input);
  const text = createProductionTextLayoutService(source, {
    fontSubstitutes: options.fontSubstitutes,
    cjkFallback: layoutCjkFallback(source, options.cjkFallback),
    measureContext: null,
    fontSet: null,
  });
  return docxNativeRouteRequests(model, text);
}

/** Main-thread preparation: enumerate, measure and admit in this realm. The
 * returned sentinel faces belong in the document's release list; they judge
 * glyph ownership while this realm lays out. */
export async function prepareMainThreadNativeRouteMetrics(
  input: LayoutSourceStore,
  model: DocxDocumentModel,
  options: Readonly<{ fontSubstitutes?: readonly FontSubstitution[]; cjkFallback?: CjkLang }> = {},
): Promise<{ nativeRouteMetrics?: Record<string, NativeRouteLineMetric>; faces: FontFace[] }> {
  const measured = await measureDocxNativeRouteLines(nativeRouteLineRequests(input, model, options), true);
  return measured.faces.length > 0
    ? { nativeRouteMetrics: admitDocxNativeRouteLines(measured.lines, activeFontSet()), faces: measured.faces }
    : { faces: [] };
}

export function createLayoutServices(
  input: DocxDocumentModel | LayoutSourceStore,
  options: {
    readonly localMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
    readonly fontSubstitutes?: readonly FontSubstitution[];
    readonly cjkFallback?: CjkLang;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    /** The run-language oracle already built on {@link measureContext} (a
     * render worker decides admission with it); otherwise the text service
     * builds its own. */
    readonly namedFamilyOracle?: import('@silurus/ooxml-core').NamedFamilyOracle;
    /** The factory that made {@link measureContext}'s canvas, when the caller
     * owns one (a Node session): symbol colour probes allocate their scratch
     * canvas through it, on the backend that measures and paints. */
    readonly createCanvas?: import('@silurus/ooxml-core').SymbolProbeCanvasFactory;
    readonly embeddedRoutes?: readonly LoadedEmbeddedFontRoute[];
    readonly officeRoutes?: readonly OfficeFontFallbackRoute[];
    readonly sourceFonts?: import('./source-font-metrics.js').DocxSourceFonts;
    /** Measured served/application face lines (served-face-lines.ts). */
    readonly servedFaceLines?: import('./served-face-lines.js').PreparedServedFaceLines;
    readonly installedSubstituteFamilies?: readonly string[];
    readonly installedFamilies?: readonly string[];
    readonly allowFootnoteContinuation?: boolean;
    /** Admitted native route lines (native-route-line-metrics.ts). */
    readonly nativeRouteMetrics?: Readonly<Record<string, NativeRouteLineMetric>>;
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
  // Canvas may grid-fit small sizes (core reference-advances.ts): one planner
  // per measuring context decides, per font shorthand, whether text is
  // measured at its own size or at the reference em under a linear scale.
  // Every layout consumer measures through the same view, and retained paint
  // reads the same planner's decisions, so layout and glyphs agree.
  const advancePlanner = canvasContext === null
    ? null
    : createTextAdvancePlanner(canvasContext, { probeClusters: TEXT_ADVANCE_PROBE_CLUSTERS });
  const measured = canvasContext === null || advancePlanner === null
    ? null
    : createReferenceAdvanceContext(canvasContext, advancePlanner);
  // The same view, typed as the concrete context, for vertical planning.
  const verticalContext = canvasContext === null || measured === null
    ? null
    : withReferenceAdvances(canvasContext, measured);
  let hasLang: boolean | undefined;
  const engineHasLang = () => hasLang ??= canvasContext !== null && 'lang' in canvasContext;
  const context: MeasurementTextContext | null = canvasContext === null || measured === null
    ? null
    : Object.freeze(Object.defineProperties({
        get textAlign() { return canvasContext.textAlign; },
        set textAlign(value: CanvasTextAlign | undefined) { if (value !== undefined) canvasContext.textAlign = value; },
        get direction() { return canvasContext.direction; },
        set direction(value: CanvasDirection | undefined) { if (value !== undefined) canvasContext.direction = value; },
        get font() { return measured.font; },
        set font(value: string) { measured.font = value; },
        get letterSpacing() { return measured.letterSpacing; },
        set letterSpacing(value: string) { measured.letterSpacing = value; },
        get fontKerning() { return canvasContext.fontKerning; },
        set fontKerning(value: CanvasFontKerning) { canvasContext.fontKerning = value; },
        measureText(text: string) { return measured.measureText(text); },
      }, {
        // Canvas `lang` (language-specific glyphs), forwarded only where the
        // engine has it. The engine is asked on first use, never at creation:
        // untagged text selects no language, so its measurement reads and
        // writes none of the caller context's language state. An engine
        // without the attribute reads as no language and ignores writes.
        lang: {
          get: () => engineHasLang() ? (canvasContext as { lang: string }).lang : undefined,
          set: (value: string) => { if (engineHasLang()) (canvasContext as { lang: string }).lang = value; },
          enumerable: true,
        },
      }));
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
        ? 'vertical-glyph-measurement:dom-vert-probe-v3'
        : 'vertical-glyph-measurement:no-dom-vert-probe-v2',
    measureRunInkExtra(text: string, lang?: string): number {
      if (canvasContext === null || verticalContext === null) {
        throw new Error('Vertical glyph measurement requires a concrete text context');
      }
      return withVertFeatureCanvasScope(
        canvasContext,
        () => verticalRunInkExtraPx(verticalContext, text, lang),
      );
    },
    planRun(input: Parameters<VerticalGlyphMeasurementService['planRun']>[0]) {
      if (canvasContext === null || verticalContext === null) {
        throw new Error('Vertical glyph planning requires a concrete text context');
      }
      // The vertical planner measures through the same reference-advance view
      // as horizontal acquisition (its `vert` toggles re-assign the logical
      // font, which re-applies the planned reference em).
      const view = verticalContext;
      return withVertFeatureCanvasScope(canvasContext, () => {
        const previousFont = view.font;
        const previousKerning = canvasContext.fontKerning;
        view.font = input.font;
        canvasContext.fontKerning = input.fontKerning;
        const previousLang = selectCanvasLanguage(canvasContext, input.lang);
        try {
          return planVerticalRunWithCapability(
            view,
            input.text,
            input.fontSizePt,
            input.letterSpacingPt,
            input.charScale,
            input.growTrRotateInk,
            (cp) => verticalVertGlyphReachable(view, cp, input.lang),
            input.writingMode,
            input.characterGridDeltaPt,
            input.characterAdjustmentPt,
          );
        } finally {
          restoreCanvasLanguage(canvasContext, previousLang);
          view.font = previousFont;
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
  const cjkFallback = layoutCjkFallback(source, options.cjkFallback);
  const services = createProductionLayoutServices(source, {
    ...options,
    cjkFallback,
    localMetrics,
    fontMetrics: inputFontMetrics,
    measureContext: context,
    ...(advancePlanner ? { textAdvancePlanner: advancePlanner } : {}),
    symbolProbeOwner: canvasContext,
    ...(options.namedFamilyOracle ? { namedFamilyOracle: options.namedFamilyOracle } : {}),
    ...(options.createCanvas ? { symbolProbeCanvas: options.createCanvas } : {}),
    // A caller canvas may belong to a popup/iframe. CSS face admission must
    // inspect the FontFaceSet that also shapes this context's glyphs.
    fontSet: (isHTMLCanvas(canvasElement) ? canvasElement.ownerDocument?.fonts : undefined)
      ?? activeFontSet(),
    verticalGlyphMeasurement,
  });
  // Body layout and text measurement share one immutable resource snapshot,
  // including caller-supplied or decoded embedded font metrics.
  const fontMetrics = services.text.fontMetrics ?? inputFontMetrics;
  attachLayoutSourceStore(services, source);
  if (advancePlanner !== null) attachTextAdvancePlanner(services, advancePlanner);
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
