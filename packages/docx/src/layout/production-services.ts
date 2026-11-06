import { acquireCanvasGlyphContours } from '../native-glyph-contours.js';
import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import { type CjkLang } from '@silurus/ooxml-core';
import {
  activeFontSet,
  canvasFontString,
  fontFaceWeightCovers,
} from '@silurus/ooxml-core';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import type { FontSubstituteScript, OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { DOCX_GOOGLE_FONTS } from '../google-font-registry.js';
import type { LoadedEmbeddedFontRoute } from '../embedded-fonts.js';
import { normalizeFontFamilyUncached } from '../line-layout.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import { createFontResolver, type FontInventoryFace, type NativeRouteLineMetric } from './font-service.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './measurement-capabilities.js';
import {
  createImageMetadataService,
  createMathMetadataService,
  mathResourceKey,
  type MathLayoutResource,
} from './resources.js';
import {
  attachPaintResourceRegistry,
  attachPrivateResourceLookup,
  attachVerticalGlyphMeasurementService,
} from './runtime-state.js';
import {
  classifyDocxFontGeneric,
  createTextLayoutService,
  nativeVisiblePrefixMatches,
  snapshotFontMetrics,
  type GlyphMeasureRequest,
  type TextLayoutService,
} from './text.js';
import type { LayoutServices } from './types.js';
import { ownsGlyphs, sameGlyphRaster } from '../glyph-raster-proof.js';

export interface LoadedFontFaceRecord {
  readonly family: string;
  readonly status: string;
  readonly style: string;
  readonly weight: string;
}

export interface ProductionLayoutServiceOptions {
  readonly allowFootnoteContinuation?: boolean;
  readonly localMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly useGoogleFonts?: boolean;
  readonly cjkFallback?: CjkLang;
  readonly mathResources?: readonly MathLayoutResource[];
  readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
  readonly measureContext: MeasurementTextContext | null;
  /** FontFaceSet of the measuring canvas, including foreign DOM realms. */
  readonly fontSet?: FontFaceSet | null;
  readonly verticalGlyphMeasurement: VerticalGlyphMeasurementService;
  /** Exact registrations returned by the DOCX loader. */
  readonly embeddedRoutes?: readonly LoadedEmbeddedFontRoute[];
  /** Exact local registrations, scoped to this FontFaceSet. */
  readonly officeRoutes?: readonly OfficeFontFallbackRoute[];
  readonly googleFaces?: readonly LoadedFontFaceRecord[];
  /** Normalized families whose authored face is installed; they are never
   * routed to a different-family Google substitute (`docxGoogleFontPlan`). */
  readonly installedSubstituteFamilies?: readonly string[];
  /** Normalized authored families the family presence probe found installed. */
  readonly installedFamilies?: readonly string[];
  /** Admitted native route lines (native-route-line-metrics.ts). */
  readonly nativeRouteMetrics?: Readonly<Record<string, NativeRouteLineMetric>>;
}

/** Registry families with a script-scoped visual substitute (core
 * substitute-script.ts), mapped to every substitute family of that script. An
 * installed authored face is never substituted, so it is not scoped either.
 * This registry delimits proof, not slot overrides: the resolver permits an
 * override only when a loaded scoped substitute wins the exact resource tuple,
 * after embedded, local and authored CSS inventory precedence. */
function docxScriptScopedFamilies(
  installed: readonly string[],
): Record<string, { script: FontSubstituteScript; substituteFamilies: string[] }> {
  const scoped = Object.entries(DOCX_GOOGLE_FONTS).filter(([, entry]) => entry.script !== undefined);
  return Object.fromEntries(scoped
    .filter(([key]) => !installed.includes(key))
    .map(([key, entry]) => [key, {
      script: entry.script!,
      substituteFamilies: [...new Set(scoped
        .filter(([, other]) => other.script === entry.script && other.loadFamily)
        .map(([, other]) => other.loadFamily!))],
    }]));
}

/** The production text service alone: font inventory, resolver and Canvas
 * measurer. Native route enumeration resolves through this same service, so
 * the routes it measures are exactly the routes layout requests. */
export function createProductionTextLayoutService(
  source: LayoutSourceStore,
  options: Omit<ProductionLayoutServiceOptions, 'verticalGlyphMeasurement'>,
): TextLayoutService {
  const fontFamilyCharsets = Object.freeze(Object.fromEntries(
    Object.entries(source.fontFamilyCharsets)
      .map(([family, charset]) => [family.trim().toLowerCase(), charset]),
  ));
  const displayFaceFamily = (family: string): string => family
    .trim()
    .replace(/^(['"])(.*)\1$/, '$2');
  const normalizedFaceFamily = (family: string): string => displayFaceFamily(family)
    .toLowerCase();
  const loadedFaceStyle = (face: LoadedFontFaceRecord): 'normal' | 'italic' | null => {
    const style = face.style.trim().toLowerCase();
    return style === 'normal' || style === 'italic' ? style : null;
  };
  const loadedFaceWeights = (face: LoadedFontFaceRecord): number[] => {
    const weight = face.weight.trim().toLowerCase();
    if (weight === 'normal') return [400];
    if (weight === 'bold') return [700];
    if (/^\d+$/.test(weight)) {
      const numeric = Number(weight);
      return numeric >= 100 && numeric <= 900 ? [numeric] : [];
    }
    // WML exposes regular/bold tuples. A loaded CSS variable face may cover
    // both; use the same descriptor rule as the local-face preflight.
    return [400, 700].filter((tupleWeight) => fontFaceWeightCovers(weight, tupleWeight));
  };
  const loadedFaces = (faces: readonly LoadedFontFaceRecord[]) => faces.flatMap((face) => {
    if (face.status !== 'loaded') return [];
    const style = loadedFaceStyle(face);
    return style == null ? [] : loadedFaceWeights(face).map((weight) => ({
      family: normalizedFaceFamily(face.family),
      displayFamily: displayFaceFamily(face.family),
      weight,
      style,
    }));
  });
  const inventory: FontInventoryFace[] = (options.embeddedRoutes ?? []).map((route) => ({
      requestedFamily: route.requestedFamily,
      resolvedFamily: route.resolvedFamily,
      resourceIdentity: route.resourceIdentity,
      source: 'embedded' as const,
      weight: route.weight,
      style: route.style,
    }));
  // Caller resources are optional substitutes. An authored embedded face wins
  // the same tuple even when its loader supplies no numeric metric; do not
  // overwrite that face's measured/default metrics with caller alias geometry.
  const embeddedTuples = new Set(inventory.map((face) =>
    `${normalizedFaceFamily(face.requestedFamily)}:${face.weight}:${face.style}`));
  const unshadowed = (metrics: Readonly<Record<string, ResolvedFontMetric>> = {}) =>
    Object.fromEntries(Object.entries(metrics).filter(([key, metric]) =>
      !metric.sourceIdentity?.startsWith('provided-sfnt:')
      || !embeddedTuples.has(`${normalizedFaceFamily(metric.requestedFamily ?? key)}:${metric.weight ?? 400}:${metric.style ?? 'normal'}`)));
  const localMetrics = snapshotFontMetrics(unshadowed(options.localMetrics));
  const localTuples = new Set(Object.entries(localMetrics).map(([key, metric]) =>
    `${normalizedFaceFamily(metric.requestedFamily ?? key)}:${metric.weight ?? 400}:${metric.style ?? 'normal'}`));
  // A document-embedded face and an application-provided exact resource outrank
  // the library's offline fallback for the same authored tuple.
  const officeRoutes = (options.officeRoutes ?? []).filter((route) => {
    const tuple = `${normalizedFaceFamily(route.requestedFamily)}:${route.weight}:${route.style}`;
    return !embeddedTuples.has(tuple) && !localTuples.has(tuple);
  });
  const officeMetrics = Object.fromEntries(officeRoutes.flatMap((route) => {
    if (route.metric.lineHeightRatio == null) return [];
    const requested = normalizedFaceFamily(route.requestedFamily);
    const key = route.weight === 400 && route.style === 'normal'
      ? requested : `${requested}:${route.weight}:${route.style}`;
    return [[key, route.metric]];
  }));
  const fontMetrics = snapshotFontMetrics({
    ...localMetrics,
    ...officeMetrics,
    ...unshadowed(options.fontMetrics),
  });
  for (const route of officeRoutes) {
    inventory.push({
      requestedFamily: route.requestedFamily,
      resolvedFamily: route.family,
      source: route.source,
      resourceIdentity: route.resourceIdentity,
      weight: route.weight,
      style: route.style,
      ...(route.coverageProbeFamily ? { coverageProbeFamily: route.coverageProbeFamily } : {}),
    });
  }
  for (const [requestedFamily, metric] of Object.entries(localMetrics)) {
    inventory.push({
      requestedFamily: metric.requestedFamily ?? requestedFamily,
      resolvedFamily: metric.family,
      source: 'local',
      ...(metric.sourceIdentity === undefined ? {} : { resourceIdentity: metric.sourceIdentity }),
      weight: metric.weight ?? 400,
      style: metric.style ?? 'normal',
    });
  }
  if (options.useGoogleFonts) {
    const successfulGoogle = loadedFaces(options.googleFaces ?? []);
    const seen = new Set<string>();
    for (const name of source.fonts.preloadNames) {
      if (!name) continue;
      const key = name.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      const entry = ownValue(DOCX_GOOGLE_FONTS, key);
      const resolvedFamily = entry?.loadFamily ?? name;
      if (!entry) continue;
      if (normalizedFaceFamily(resolvedFamily) !== normalizedFaceFamily(name)
        && options.installedSubstituteFamilies?.includes(key)) continue;
      for (const loaded of successfulGoogle.filter(
        (face) => face.family === normalizedFaceFamily(resolvedFamily),
      )) {
        inventory.push({
          requestedFamily: name,
          resolvedFamily: loaded.displayFamily,
          source: normalizedFaceFamily(resolvedFamily) === normalizedFaceFamily(name)
            ? 'google' : 'substitute',
          ...(entry.script === undefined ? {} : { script: entry.script }),
          weight: loaded.weight,
          style: loaded.style,
        });
      }
    }
  }
  const context = options.measureContext;
  const routedFontFamilies = [...new Set([
    ...Object.keys(source.fonts.familyClasses),
    ...Object.keys(source.fonts.familyPitches),
    ...source.fonts.renderedFamilies,
    ...(source.fonts.majorFamily ? [source.fonts.majorFamily] : []),
    ...(source.fonts.minorFamily ? [source.fonts.minorFamily] : []),
  ])];
  // An application-owned @font-face with the authored family is a real Canvas
  // route, even though the library cannot read its font bytes. Snapshot loaded
  // tuples at service creation so line layout uses Canvas's own line box rather
  // than the pinned reference for a different version of the same family.
  // A face loaded after this snapshot needs a new layout service/pagination.
  const requestedNames = new Set(routedFontFamilies.map(normalizedFaceFamily));
  const fontSet = options.fontSet === undefined ? activeFontSet() : options.fontSet;
  // Any loaded application face declaring an authored family shadows the
  // system family for every tuple, whatever its descriptors (oblique, weight
  // 950...), even ones the exact-tuple inventory below cannot represent.
  const shadowedFamilies = new Set<string>();
  if (fontSet && typeof fontSet[Symbol.iterator] === 'function') {
    for (const face of fontSet) {
      if (!requestedNames.has(normalizedFaceFamily(face.family))) continue;
      if (face.status === 'loaded') shadowedFamilies.add(normalizedFaceFamily(face.family));
      for (const loaded of loadedFaces([face])) {
        inventory.push({
          requestedFamily: loaded.displayFamily,
          resolvedFamily: loaded.displayFamily,
          source: 'css',
          weight: loaded.weight,
          style: loaded.style,
        });
      }
    }
  }
  // An explicitly directed request uses the same metric implementation as
  // ordinary acquisition, but only a context exposing Canvas direction opts in.
  // Unknown structural/custom contexts retain their ordinary contract.
  const supportsDirectedContext = context !== null
    && ['ltr', 'rtl', 'inherit'].includes(context.direction ?? '')
    && ['start', 'end', 'left', 'right', 'center'].includes(context.textAlign ?? '');
  const measure = (request: Readonly<GlyphMeasureRequest>) => {
    if (!context) return {
      advancePt: [...request.text].length * request.fontSizePt * 0.5,
      ascentPt: request.fontSizePt * 0.8,
      descentPt: request.fontSizePt * 0.2,
    };
    const previousFont = context.font;
    const previousLetterSpacing = context.letterSpacing;
    const previousKerning = context.fontKerning;
    const previousDirection = request.baseDirection !== undefined ? context.direction : undefined;
    const previousTextAlign = request.baseDirection !== undefined ? context.textAlign : undefined;
    try {
      if (request.baseDirection !== undefined) {
        context.direction = request.baseDirection;
        // Native ink bounds must use the same physical left origin as retained
        // point geometry. Canvas start + RTL would instead use a right anchor.
        context.textAlign = 'left';
        if (context.direction !== request.baseDirection || context.textAlign !== 'left') {
          throw new TypeError('Canvas did not honor directed measurement and left anchor');
        }
      }
      context.font = canvasFontString(
        request.fontRoute,
        request.fontSizePt,
        request.weight,
        request.style,
      );
      context.letterSpacing = `${request.letterSpacingPt}px`;
      if (request.kerning != null) context.fontKerning = request.kerning ? 'normal' : 'none';
      const metrics = context.measureText(request.text);
      const horizontalInkBoundsAreTight =
        Number.isFinite(metrics.actualBoundingBoxLeft)
        && Number.isFinite(metrics.actualBoundingBoxRight);
      // Retain the historical full-advance fallback for consumers that need
      // a stable ink box (ruby, decoration, hit geometry), but label whether
      // the horizontal edges are genuinely tight. Whitespace-trimming
      // consumers must not infer sidebearings from the fallback box.
      const inkBounds = {
        xMinPt: horizontalInkBoundsAreTight
          ? -metrics.actualBoundingBoxLeft : 0,
        xMaxPt: horizontalInkBoundsAreTight
          ? metrics.actualBoundingBoxRight : metrics.width,
        ascentPt: metrics.actualBoundingBoxAscent,
        descentPt: metrics.actualBoundingBoxDescent,
      };
      return {
        advancePt: metrics.width,
        ...(request.baseDirection !== undefined ? { nativeBaseDirection: request.baseDirection } : {}),
        ascentPt: metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent ?? 0,
        descentPt: metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent ?? 0,
        ...(Object.values(inkBounds).every(Number.isFinite) ? {
          inkBounds,
          ...(horizontalInkBoundsAreTight ? { horizontalInkBoundsAreTight: true } : {}),
        } : {}),
      };
    } finally {
      if (request.baseDirection !== undefined) {
        context.direction = previousDirection;
        context.textAlign = previousTextAlign;
      }
      context.font = previousFont;
      context.letterSpacing = previousLetterSpacing;
      if (request.kerning != null) context.fontKerning = previousKerning;
    }
  };
  return createTextLayoutService({
    ...(supportsDirectedContext ? {
      acquireProjectableNativeUnits: (requests, shape) => {
        // Each request remains one native word operation plus optional charged
        // ASCII suffix. Joined SPACE+Y probes are metrics, not new paint words.
        // Do not manufacture an isolated SPACE or split a contextual multiword
        // operation merely to make it look projectable.
        if (requests.some(request => !request.text.trimEnd()
          || request.text.replace(/ +$/u, '').includes(' '))) return undefined;
        const results = requests.map(request => shape({ ...request, measure: true, clusterGeometry: 'spaces' }));
        for (let index = 0; index < results.length; index++) {
          const request = requests[index]!, entry = results[index]!, result = entry.shape;
          if (result.spans.length !== 1) return undefined;
          if (request.baseDirection === 'rtl' && request.text.endsWith(' ')
            && !nativeVisiblePrefixMatches(request, result, entry.visiblePrefix)) return undefined;
          if (result.nativeBaseDirection !== request.baseDirection) throw new TypeError('Native unit receipt mismatch');
          if (!Number.isFinite(result.advancePt) || result.advancePt < 0
            || result.spans.map(span => span.text).join('') !== request.text) return undefined;
          let previous = 0;
          const boundaries = new Set([0, request.text.length, ...result.graphemeBoundaries]);
          for (const cluster of result.clusters ?? []) {
            if (request.text.slice(cluster.range.start, cluster.range.end) !== ' '
              || !boundaries.has(cluster.range.start) || !boundaries.has(cluster.range.end)
              || !Number.isFinite(cluster.offsetPt) || !Number.isFinite(cluster.advancePt)
              || cluster.offsetPt < previous || cluster.advancePt < 0
              || cluster.offsetPt + cluster.advancePt > result.advancePt) return undefined;
            previous = cluster.offsetPt + cluster.advancePt;
          }
          const spaces = request.text.split(' ').length - 1;
          if (spaces !== (result.clusters?.length ?? 0)) return undefined;
        }
        return results;
      },
    } : {}),
    ...(context ? { glyphContours: acquireCanvasGlyphContours, glyphContourFingerprint: 'canvas-half-coverage-contours-v2',
      ...(supportsDirectedContext ? { glyphContoursSupportDirection: true as const } : {}) } : {}),
    fonts: createFontResolver(inventory, {
      installedFamilies: options.installedFamilies,
      shadowedFamilies: [...shadowedFamilies],
      regionalFamilyLists: Object.fromEntries((['sc', 'tc', 'hk', 'jp', 'kr'] as const).map(region => [
        region,
        Object.fromEntries([...new Set([...routedFontFamilies, 'sans-serif', 'serif', 'monospace'])]
          .map(family => [family, normalizeFontFamilyUncached(
            family, source.fonts.familyClasses, source.fonts.familyPitches, region,
          )])),
      ])),
      scriptScopedFamilies: options.useGoogleFonts
        ? docxScriptScopedFamilies(options.installedSubstituteFamilies ?? [])
        : undefined,
      nativeRouteMetrics: options.nativeRouteMetrics,
      nativeFamilyLists: Object.fromEntries(routedFontFamilies.map((family) => [
        family,
        normalizeFontFamilyUncached(
          family,
          source.fonts.familyClasses,
          source.fonts.familyPitches,
        ),
      ])),
    }),
    cjkFallback: options.cjkFallback,
    fontMetrics,
    eastAsiaFontCharsets: fontFamilyCharsets,
    genericFamilies: Object.fromEntries(routedFontFamilies.map((family) => [
      family,
      classifyDocxFontGeneric(family, source.fonts.familyClasses, source.fonts.familyPitches),
    ])),
    measurer: {
      // Vertical OpenType capability is consulted only by vertical acquisition.
      // DOM-dependent vertical documents cannot retain worker mode, so folding
      // that unused capability into the general text snapshot would make equal
      // horizontal main/worker services advertise unequal cache identities.
      // v2 proves synthesized faces; v4 additionally owns temporary projectable
      // directed source bundles. Capability participates in the stable semantic
      // fingerprint, independently of source occurrence or cache history.
      fingerprint: supportsDirectedContext ? 'canvas-projectable-directed-prefix-text-metrics-v5'
        : context ? 'canvas-text-metrics-v2' : 'deterministic-text-metrics-v1',
      ...(context ? { sameGlyphRaster, ownsGlyphs } : {}),
      measure,
      ...(supportsDirectedContext ? { measureDirected: measure, supportsDirectedAuxiliaries: true as const } : {}),
    },
  });
}

export function createProductionLayoutServices(
  source: LayoutSourceStore,
  options: ProductionLayoutServiceOptions,
): LayoutServices {
  const text = createProductionTextLayoutService(source, options);
  const mathResources = options.mathResources ?? source.mathOccurrences.map(({ display, source: occurrenceSource }) => ({
    resourceKey: mathResourceKey(occurrenceSource, display ? 'display' : 'inline'),
    widthEm: 0,
    ascentEm: 0,
    descentEm: 0,
    available: false,
    diagnostics: [{
      code: 'UNSUPPORTED_FEATURE' as const,
      severity: 'warning' as const,
      message: 'The optional math renderer is unavailable; using the deterministic text fallback',
    }],
  }));
  const imageMetadata = source.imageMetadata;
  const services: LayoutServices = Object.freeze({
    text,
    images: createImageMetadataService(imageMetadata),
    math: createMathMetadataService(mathResources),
    ...(options.allowFootnoteContinuation === true
      ? { allowFootnoteContinuation: true as const }
      : {}),
    verticalGlyphFingerprint: options.verticalGlyphMeasurement.fingerprint,
  });
  const occurrenceKeys = source.mathOccurrences.map(({ source: occurrenceSource, display }) =>
    mathResourceKey(occurrenceSource, display ? 'display' : 'inline'));
  const metadataKeys = mathResources.map((resource) => resource.resourceKey);
  const missingMetadata = occurrenceKeys.filter((key) => !metadataKeys.includes(key));
  const extraMetadata = metadataKeys.filter((key) => !occurrenceKeys.includes(key));
  if (missingMetadata.length || extraMetadata.length) {
    throw new Error(
      `Math metadata membership mismatch: missing [${missingMetadata.join(', ')}]; extra [${extraMetadata.join(', ')}]`,
    );
  }
  attachPrivateResourceLookup(
    services,
    options.mathDrawables ?? new Map(),
    mathResources.filter((resource) => resource.available !== false)
      .map((resource) => resource.resourceKey),
  );
  attachPaintResourceRegistry(
    services,
    source.paintResources,
  );
  attachVerticalGlyphMeasurementService(services, options.verticalGlyphMeasurement);
  return services;
}
