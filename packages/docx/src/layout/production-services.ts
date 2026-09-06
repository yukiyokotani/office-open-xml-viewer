import {
  canvasFontString,
  measureResolvedCanvasFontBoxRatio,
  normalizeFontMetricFamily,
} from '@silurus/ooxml-core';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { DOCX_GOOGLE_FONTS } from '../google-fonts.js';
import { normalizeFontFamilyUncached } from '../line-layout.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import { createFontResolver, type FontInventoryFace } from './font-service.js';
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
  snapshotFontMetrics,
  type GlyphMeasureRequest,
} from './text.js';
import type { LayoutServices } from './types.js';
import { wordResolvedEastAsianSingleLineRatio } from './line-compatibility.js';
import type { DocxResolvedFontMetricCandidate } from '../document-content.js';

export interface LoadedFontFaceRecord {
  readonly family: string;
  readonly status: string;
  readonly style: string;
  readonly weight: string;
}

export interface ProductionLayoutServiceOptions {
  readonly localMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly useGoogleFonts?: boolean;
  readonly mathResources?: readonly MathLayoutResource[];
  readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
  readonly measureContext: MeasurementTextContext | null;
  readonly verticalGlyphMeasurement: VerticalGlyphMeasurementService;
  readonly embeddedFaces?: readonly LoadedFontFaceRecord[];
  readonly googleFaces?: readonly LoadedFontFaceRecord[];
  /** Derive Word line allocation only from the concrete authored Canvas face
   * that proves coverage of its fontTable East-Asian charset. */
  readonly measureResolvedFontMetrics?: boolean;
  readonly resolvedFontMetricCandidates?: readonly DocxResolvedFontMetricCandidate[];
}

function canvasResolvedFontMetrics(
  candidates: readonly DocxResolvedFontMetricCandidate[],
  context: MeasurementTextContext | null,
): Readonly<Record<string, ResolvedFontMetric>> {
  if (!context) return {};
  const metrics: Record<string, ResolvedFontMetric> = {};
  for (const candidate of candidates) {
    const family = candidate.family.trim();
    if (!family) continue;
    const key = normalizeFontMetricFamily(family);
    if (metrics[key]) continue;
    // The document-content projection proves this family actually wins a
    // rendered script slot. Equal selected/control glyph ink means Canvas
    // silently substituted a fallback, so no resource metric is claimed.
    const fontBoxRatio = measureResolvedCanvasFontBoxRatio(
      context,
      family,
      { text: candidate.probeText, emPx: 100 },
    );
    if (!(fontBoxRatio != null && fontBoxRatio > 0)) continue;
    const eastAsianLineHeightRatio = wordResolvedEastAsianSingleLineRatio(fontBoxRatio);
    if (!(eastAsianLineHeightRatio > 0)) continue;
    metrics[key] = Object.freeze({
      family,
      requestedFamily: family,
      weight: 400,
      style: 'normal',
      sourceIdentity: `canvas-resolved:${family}`,
      synthesized: false,
      fontBoxRatio,
      ...(candidate.appliesToLatin ? { lineHeightRatio: eastAsianLineHeightRatio } : {}),
      eastAsianLineHeightRatio,
    });
  }
  return Object.freeze(metrics);
}

export function createProductionLayoutServices(
  source: LayoutSourceStore,
  options: ProductionLayoutServiceOptions,
): LayoutServices {
  const measuredFontMetrics = options.measureResolvedFontMetrics
    ? canvasResolvedFontMetrics(
        options.resolvedFontMetricCandidates ?? [],
        options.measureContext,
      )
    : {};
  const localMetrics = snapshotFontMetrics(options.localMetrics);
  const fontMetrics = snapshotFontMetrics({
    ...measuredFontMetrics,
    ...localMetrics,
    ...options.fontMetrics,
  });
  const fontFamilyCharsets = Object.freeze(Object.fromEntries(
    Object.entries(source.fontFamilyCharsets)
      .map(([family, charset]) => [family.trim().toLowerCase(), charset]),
  ));
  const displayFaceFamily = (family: string): string => family
    .trim()
    .replace(/^(['"])(.*)\1$/, '$2');
  const normalizedFaceFamily = (family: string): string => displayFaceFamily(family)
    .toLocaleLowerCase('en-US');
  const loadedFaceStyle = (face: LoadedFontFaceRecord): 'normal' | 'italic' | null => {
    const style = face.style.trim().toLocaleLowerCase('en-US');
    return style === 'normal' || style === 'italic' ? style : null;
  };
  const loadedFaceWeight = (face: LoadedFontFaceRecord): number | null => {
    const weight = face.weight.trim().toLocaleLowerCase('en-US');
    if (weight === 'normal') return 400;
    if (weight === 'bold') return 700;
    if (!/^\d+$/.test(weight)) return null;
    const numeric = Number(weight);
    return numeric >= 100 && numeric <= 900 ? numeric : null;
  };
  const loadedFaces = (faces: readonly LoadedFontFaceRecord[]) => faces.flatMap((face) => {
    if (face.status !== 'loaded') return [];
    const weight = loadedFaceWeight(face);
    const style = loadedFaceStyle(face);
    return weight == null || style == null ? [] : [{
      family: normalizedFaceFamily(face.family),
      displayFamily: displayFaceFamily(face.family),
      weight,
      style,
    }];
  });
  const successfulEmbedded = new Map(loadedFaces(options.embeddedFaces ?? []).map((loaded) => [
    `${loaded.family}:${loaded.weight}:${loaded.style}`, loaded,
  ]));
  const inventory: FontInventoryFace[] = source.fonts.embeddedFonts.flatMap((font) => {
    const weight = font.style === 'bold' || font.style === 'boldItalic' ? 700 : 400;
    const style = font.style === 'italic' || font.style === 'boldItalic' ? 'italic' as const : 'normal' as const;
    const loaded = successfulEmbedded.get(`${normalizedFaceFamily(font.fontName)}:${weight}:${style}`);
    return loaded ? [{
      requestedFamily: font.fontName,
      resolvedFamily: loaded.displayFamily,
      source: 'embedded' as const,
      weight,
      style,
    }] : [];
  });
  for (const [requestedFamily, metric] of Object.entries(localMetrics)) {
    inventory.push({
      requestedFamily: metric.requestedFamily ?? requestedFamily,
      resolvedFamily: metric.family,
      source: 'local',
      weight: metric.weight ?? 400,
      style: metric.style ?? 'normal',
    });
  }
  if (options.useGoogleFonts) {
    const successfulGoogle = loadedFaces(options.googleFaces ?? []);
    const seen = new Set<string>();
    for (const name of source.fonts.preloadNames) {
      if (!name) continue;
      const key = name.toLocaleLowerCase('en-US');
      if (seen.has(key)) continue;
      seen.add(key);
      const entry = DOCX_GOOGLE_FONTS[key];
      const resolvedFamily = entry?.loadFamily ?? name;
      if (!entry) continue;
      for (const loaded of successfulGoogle.filter(
        (face) => face.family === normalizedFaceFamily(resolvedFamily),
      )) {
        inventory.push({
          requestedFamily: name,
          resolvedFamily: loaded.displayFamily,
          source: normalizedFaceFamily(resolvedFamily) === normalizedFaceFamily(name)
            ? 'google' : 'substitute',
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
  const text = createTextLayoutService({
    fonts: createFontResolver(inventory, {
      nativeFamilyLists: Object.fromEntries(routedFontFamilies.map((family) => [
        family,
        normalizeFontFamilyUncached(
          family,
          source.fonts.familyClasses,
          source.fonts.familyPitches,
        ),
      ])),
    }),
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
      fingerprint: context ? 'canvas-text-metrics-v1' : 'deterministic-text-metrics-v1',
      measure(request: Readonly<GlyphMeasureRequest>) {
        if (!context) return {
          advancePt: [...request.text].length * request.fontSizePt * 0.5,
          ascentPt: request.fontSizePt * 0.8,
          descentPt: request.fontSizePt * 0.2,
        };
        const previousFont = context.font;
        const previousLetterSpacing = context.letterSpacing;
        const previousKerning = context.fontKerning;
        try {
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
            ascentPt: metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent ?? 0,
            descentPt: metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent ?? 0,
            ...(Object.values(inkBounds).every(Number.isFinite) ? {
              inkBounds,
              ...(horizontalInkBoundsAreTight ? { horizontalInkBoundsAreTight: true } : {}),
            } : {}),
          };
        } finally {
          context.font = previousFont;
          context.letterSpacing = previousLetterSpacing;
          if (request.kerning != null) context.fontKerning = previousKerning;
        }
      },
    },
  });
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
