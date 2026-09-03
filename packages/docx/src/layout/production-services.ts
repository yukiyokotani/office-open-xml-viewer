import { canvasFontString, providerFontFamily, providerFontSource } from '@silurus/ooxml-core';
import type { FontFamilyRoutes, ResolvedLocalFontMetric } from '@silurus/ooxml-core';
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
import { createPaintResourceRegistry } from './paint-resources.js';
import {
  attachPaintResourceRegistry,
  attachPrivateResourceLookup,
  attachVerticalGlyphMeasurementService,
} from './runtime-state.js';
import {
  classifyDocxFontGeneric,
  createTextLayoutService,
  snapshotLocalMetrics,
  type GlyphMeasureRequest,
} from './text.js';
import type { LayoutServices } from './types.js';

export interface LoadedFontFaceRecord {
  readonly family: string;
  readonly status: string;
  readonly style: string;
  readonly weight: string;
}

export interface ProductionLayoutServiceOptions {
  readonly localMetrics?: Readonly<Record<string, ResolvedLocalFontMetric>>;
  readonly mathResources?: readonly MathLayoutResource[];
  readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
  readonly measureContext: MeasurementTextContext | null;
  readonly verticalGlyphMeasurement: VerticalGlyphMeasurementService;
  readonly embeddedFaces?: readonly LoadedFontFaceRecord[];
  readonly providerRoutes?: FontFamilyRoutes;
}

export function createProductionLayoutServices(
  source: LayoutSourceStore,
  options: ProductionLayoutServiceOptions,
): LayoutServices {
  const localMetrics = snapshotLocalMetrics(options.localMetrics);
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
  for (const requestedFamily of Object.keys(options.providerRoutes ?? {})) {
    const resolvedFamily = providerFontFamily(options.providerRoutes, requestedFamily);
    const source = providerFontSource(options.providerRoutes, requestedFamily);
    if (resolvedFamily && source) inventory.push({ requestedFamily, resolvedFamily, source });
  }
  for (const [requestedFamily, metric] of Object.entries(localMetrics)) {
    inventory.push({
      requestedFamily: metric.requestedFamily ?? requestedFamily,
      resolvedFamily: metric.family,
      source: 'local',
      weight: metric.weight ?? 400,
      style: metric.style ?? 'normal',
    });
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
    localMetrics,
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
    options.providerRoutes && Object.keys(options.providerRoutes).length > 0
      ? createPaintResourceRegistry(source.paintResources.descriptors.map((descriptor) => (
          descriptor.kind === 'chart'
            ? {
                ...descriptor,
                model: { ...descriptor.model, providerFontRoutes: options.providerRoutes },
              }
            : descriptor
        )))
      : source.paintResources,
  );
  attachVerticalGlyphMeasurementService(services, options.verticalGlyphMeasurement);
  return services;
}
