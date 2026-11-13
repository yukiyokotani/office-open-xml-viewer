// Variable-font line metrics for font-source faces: reading a variable face's
// fvar/avar/MVAR (open-type-variations.ts) and instancing it. Only the
// font-source session reads served faces' tables and only DOCX layout of
// served faces instances them, so this module stays out of the static graph
// that every document loads (open-type-metrics.ts keeps the static parsers).
import {
  lineMetricsVaryAlong,
  metricDelta,
  parseOpenTypeVariation,
  variationCoordinates,
  type LineMetricMvarTag,
} from './open-type-variations.js';
import { parseOpenTypeResourceMetrics, type OpenTypeLineMetrics } from './open-type-metrics.js';

/** {@link parseOpenTypeResourceMetrics} that also reads a variable face
 * (its default instance plus the variation data {@link instanceOpenTypeLineMetrics}
 * needs). */
export function parseVariableOpenTypeResourceMetrics(
  bytes: Uint8Array,
  faceIndex?: number,
): OpenTypeLineMetrics | null {
  return parseOpenTypeResourceMetrics(bytes, faceIndex, { variations: parseOpenTypeVariation });
}

/**
 * Line metrics of one instance of a variable font: the default-instance
 * values plus the MVAR deltas at `axisValues` (user-space values by axis tag;
 * a missing axis takes its fvar default). A static font is returned as is.
 * MVAR tags follow the OpenType value-tag table: `hasc`/`hdsc`/`hlgp` vary
 * the OS/2 typo values, `hcla`/`hcld` usWin; the hhea values have no tag and
 * never vary. Returns null when a value of `opszTags` (default: every one of
 * them) varies along `opsz`: the browser sets that axis from the font size
 * (CSS `font-optical-sizing: auto`), so one size-free value does not exist.
 * A caller passes only the tags its projection reads.
 */
export function instanceOpenTypeLineMetrics(
  metrics: OpenTypeLineMetrics,
  axisValues: Readonly<Record<string, number>>,
  opszTags?: readonly LineMetricMvarTag[],
): OpenTypeLineMetrics | null {
  const variation = metrics.variation;
  if (!variation) return metrics;
  if (lineMetricsVaryAlong(variation, 'opsz', opszTags)) return null;
  const coords = variationCoordinates(variation, axisValues);
  const delta = (tag: Parameters<typeof metricDelta>[1]) => metricDelta(variation, tag, coords);
  const { variation: _variation, ...base } = metrics;
  void _variation;
  return Object.freeze({
    ...base,
    ...(metrics.winAscent === undefined ? {} : {
      winAscent: metrics.winAscent + delta('hcla'),
      winDescent: (metrics.winDescent ?? 0) + delta('hcld'),
    }),
    ...(metrics.typoAscent === undefined ? {} : {
      typoAscent: metrics.typoAscent + delta('hasc'),
      typoDescent: (metrics.typoDescent ?? 0) + delta('hdsc'),
      typoLineGap: (metrics.typoLineGap ?? 0) + delta('hlgp'),
    }),
  });
}
