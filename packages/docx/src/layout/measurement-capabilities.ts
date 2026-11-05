import type { WritingMode } from './types.js';

/** Canvas capabilities that retained layout may use synchronously.
 *
 * Vertical OpenType feature probing is deliberately excluded: it needs the
 * backing DOM canvas and is injected separately as a document-scoped service.
 */
export interface MeasurementTextContext {
  font: string;
  letterSpacing: string;
  fontKerning: CanvasFontKerning;
  /** Canvas `lang`, present only where the engine supports it. */
  lang?: string;
  measureText(text: string): TextMetrics;
}

/** Document-scoped vertical glyph measurement bound to the same concrete
 * context as {@link MeasurementTextContext}, without exposing its canvas or
 * paint methods to acquisition. Calls are synchronous because the underlying
 * OpenType feature probe mutates and restores font state in one stack frame. */
export interface VerticalGlyphMeasurementService {
  readonly fingerprint: string;
  measureRunInkExtra(text: string): number;
  planRun(input: Readonly<{
    text: string;
    font: string;
    fontKerning: CanvasFontKerning;
    fontSizePt: number;
    letterSpacingPt: number;
    /** Per-scalar grid/fit contribution, independent of run tracking. */
    characterGridDeltaPt?: number;
    charScale: number;
    growTrRotateInk: boolean;
    writingMode: WritingMode;
    /** Punctuation compression and width balance of a source range, floored
     * with the cell (WORD_CONDENSED_SPACING_ADVANCE_FLOOR); when supplied,
     * reported advances include it. */
    characterAdjustmentPt?: (start: number, end: number) => number;
  }>): readonly Readonly<{
    range: Readonly<{ start: number; end: number }>;
    text: string;
    orientation: 'upright' | 'rotate' | 'sideways';
    originPt: number;
    advancePt: number;
    drawOffsetPt: Readonly<{ xPt: number; yPt: number }>;
    verticalFeature: boolean;
    blockAxisInkBounds?: Readonly<{ startPt: number; endPt: number }>;
    floored?: true;
    naturalAdvancePt?: number;
  }>[];
}
