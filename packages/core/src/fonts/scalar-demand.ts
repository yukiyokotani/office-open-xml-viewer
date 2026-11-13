/**
 * Demanded Unicode scalars: the input to `unicode-range` subset laziness.
 *
 * A font source may split one family into many `unicode-range` subsets (a
 * Noto CJK stylesheet declares about a hundred). The painting realm loads only
 * the subsets whose range intersects the scalars the document will paint, so
 * it must know those scalars before layout. This module is the one collector
 * every format feeds with the text it paints; the format owns *which* text
 * (its traversal and its unmodelled paths), this module owns the scalar rules.
 *
 * Loading policy, not OOXML rendering behaviour. The answer is either an
 * explicit scalar list or `'all'` (load every subset). `'all'` is the
 * conservative answer and never an error; the collector returns it whenever
 * the painted scalars cannot be proven from the input text:
 *
 * - strong-RTL text: Canvas performs UAX #9 mirroring, which is unmodelled
 *   here, so a mirrored glyph could be painted that the text does not spell;
 * - PUA `U+F020–F0FF` outside a symbol-encoded family: an unmodelled
 *   symbol-font remapping could paint a different scalar than the stored one.
 *   Text in a symbol-encoded family ({@link FontDemandScalarCollector.addSymbolText})
 *   demands instead every scalar its conversion may paint, a finite set
 *   from the conversion tables (`symbol-font.ts`);
 * - any scalar that is not normalization-inert ({@link fontDemandInertScalar}):
 *   canonical composition or decomposition could paint a scalar the text does
 *   not spell, and malformed/unassigned input proves nothing;
 * - growth beyond the caller's JSON byte budget (the collected list is carried
 *   in a bounded projection) or beyond `HARD_MAX_FONT_DEMAND_SCALARS`.
 *
 * The production Canvas measurement sentinels (`ii` tracking, `M` baseline,
 * space for tabs) are always included because the renderers measure them with
 * every face they paint.
 */
import { hasStrongRtl } from '../text/bidi/line-order.js';
import {
  FONT_BASELINE_SENTINEL,
  FONT_SPACE_SENTINEL,
  FONT_TRACKING_SENTINEL,
} from '../internal/font-measurement-sentinels.js';
import { HARD_MAX_FONT_DEMAND_SCALARS } from '../worker/resource-policy.generated.js';
import { fontDemandInertScalar } from './inert-scalars.js';
import { symbolFontKey, symbolPaintCandidates, type SymbolFontKey } from './symbol-font.js';

/** `'all'`, or the demanded scalars in first-seen order. */
export type FontDemandScalars = 'all' | readonly number[];

export interface FontDemandScalarOptions {
  /** JSON byte allowance for the collected (uncommitted) scalar list, counted
   * as its `[n,n,…]` serialization; exceeding it yields `'all'`. */
  readonly budget: number;
  /** Scalars an earlier delta already demanded: neither collected again nor
   * charged, but counted against `HARD_MAX_FONT_DEMAND_SCALARS`. */
  readonly committed?: ReadonlySet<number>;
}

const NO_SCALARS: ReadonlySet<number> = new Set();

/** Incremental form for formats that interleave their own unmodelled-path
 * checks with text collection. Once a text is declined, the collector stays
 * `'all'`. */
export class FontDemandScalarCollector {
  private readonly points = new Set<number>();
  private readonly committed: ReadonlySet<number>;
  private readonly budget: number;
  // The empty delta still serializes as "[]"; charge it before any scalar.
  private bytes = 2;
  private declined: boolean;

  constructor(options: FontDemandScalarOptions) {
    this.budget = options.budget;
    this.committed = options.committed ?? NO_SCALARS;
    this.declined = this.bytes > this.budget;
    if (!this.declined) {
      this.add(FONT_BASELINE_SENTINEL + FONT_TRACKING_SENTINEL + FONT_SPACE_SENTINEL);
    }
  }

  /** True once some text could not be proven; the result is then `'all'`. */
  get isAll(): boolean { return this.declined; }

  /** Collect one painted text. Returns false (and becomes `'all'`) when the
   * text's painted scalars cannot be proven from its spelling. */
  add(text: string): boolean {
    if (this.declined) return false;
    if (hasStrongRtl(text)) return this.decline();
    for (const scalar of text) {
      const cp = scalar.codePointAt(0) as number;
      if ((cp >= 0xf020 && cp <= 0xf0ff) || !this.admit(cp)) return this.decline();
    }
    return true;
  }

  /**
   * Collect one text painted in a font list that may name symbol-encoded
   * families (`fontFamilies`, any order; `symbol-font.ts`). Each character
   * the painter may convert demands every scalar its conversion may write in
   * any of those families ({@link symbolPaintCandidates}): which ones paint
   * depends on what the painting font draws, judged only at paint time. The
   * rest of the text is collected as by {@link add}. Without a symbol family
   * this is {@link add}.
   */
  addSymbolText(text: string, fontFamilies: Iterable<string | null | undefined>): boolean {
    if (this.declined) return false;
    const keys = new Set<SymbolFontKey>();
    for (const family of fontFamilies) {
      const key = symbolFontKey(family);
      if (key) keys.add(key);
    }
    if (!keys.size) return this.add(text);
    let rest = '';
    for (const ch of text) {
      const cp = ch.codePointAt(0) as number;
      if (!(cp >= 0xf020 && cp <= 0xf0ff) && (cp < 0x20 || cp > 0xff)) {
        rest += ch;
        continue;
      }
      for (const key of keys) {
        for (const scalar of symbolPaintCandidates(ch, key)) {
          if (!this.admit(scalar.codePointAt(0) as number)) return this.decline();
        }
      }
    }
    return this.add(rest);
  }

  /** Demand one normalization-inert scalar within the budgets. */
  private admit(cp: number): boolean {
    if (!fontDemandInertScalar(cp)) return false;
    if (this.committed.has(cp) || this.points.has(cp)) return true;
    this.bytes += String(cp).length + 1;
    if (this.bytes > this.budget
      || this.committed.size + this.points.size >= HARD_MAX_FONT_DEMAND_SCALARS) return false;
    this.points.add(cp);
    return true;
  }

  /** Mark an unmodelled paint path: the result becomes `'all'`. */
  decline(): false {
    this.declined = true;
    return false;
  }

  result(): FontDemandScalars {
    return this.declined ? 'all' : [...this.points];
  }
}

/** One-shot collection over every painted text of a document or delta. */
export function collectFontDemandScalars(
  texts: Iterable<string>,
  options: FontDemandScalarOptions,
): FontDemandScalars {
  const collector = new FontDemandScalarCollector(options);
  for (const text of texts) if (!collector.add(text)) break;
  return collector.result();
}
