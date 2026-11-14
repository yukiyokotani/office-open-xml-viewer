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
 * - an unassigned scalar or a lone surrogate: malformed or unknown input
 *   proves nothing;
 * - a cluster whose canonical closure outgrows a sane cluster
 *   (`canonical-closure.ts`).
 *
 * A scalar that is not normalization-inert ({@link fontDemandInertScalar})
 * is demanded with its cluster's canonical closure: shaping may decompose or
 * compose it (`が` can paint `か` and U+3099; `e` U+0301 can paint `é`), so
 * every scalar of every canonically equivalent rendering is demanded. A
 * cluster that starts with a joiner also demands U+25CC, the dotted circle a
 * shaper may insert for a mark without a base.
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
import { canonicalClusterScalars, isCanonicallyDescribed, isClusterJoiner } from './canonical-closure.js';
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
    // A cluster: a starter and the joiners after it (canonical-closure.ts).
    let cluster: number[] = [];
    let canonical = false;
    const flush = (): boolean => {
      if (!cluster.length) return true;
      const scalars = canonical ? canonicalClusterScalars(cluster) : cluster;
      cluster = [];
      canonical = false;
      if (!scalars) return false;
      for (const cp of scalars) if (!this.admitScalar(cp)) return false;
      return true;
    };
    for (const scalar of text) {
      const cp = scalar.codePointAt(0) as number;
      if (cp >= 0xf020 && cp <= 0xf0ff) return this.decline();
      const inert = fontDemandInertScalar(cp);
      if (!inert && !isCanonicallyDescribed(cp)) return this.decline();
      if (inert || !isClusterJoiner(cp)) {
        if (!flush()) return this.decline();
      } else if (!cluster.length) {
        // A mark without a base: a shaper may paint a dotted circle for it.
        if (!this.admitScalar(0x25cc)) return this.decline();
      }
      cluster.push(cp);
      if (!inert) canonical = true;
    }
    if (!flush()) return this.decline();
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
    return fontDemandInertScalar(cp) && this.admitScalar(cp);
  }

  /** Demand one scalar within the budgets. */
  private admitScalar(cp: number): boolean {
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
