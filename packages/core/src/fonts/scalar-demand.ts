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
 * - a scalar that is not normalization-inert and that the host does not know
 *   as assigned, or a lone surrogate: malformed or unknown input proves
 *   nothing;
 * - a cluster whose canonical closure outgrows a sane cluster
 *   (`canonical-closure.ts`), including the closure of the joiners that
 *   start texts with a scalar they may compose with;
 * - growth beyond the caller's JSON byte budget (the collected list is carried
 *   in a bounded projection) or beyond `HARD_MAX_FONT_DEMAND_SCALARS`.
 *
 * A scalar that is not normalization-inert ({@link fontDemandInertScalar})
 * is demanded with its cluster's canonical closure: shaping may decompose or
 * compose it (`が` can paint `か` and U+3099; `e` U+0301 can paint `é`), so
 * every scalar of every canonically equivalent rendering is demanded. A text
 * that starts with joiners may be shaped after any other painted text, so
 * those joiners are also closed with every demanded scalar they may compose
 * with.
 * Any combining mark or joiner also demands U+25CC, the dotted circle a
 * shaper may insert for a mark without a valid base, and the few scalars
 * shapers split outside canonical equivalence demand their parts
 * (`shaperSplitScalars`).
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
import { canonicalClusterScalars, isCanonicallyKnown, isClusterJoiner, mayCompose, shaperSplitScalars } from './canonical-closure.js';
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

/** Any combining mark (General_Category M). */
const MARK_TEXT = /\p{M}/u;


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
  /** The cluster being collected (a starter and the joiners after it,
   * `canonical-closure.ts`). */
  private cluster: number[] = [];
  /** Whether {@link cluster} holds a scalar that is not normalization-inert. */
  private canonical = false;
  /** Every joiner of a run of joiners some text starts with. Layout may
   * shape such a text after any other painted text (adjacent runs of one
   * format, runs a deletion or an unchosen optional hyphen separated,
   * case-mapped runs), so {@link result} closes these joiners with every
   * demanded scalar they may compose with. */
  private readonly leading = new Set<number>();
  /** Whether U+25CC is demanded: a shaper may paint a dotted circle for a
   * mark or joiner it finds without a valid base (HarfBuzz broken clusters,
   * Hangul tone marks), wherever it stands in the text. */
  private dottedCircle = false;

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
    if (!this.dottedCircle && MARK_TEXT.test(text) && !this.demandDottedCircle()) return this.decline();
    // Whether this text has not yet reached its first starter.
    let leading = true;
    for (const scalar of text) {
      const cp = scalar.codePointAt(0) as number;
      if (cp >= 0xf020 && cp <= 0xf0ff) return this.decline();
      // A soft hyphen is demanded but leaves the cluster open: a painter may
      // strip it and draw the scalars around it as one text (DOCX §17.3.3.29).
      if (cp === 0xad) {
        if (!this.admitScalar(cp)) return this.decline();
        continue;
      }
      const inert = fontDemandInertScalar(cp);
      if (!inert && !isCanonicallyKnown(cp)) return this.decline();
      if (cp >= 0xe31 && cp <= 0x2011) {
        for (const split of shaperSplitScalars(cp)) if (!this.admitScalar(split)) return this.decline();
      }
      if (inert || !isClusterJoiner(cp)) {
        if (!this.flush()) return this.decline();
        leading = false;
      } else {
        if (!this.dottedCircle && !this.demandDottedCircle()) return this.decline();
        if (leading) this.leading.add(cp);
      }
      this.cluster.push(cp);
      if (!inert) this.canonical = true;
    }
    return this.flush() || this.decline();
  }

  private demandDottedCircle(): boolean {
    this.dottedCircle = true;
    return this.admitScalar(0x25cc);
  }

  /** Close the leading joiners ({@link leading}) with each demanded scalar
   * they may compose with, to a fixed point: a scalar composed with one run
   * may compose with another run that follows it. The closure of all of
   * them together is a superset of each run's. */
  private closeLeadingJoiners(): boolean {
    const joiners = [...this.leading];
    const tried = new Set<number>();
    for (let pending = joiners.length ? [...this.committed, ...this.points] : []; pending.length;) {
      const next: number[] = [];
      for (const starter of pending) {
        if (tried.has(starter)) continue;
        tried.add(starter);
        if (!joiners.some((joiner) => mayCompose(starter, joiner))) continue;
        const scalars = canonicalClusterScalars([starter, ...joiners]);
        if (!scalars) return false;
        for (const cp of scalars) {
          if (this.committed.has(cp) || this.points.has(cp)) continue;
          if (!this.admitScalar(cp)) return false;
          next.push(cp);
        }
      }
      pending = next;
    }
    this.leading.clear();
    return true;
  }

  /** Demand the pending cluster: its scalars, or its canonical closure when
   * it holds a scalar that is not normalization-inert. */
  private flush(): boolean {
    if (!this.cluster.length) return true;
    const scalars = this.canonical ? canonicalClusterScalars(this.cluster) : this.cluster;
    this.cluster = [];
    this.canonical = false;
    if (!scalars) return false;
    for (const cp of scalars) if (!this.admitScalar(cp)) return false;
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
    if (!this.declined && !this.closeLeadingJoiners()) this.decline();
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
