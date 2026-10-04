// PPTX-only font-slot compatibility. ECMA-376 Part 1 §21.1.2.3 supplies
// the normative Unicode table (the otherwise slot is ea); §21.1.2.3.9 supplies
// lang. [MS-OI29500] §2.1.1397 discusses font substitution, not these measured
// slot deviations. Do not infer a slot from font fallback or a line-break class.
// DOCX uses WordprocessingML §17.3.2.26; Excel's DrawingML adapter has its own
// Office behaviour. The measured PowerPoint overrides must not affect either.
//
// Observed Office deviations take precedence for the tested pairs. They apply
// only to the code points and language identifiers tested, not a whole block or
// every regional variant. Swept characters without a distinguishable PDF font
// use the normative table. Unmeasured scripts retain the pre-1653 routing until
// distinct-face controls settle Office's behaviour; no new script override is
// inferred from a fallback font. This is intentionally incomplete Office coverage.
// The extra Windows cycles settle European digits and exact Myanmar-extension
// scalars. Replacement 10/18/32 pt boundary controls validate Japanese U+201F
// across three face triples, but reject general symbol deviations at the tested
// endpoints. Those overrides are withdrawn below; other unmeasured pairs retain
// prior routing. See POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE for the full matrix.
import { graphemeClusterOffsets, isCjkBreakChar, isComplexScriptCodePoint } from '@silurus/ooxml-core';
export { POWERPOINT_FONT_SLOT_EVIDENCE } from './font-slot-evidence.js';

export type PowerPointFontSlot = 'latin' | 'ea' | 'cs' | 'sym';
type SlotRange = readonly [start: number, end: number, slot: PowerPointFontSlot];

/** Sorted, disjoint ECMA-376 §21.1.2.3 exceptions to the otherwise ea slot. */
const NORMATIVE_SLOT_RANGES: readonly SlotRange[] = [
  [0x0000, 0x00A6, 'latin'],
  [0x00A9, 0x00AF, 'latin'],
  [0x00B2, 0x00B3, 'latin'],
  [0x00B5, 0x00D6, 'latin'],
  [0x00D8, 0x00F6, 'latin'],
  [0x00F8, 0x058F, 'latin'],
  [0x0590, 0x074F, 'cs'],
  [0x0780, 0x07BF, 'cs'],
  [0x0900, 0x109F, 'cs'],
  [0x10A0, 0x10FF, 'latin'],
  [0x1200, 0x137F, 'latin'],
  [0x13A0, 0x177F, 'latin'],
  [0x1780, 0x18AF, 'cs'],
  [0x1D00, 0x1D7F, 'latin'],
  [0x1E00, 0x1FFF, 'latin'],
  [0x2000, 0x200B, 'latin'],
  [0x200C, 0x200F, 'cs'],
  [0x2010, 0x2029, 'latin'],
  [0x202A, 0x202F, 'cs'],
  [0x2030, 0x2046, 'latin'],
  [0x204A, 0x245F, 'latin'],
  [0x2670, 0x2671, 'cs'],
  [0x27C0, 0x2BFF, 'latin'],
  [0xF000, 0xF0FF, 'sym'],
  [0xFB00, 0xFB17, 'latin'],
  [0xFB1D, 0xFB4F, 'cs'],
  [0xFE50, 0xFE6F, 'latin'],
  [0x1D400, 0x1D7FF, 'latin'],
];

const EN_US: readonly SlotRange[] = [
  // Observed interior routing is retained. Replacement cycles under en-US/ja-JP
  // at 10/18/32 pt reject the former endpoints: U+24FF follows ea in one complete
  // triple (two others substitute); U+259F follows a fixed face in all three
  // triples. Neither supports a general latin override. ECMA's otherwise ea
  // applies to those endpoints; U+2500 remains latin in its complete triple.
  [0x2500, 0x259E, 'latin'],
  // U+2619 and U+2670/2671 each have 18 inconsistent cycles (three triples ×
  // three sizes × en-US/ja-JP), including nearby U+261A/266F/2672 counterexamples.
  // Withdraw their single-triple latin overrides: normative otherwise ea for
  // U+2619, explicit cs for U+2670/2671 (§21.1.2.3). Office slot independence
  // remains unresolved; no font-name-dependent compatibility rule is inferred.
  [0x2680, 0x2691, 'latin'],
  [0x2698, 0x2698, 'latin'],
  [0x269A, 0x269A, 'latin'],
  [0x269D, 0x269F, 'latin'],
  [0x26A2, 0x26A6, 'latin'],
  [0x26A8, 0x26A9, 'latin'],
  [0x26AC, 0x26AF, 'latin'],
  [0x26B2, 0x26BC, 'latin'],
  [0x26BF, 0x26C3, 'latin'],
  [0x26C6, 0x26C7, 'latin'],
  [0x26C9, 0x26CD, 'latin'],
  [0x26D0, 0x26D0, 'latin'],
  [0x26D2, 0x26D2, 'latin'],
  [0x26D5, 0x26E8, 'latin'],
  [0x26EB, 0x26EF, 'latin'],
  [0x26F6, 0x26F6, 'latin'],
  [0x26FB, 0x26FC, 'latin'],
  [0x26FE, 0x2700, 'latin'],
  [0x275F, 0x2760, 'latin'],
  [0x2768, 0x2775, 'latin'],
];

// Observed Japanese quote endpoint: all three cyclic assignments across three
// face triples and 10/18/32 pt select latin for U+201F, while U+201D/E select ea.
// en-US selects latin for all three. This supports the exact Japanese override,
// not other language IDs, surrounding contexts or the remaining symbol ranges.
const JA_JP: readonly SlotRange[] = [[0x201F, 0x201F, 'latin'], ...EN_US];

const KO_KR: readonly SlotRange[] = [
  [0x201F, 0x201F, 'latin'],
];

const ZH_CN = KO_KR;

const ZH_TW = KO_KR;

const HE_IL: readonly SlotRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x00BB, 0x00BB, 'cs'],
  [0x00D7, 0x00D7, 'latin'],
  [0x00F7, 0x00F7, 'latin'],
  [0x2047, 0x2048, 'latin'],
];

const AR_SA = HE_IL;

const TH_TH: readonly SlotRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x00D7, 0x00D7, 'latin'],
  [0x00F7, 0x00F7, 'latin'],
  [0x2047, 0x2048, 'latin'],
];

const HI_IN = TH_TH;

// Extra tagged-PDF cyclic controls at 20 pt settle 0–9 both alone and between
// Latin letters for these exact language IDs. en-US with each tested altLang
// retains latin digits; altLang does not replace the classification language.
const CS_DIGITS: readonly SlotRange[] = [[0x0030, 0x0039, 'cs']];

// Complete 20 pt cyclic controls select cs for isolated » × ÷, U+2018–201E
// and U+2047/U+2048, and latin between Latin letters (tested with A/B), for
// these exact lang IDs.
// This is observed PowerPoint itemization, not ECMA's scalar table. Native
// neighbours, punctuation sequences and other lang IDs retain prior routing.
const CONTEXTUAL_CS_LANGUAGES = new Set([
  'ar-eg', 'ar-sa', 'fa-ir', 'he', 'he-il', 'hi-in',
  'syr-sy', 'th-th', 'ug-cn', 'ur-in', 'ur-pk', 'yi-001',
]);
const LATIN_LETTER_RE = /^(?=\p{Script=Latin})\p{Letter}$/u;

// ECMA-376 §21.1.2.3's otherwise-ea rule, confirmed by all three cyclic faces
// for these exact Myanmar extension scalars under en-US/my-MM/ja-JP at 20 pt.
// Standalone U+A9E5/U+AA7B–AA7D also have 12 complete exact-scalar cycles;
// marks do not imply substitution. Unassigned U+A9FF has no scalar observation.
// Other language IDs retain the previous script policy rather than extrapolate.
const MYANMAR_EXTENSION_EA: readonly SlotRange[] = [
  [0xA9E0, 0xA9FE, 'ea'], [0xAA60, 0xAA7F, 'ea'],
];
const MYANMAR_MEASURED_LANGUAGES = new Set(['en-us', 'my-mm', 'ja-jp']);

const OBSERVED_OVERRIDES = new Map<string, readonly SlotRange[]>([
  ['en-us', EN_US], ['ja-jp', JA_JP], ['ko-kr', KO_KR],
  ['zh-cn', ZH_CN], ['zh-tw', ZH_TW], ['he-il', HE_IL],
  // Additional 20 pt cycles settle U+201F for these exact quote-language IDs.
  ['zh-hk', KO_KR], ['zh-mo', KO_KR], ['zh-sg', KO_KR], ['ii-cn', KO_KR],
  ['ar-sa', AR_SA], ['th-th', TH_TH], ['hi-in', HI_IN],
  ['fa-ir', CS_DIGITS], ['ur-pk', CS_DIGITS],
  ['yi-001', CS_DIGITS], ['syr-sy', CS_DIGITS], ['ug-cn', CS_DIGITS],
  ['ar-eg', CS_DIGITS], ['he', CS_DIGITS], ['ur-in', CS_DIGITS],
]);

// The full sweep's block boundaries, including excluded/unextractable scalars.
// A glyph-fallback observation does not select a slot: specification defaults
// apply within these blocks. Outside them the existing script policy is kept.
const SWEPT_BLOCKS: readonly (readonly [number, number])[] = [
  [0x0021, 0x007E], [0x00A1, 0x017F], [0x02B0, 0x02FF], [0x0370, 0x04FF],
  [0x2000, 0x209F], [0x20A0, 0x20CF], [0x2100, 0x23FF], [0x2460, 0x27BF],
  [0x3000, 0x303F], [0x3200, 0x33FF], [0xFE30, 0xFE4F], [0xFF01, 0xFF9F], [0xFFE0, 0xFFEE],
];
const EXISTING_INDIC_CS_RE = /[\p{Script=Devanagari}\p{Script=Thai}\p{Script=Bengali}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Oriya}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Lao}\p{Script=Myanmar}\p{Script=Tibetan}]/u;

const NORMATIVE_EA_QUOTE_LANGUAGES = new Set([
  'ii-cn', 'ja-jp', 'ko-kr', 'zh-cn', 'zh-hk', 'zh-mo', 'zh-sg', 'zh-tw',
]);

function slotInRanges(cp: number, ranges: readonly SlotRange[]): PowerPointFontSlot | undefined {
  let lo = 0;
  let hi = ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const [start, end, slot] = ranges[mid];
    if (cp < start) hi = mid - 1;
    else if (cp > end) lo = mid + 1;
    else return slot;
  }
  return undefined;
}

/** Paragraph context uses UTF-16 offsets before any display mapping. */
export interface PowerPointFontContext { text: string; offset: number }

/** Without paragraph context, preserve the historical scalar routing. */
export function powerPointFontSlot(cp: number, lang?: string, context?: PowerPointFontContext): PowerPointFontSlot {
  const language = lang?.toLowerCase() ?? '';
  if (context && CONTEXTUAL_CS_LANGUAGES.has(language)
    && (cp === 0xbb || cp === 0xd7 || cp === 0xf7 || (cp >= 0x2018 && cp <= 0x201e) || cp === 0x2047 || cp === 0x2048)) {
    if (context.text === String.fromCodePoint(cp)) return 'cs';
    // A neighbouring scalar can occupy two UTF-16 code units. Reading just a
    // code unit here would make a supplementary Latin letter a run-seam bug.
    const beforeIndex = context.offset > 1
      && context.text.charCodeAt(context.offset - 1) >= 0xdc00
      && context.text.charCodeAt(context.offset - 1) <= 0xdfff ? context.offset - 2 : context.offset - 1;
    const before = context.text.codePointAt(beforeIndex);
    const after = context.text.codePointAt(context.offset + String.fromCodePoint(cp).length);
    if (before !== undefined && after !== undefined
      && LATIN_LETTER_RE.test(String.fromCodePoint(before))
      && LATIN_LETTER_RE.test(String.fromCodePoint(after))) return 'latin';
  }
  if (cp >= 0xf000 && cp <= 0xf0ff) return 'sym';
  const overrides = OBSERVED_OVERRIDES.get(language);
  const observed = overrides && slotInRanges(cp, overrides);
  if (observed !== undefined) return observed;
  if (MYANMAR_MEASURED_LANGUAGES.has(language)) {
    const extensionSlot = slotInRanges(cp, MYANMAR_EXTENSION_EA);
    if (extensionSlot !== undefined) return extensionSlot;
  }
  if (cp >= 0x2018 && cp <= 0x201f && NORMATIVE_EA_QUOTE_LANGUAGES.has(language)) return 'ea';
  if (SWEPT_BLOCKS.some(([start, end]) => cp >= start && cp <= end)) {
    return slotInRanges(cp, NORMATIVE_SLOT_RANGES) ?? 'ea';
  }
  // The role-permutation controls distinguish Jamo's ea slot (unlike the
  // unresolved scripts). Do not confuse conjoining Jamo with break units.
  if (cp >= 0x1100 && cp <= 0x11ff) return 'ea';
  if (isComplexScriptCodePoint(cp) || EXISTING_INDIC_CS_RE.test(String.fromCodePoint(cp))) return 'cs';
  return isCjkBreakChar(cp) ? 'ea' : 'latin';
}

/** Measured ja-JP display mapping; still uses U+005C's latin slot. */
export function powerPointDisplayCluster(cluster: string, lang?: string): string {
  return lang?.toLowerCase() === 'ja-jp' && cluster.startsWith('\\')
    ? `¥${cluster.slice(1)}` : cluster;
}

interface PowerPointFontUnit {
  start: number;
  end: number;
  slot: PowerPointFontSlot;
}

/** Resolve slots once on paragraph text, before layout, paint or font loading.
 * Extenders normally inherit the base run's slot/format. The measured exception
 * is precisely U+1000 plus one of U+A9E5/U+AA7B–AA7D under en-US/my-MM/ja-JP:
 * all 24 single-run/seam cycles select cs for the base and ea for the mark.
 * Record those two-scalar clusters as separate slot units. This is attribution
 * metadata, not permission to split a Canvas shaping call: Canvas cannot attach
 * a mark across calls with different fonts. `graphemeEnds` preserves the original
 * paint boundaries independently of these units. Standalone marks use their
 * independently measured ea slot. Other bases, longer clusters and
 * language-changing seams retain previous inheritance:
 * they have no complete split-font evidence.
 * Font units do not redefine Unicode graphemes or core's line-break policy.
 */
export function powerPointFontRouting(runs: readonly { text: string | null; lang?: string }[]): {
  text: string;
  starts: number[];
  units: PowerPointFontUnit[];
  graphemeEnds: readonly number[];
  eastAsianText: string[];
} {
  const starts: number[] = [];
  let text = '';
  for (const run of runs) {
    starts.push(text.length);
    text += run.text ?? '\n'; // objects and explicit breaks end contextual text
  }
  const bounds = graphemeClusterOffsets(text);
  bounds.push(text.length);
  const units: PowerPointFontUnit[] = [];
  let runIndex = 0;
  const languageAt = (offset: number) => {
    while (runIndex + 1 < starts.length && starts[runIndex + 1] <= offset) runIndex++;
    return runs[runIndex]?.lang;
  };
  let start = 0;
  let previousLanguage = '';
  const isSplitMark = (cp: number) => cp === 0xa9e5 || (cp >= 0xaa7b && cp <= 0xaa7d);
  for (const end of bounds) {
    const lang = languageAt(start);
    const cluster = text.slice(start, end);
    const cp = text.codePointAt(start) ?? 0;
    let slot = powerPointFontSlot(cp, lang, { text, offset: start });
    const language = lang?.toLowerCase() ?? '';
    // Spacing marks may already be separate ICU graphemes. The measured
    // base/mark routing must not depend on that segmentation distinction.
    const previous = units.at(-1);
    if (cluster.length === 1 && isSplitMark(cp)
      && previous?.start === start - 1 && text.charCodeAt(start - 1) === 0x1000
      && MYANMAR_MEASURED_LANGUAGES.has(language) && previousLanguage === language) slot = 'ea';
    if (cluster.length === 2 && cp === 0x1000
      && isSplitMark(cluster.charCodeAt(1))
      && MYANMAR_MEASURED_LANGUAGES.has(language)
      && languageAt(start + 1)?.toLowerCase() === language) {
      units.push({ start, end: start + 1, slot }, { start: start + 1, end, slot: 'ea' });
    } else if (end > start) {
      units.push({ start, end, slot });
    }
    previousLanguage = language;
    start = end;
  }
  // Partition selected EA text by authored run with a forward cursor. Neither
  // a long cross-run cluster nor many short runs may cause quadratic rescans.
  const eastAsianText = runs.map(() => '');
  runIndex = 0;
  for (const unit of units) {
    let offset = unit.start;
    while (offset < unit.end) {
      while (runIndex + 1 < starts.length && starts[runIndex + 1] <= offset) runIndex++;
      const end = Math.min(unit.end, starts[runIndex + 1] ?? text.length);
      if (unit.slot === 'ea' && runs[runIndex]?.text !== null) eastAsianText[runIndex] += text.slice(offset, end);
      offset = end;
    }
  }
  return { text, starts, units, graphemeEnds: bounds, eastAsianText };
}
