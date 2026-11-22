/* ------------------------------------------------------------------ *
 * Japanese line-breaking (kinsoku shori / 禁則処理)
 *
 * ECMA-376 §17.3.1.16 `w:kinsoku` is a paragraph property for East Asian
 * typography line-breaking rules. Format adapters own interpretation and
 * inheritance; these shared helpers accept the caller's `enabled` toggle,
 * defaulting to true. The primitive does not resolve Word paragraph
 * properties or assume that the property lives in settings.xml.
 *
 * §17.15.1.58 `w:noLineBreaksAfter` / §17.15.1.59 `w:noLineBreaksBefore`
 * let a document override the character set used by the kinsoku engine
 * for a given language (`w:lang`):
 *   - noLineBreaksBefore (§17.15.1.59): characters that "cannot begin a
 *     line" (行頭禁則 — line-start-forbidden).
 *   - noLineBreaksAfter  (§17.15.1.58): characters that "cannot end a
 *     line" (行末禁則 — line-end-forbidden).
 * The spec states the `w:val` "specifies the set of characters" — it is
 * the COMPLETE set, so a present override REPLACES the application's
 * default set for that language (it does not extend it). When the
 * element is absent the application's own default set is used. We
 * implement replace-vs-default exactly per that wording.
 *
 * The default sets below are Word's documented Japanese kinsoku tables
 * (Tools ▸ Options ▸ Typography ▸ "Use default kinsoku rules"). They
 * coincide with JIS X 4051 §6.1 (行頭禁則文字 / 行末禁則文字). We encode
 * them as two flat string constants (data, not scattered conditionals);
 * membership is a Set lookup.
 *
 * Word applies the document-customizable East-Asian kinsoku to the
 * per-character break path. The `lineStartForbidden` set ALSO contains the
 * ASCII non-starters (!),.:;?]}) — these belong to the UNIVERSAL Latin line
 * breaking rule (UAX#14 LB13: no break before a closing/mid-punctuation), not
 * to East-Asian typography. The renderers' Latin wrap paths consult the DEFAULT
 * set for those (independent of the §17.3.1.16 toggle and of any custom
 * §17.15.1.59 override) so a Latin comma is never orphaned at a line start.
 * ------------------------------------------------------------------ */

/** §17.15.1.59 default 行頭禁則 — characters that may NOT begin a line.
 *  Closing brackets/quotes, mid/end punctuation, small kana, prolonged
 *  sound mark, iteration marks, and their halfwidth forms. */
const KINSOKU_DEFAULT_LINE_START_FORBIDDEN =
  // closing brackets / quotes (fullwidth)
  '”’）〕］｝〉》」』】〙〗〟｠»' +
  // mid / end punctuation (fullwidth)
  '、。，．・：；／？！‐ー゠–〜～' +
  // small kana
  'ぁぃぅぇぉっゃゅょゎゕゖ' +
  'ァィゥェォッャュョヮヵヶ' +
  'ㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿ' +
  // iteration / sound marks
  '々〻ゝゞヽヾ゛゜' +
  // misc trailing symbols
  '％‰℃°′″' +
  // halfwidth forms (cannot start a line either)
  '｡｣､･ｰﾞﾟ' +
  '!),.:;?]}｠';

/** §17.15.1.58 default 行末禁則 — characters that may NOT end a line.
 *  Opening brackets / quotes and currency/lead symbols. */
const KINSOKU_DEFAULT_LINE_END_FORBIDDEN =
  // opening brackets / quotes (fullwidth)
  '“‘（〔［｛〈《「『【〘〖〝｟«' +
  // currency / lead symbols
  '＄￥＃￡￠' +
  // halfwidth opening forms
  '([{｟';

/** Resolved kinsoku configuration for a document.
 *  `enabled` reflects §17.3.1.16; the two sets are §17.15.1.59 (start) /
 *  §17.15.1.58 (end) (custom sets replace the defaults — see
 *  resolveKinsokuRules). */
export interface KinsokuRules {
  enabled: boolean;
  /** Code points forbidden at line START (行頭禁則). */
  lineStartForbidden: Set<number>;
  /** Code points forbidden at line END (行末禁則). */
  lineEndForbidden: Set<number>;
}

function codePointSet(text: string): Set<number> {
  const out = new Set<number>();
  for (const ch of text) out.add(ch.codePointAt(0)!);
  return out;
}

/** Build the active {@link KinsokuRules} from the document settings.
 *  - `enabled` defaults to TRUE when undefined (§17.3.1.16 default).
 *  - A non-undefined custom set REPLACES the default for that direction
 *    (§17.15.1.58 / §17.15.1.59 "specifies the set of characters"). An empty
 *    string is a legitimate replacement that disables that direction.
 */
export function resolveKinsokuRules(settings?: {
  kinsoku?: boolean;
  noLineBreaksBefore?: string;
  noLineBreaksAfter?: string;
}): KinsokuRules {
  return {
    enabled: settings?.kinsoku !== false,
    lineStartForbidden: codePointSet(
      settings?.noLineBreaksBefore ?? KINSOKU_DEFAULT_LINE_START_FORBIDDEN,
    ),
    lineEndForbidden: codePointSet(
      settings?.noLineBreaksAfter ?? KINSOKU_DEFAULT_LINE_END_FORBIDDEN,
    ),
  };
}

/** The default Japanese kinsoku rules (no document overrides). */
export const DEFAULT_KINSOKU_RULES: KinsokuRules = resolveKinsokuRules();
