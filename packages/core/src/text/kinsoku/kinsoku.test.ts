import { describe, it, expect } from 'vitest';
import {
  kinsokuAdjustedSplit,
  resolveKinsokuRules,
  DEFAULT_KINSOKU_RULES,
} from './index.js';

// ECMA-376 §17.15.1.58–.60 Japanese line-breaking (kinsoku) tests.
//
// kinsokuAdjustedSplit operates on an array of single code points and a
// proposed split index `splitAt`: chars[0..splitAt) stay on the current line
// (head), chars[splitAt..] overflow to the next line (tail). It retracts the
// split leftwards so the tail never begins with a 行頭禁則 char and the head
// never ends with a 行末禁則 char.

const cp = (s: string): string[] => [...s];

describe('kinsokuAdjustedSplit — line-start-forbidden (行頭禁則 / 追い出し)', () => {
  it('pulls the preceding char down so a line never begins with 、', () => {
    // "あいの、うえ" with a natural break that would put 、 at the tail start.
    // chars: あ(0) い(1) の(2) 、(3) う(4) え(5)
    // raw break after の (splitAt=3) → tail = "、うえ" begins with 、(forbidden).
    const chars = cp('あいの、うえ');
    const adjusted = kinsokuAdjustedSplit(chars, 3, DEFAULT_KINSOKU_RULES);
    // Retract one: split=2 → head="あい", tail="の、うえ" begins with の (legal).
    expect(adjusted).toBe(2);
    expect(chars.slice(0, adjusted).join('')).toBe('あい');
    expect(chars.slice(adjusted).join('')).toBe('の、うえ');
  });

  it('never begins a line with 。', () => {
    // raw break before 。 ⇒ retract.
    const chars = cp('簡潔に。説明');
    // chars: 簡(0)潔(1)に(2)。(3)説(4)明(5); break at 3 → tail begins 。
    const adjusted = kinsokuAdjustedSplit(chars, 3, DEFAULT_KINSOKU_RULES);
    expect(adjusted).toBe(2);
    expect(chars.slice(adjusted).join('').startsWith('。')).toBe(false);
  });

  it('retracts past consecutive forbidden chars', () => {
    // "あ、。い" break at 1 → tail "、。い" forbidden; but break is already 1.
    // Use break at 3: あ(0)、(1)。(2)い(3) — tail "い" legal already.
    // Construct a case where two forbidden chars are at the tail head:
    // break at 2 → tail "。い" forbidden → retract to 1 → tail "、。い" forbidden
    // → retract to 0... but minSplit default 1 floors at 1; still forbidden →
    // fall back to original split (2).
    const chars = cp('あ、。い');
    const adjusted = kinsokuAdjustedSplit(chars, 2, DEFAULT_KINSOKU_RULES);
    // floor reached while still illegal → unrestricted fallback.
    expect(adjusted).toBe(2);
  });
});

describe('kinsokuAdjustedSplit — line-end-forbidden (行末禁則)', () => {
  it('does not leave an opening bracket 「 as the last char of a line', () => {
    // あい「うえ: break at 3 → head "あい「" ends with 「 (forbidden at line end).
    const chars = cp('あい「うえ');
    const adjusted = kinsokuAdjustedSplit(chars, 3, DEFAULT_KINSOKU_RULES);
    // retract to 2 → head "あい", tail "「うえ" (legal: 「 may begin a line).
    expect(adjusted).toBe(2);
    expect(chars.slice(0, adjusted).join('').endsWith('「')).toBe(false);
  });

  it('does not leave （ as the last char of a line', () => {
    const chars = cp('テスト（内容');
    // テ(0)ス(1)ト(2)（(3)内(4)容(5); break at 4 → head ends with （.
    const adjusted = kinsokuAdjustedSplit(chars, 4, DEFAULT_KINSOKU_RULES);
    expect(adjusted).toBe(3);
    expect(chars.slice(0, adjusted).join('').endsWith('（')).toBe(false);
  });
});

describe('kinsokuAdjustedSplit — bounds & no-op cases', () => {
  it('returns splitAt unchanged at the head edge (splitAt<=0)', () => {
    expect(kinsokuAdjustedSplit(cp('、あい'), 0, DEFAULT_KINSOKU_RULES)).toBe(0);
  });

  it('returns splitAt unchanged at the tail edge (splitAt>=len)', () => {
    const chars = cp('あいう');
    expect(kinsokuAdjustedSplit(chars, 3, DEFAULT_KINSOKU_RULES)).toBe(3);
  });

  it('leaves a legal split untouched', () => {
    const chars = cp('あいうえお');
    expect(kinsokuAdjustedSplit(chars, 2, DEFAULT_KINSOKU_RULES)).toBe(2);
  });

  it('pathological all-forbidden run falls back without hanging', () => {
    // Every tail-start is forbidden; must terminate and return the input split.
    const chars = cp('、。、。、。');
    const adjusted = kinsokuAdjustedSplit(chars, 3, DEFAULT_KINSOKU_RULES);
    expect(adjusted).toBe(3); // unrestricted fallback, no infinite loop
  });

  it('allows retraction to an empty head when minSplit=0 (line already has content)', () => {
    // chars: の(0)、(1)あ(2); break at 1 → tail "、あ" forbidden → retract to 0.
    const chars = cp('の、あ');
    const adjusted = kinsokuAdjustedSplit(chars, 1, DEFAULT_KINSOKU_RULES, 0);
    expect(adjusted).toBe(0); // whole run pushed to next line (追い出し)
  });
});

describe('resolveKinsokuRules — §17.15.1.58 toggle', () => {
  it('defaults kinsoku to ON when settings absent', () => {
    expect(resolveKinsokuRules().enabled).toBe(true);
    expect(resolveKinsokuRules(undefined).enabled).toBe(true);
    expect(resolveKinsokuRules({}).enabled).toBe(true);
  });

  it('kinsoku=false (w:val="0") disables', () => {
    const rules = resolveKinsokuRules({ kinsoku: false });
    expect(rules.enabled).toBe(false);
    // When disabled, the split is never adjusted.
    expect(kinsokuAdjustedSplit(cp('あいの、うえ'), 3, rules)).toBe(3);
  });

  it('default sets include the reported start-forbidden chars', () => {
    const rules = resolveKinsokuRules();
    for (const ch of '、。」）') {
      expect(rules.lineStartForbidden.has(ch.codePointAt(0)!)).toBe(true);
    }
  });

  it('default sets include the reported end-forbidden chars', () => {
    const rules = resolveKinsokuRules();
    for (const ch of '「（《') {
      expect(rules.lineEndForbidden.has(ch.codePointAt(0)!)).toBe(true);
    }
  });
});

describe('resolveKinsokuRules — §17.15.1.59/.60 custom sets REPLACE defaults', () => {
  it('custom noLineBreaksBefore replaces the default start set', () => {
    const rules = resolveKinsokuRules({ noLineBreaksBefore: '〇' });
    // The custom char is now start-forbidden...
    expect(rules.lineStartForbidden.has('〇'.codePointAt(0)!)).toBe(true);
    // ...and the default 、 is NO LONGER forbidden (replaced, not extended).
    expect(rules.lineStartForbidden.has('、'.codePointAt(0)!)).toBe(false);
    // So a break leaving 、 at line start is now legal (not retracted).
    expect(kinsokuAdjustedSplit(cp('あいの、うえ'), 3, rules)).toBe(3);
    // But a break leaving 〇 at line start IS retracted.
    expect(kinsokuAdjustedSplit(cp('あい〇うえ'), 2, rules)).toBe(1);
  });

  it('custom noLineBreaksAfter replaces the default end set', () => {
    const rules = resolveKinsokuRules({ noLineBreaksAfter: '＠' });
    expect(rules.lineEndForbidden.has('＠'.codePointAt(0)!)).toBe(true);
    expect(rules.lineEndForbidden.has('「'.codePointAt(0)!)).toBe(false);
  });

  it('empty custom set is a legitimate replacement (disables that direction)', () => {
    const rules = resolveKinsokuRules({ noLineBreaksBefore: '' });
    expect(rules.lineStartForbidden.size).toBe(0);
    expect(kinsokuAdjustedSplit(cp('あいの、うえ'), 3, rules)).toBe(3);
  });
});
