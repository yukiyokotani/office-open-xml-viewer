import { describe, expect, it } from 'vitest';
import { POWERPOINT_FONT_SLOT_EVIDENCE, powerPointFontSlot } from './font-slot-compatibility.js';
import { POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE } from './font-slot-boundary-evidence.js';
import { POWERPOINT_EXTRA_FONT_SLOT_EVIDENCE } from './font-slot-evidence.js';

describe('PowerPoint slot compatibility evidence', () => {
  it('matches original observations except explicitly withdrawn boundary overrides', () => {
    // This independently recorded PDF corpus catches broadened symbol ranges,
    // lost language overrides and inverted endpoints; fallback fonts are not a
    // slot oracle. Renderer wiring is exercised in font-slot-render.test.ts.
    for (const [lang, ranges] of Object.entries(POWERPOINT_FONT_SLOT_EVIDENCE.coverage)) {
      for (const [start, end, outcome] of ranges) {
        if (outcome !== 'latin' && outcome !== 'ea' && outcome !== 'cs') continue;
        for (let cp = start; cp <= end; cp++) {
          // Preserve the historical evidence; replacement cycles reject these
          // overrides only for the two original symbol-sweep language IDs.
          if ((lang === 'en-us' || lang === 'ja-jp')
            && POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE.withdrawnOverrides.some((withdrawn) => withdrawn === cp)) continue;
          expect(powerPointFontSlot(cp, lang), `${lang} U+${cp.toString(16)}`).toBe(outcome);
        }
      }
    }
  });

  it('matches independently extracted cyclic controls with consistent scalar outcomes', () => {
    for (const [lang, ranges] of Object.entries(POWERPOINT_EXTRA_FONT_SLOT_EVIDENCE.coverage)) {
      for (const [start, end, outcome] of ranges) {
        for (let cp = start; cp <= end; cp++) {
          expect(powerPointFontSlot(cp, lang), `${lang} U+${cp.toString(16)}`).toBe(outcome);
        }
      }
    }
  });

  it('accounts for the complete sweep, including fallback and unextractable glyphs', () => {
    const counts = Object.fromEntries(Object.entries(POWERPOINT_FONT_SLOT_EVIDENCE.coverage)
      .map(([lang, ranges]) => [lang, ranges.reduce((sum, [start, end]) => sum + end - start + 1, 0)]));
    expect(counts).toEqual({ 'en-us': 3326, 'ja-jp': 3326, 'ko-kr': 339,
      'zh-cn': 339, 'zh-tw': 339, 'he-il': 339, 'ar-sa': 339, 'th-th': 339, 'hi-in': 339 });
  });

  it('uses normative defaults for unextractable symbols and retains unmeasured script policy', () => {
    expect(powerPointFontSlot(0x25aa, 'en-US')).toBe('ea');
    expect(powerPointFontSlot(0x2049, 'en-US')).toBe('ea');
    expect(powerPointFontSlot(0x201e, 'zh-HK')).toBe('ea');
    expect(powerPointFontSlot(0x1800, 'en-US')).toBe('latin'); // Mongolian: no new compatibility claim
    expect(powerPointFontSlot(0xa000, 'en-US')).toBe('latin'); // Yi: pending distinct-face controls
    expect(powerPointFontSlot(0xfe8e, 'ar-SA')).toBe('cs'); // retain Arabic shaping-script routing
    expect(powerPointFontSlot(0x1100, 'en-US')).toBe('ea'); // settled Jamo role permutation
    expect(powerPointFontSlot(0xf000, 'en-US')).toBe('sym');
    expect(powerPointFontSlot(0xf0ff, 'en-US')).toBe('sym');
    expect(powerPointFontSlot(0xf100, 'en-US')).toBe('latin');
    expect(powerPointFontSlot(0x31, 'AR-sa')).toBe('cs');
    expect(powerPointFontSlot(0xa9e5, 'en-US')).toBe('ea'); // exact standalone-mark cycles
    expect(powerPointFontSlot(0xa9ff, 'en-US')).toBe('latin'); // unassigned gap
    expect(powerPointFontSlot(0xaa7b, 'my-MM')).toBe('ea'); // exact standalone-mark cycles
    expect(powerPointFontSlot(0xa9e0, 'fr-FR')).toBe('cs'); // unmeasured language retains policy
    expect(powerPointFontSlot(0xbb, 'he-IL')).toBe('cs'); // original sweep routing
    expect(powerPointFontSlot(0x30, 'fa')).toBe('latin'); // do not infer other region/language IDs
    expect(powerPointFontSlot(0x31, 'constructor')).toBe('latin'); // untrusted document language
  });
});
