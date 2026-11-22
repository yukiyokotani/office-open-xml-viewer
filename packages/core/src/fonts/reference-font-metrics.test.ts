import { describe, expect, it } from 'vitest';
import { findReferenceFontMetrics } from './reference-font-metrics.js';

describe('findReferenceFontMetrics', () => {
  it('preserves conflicting source profiles instead of choosing a same-name winner', () => {
    const matches = findReferenceFontMetrics('Times New Roman', {
      weight: 400,
      style: 'normal',
    });

    expect(matches.map((profile) => [profile.source, profile.hhea])).toEqual([
      ['office-mac', [1825, -443, 0]],
      ['macos-supplemental', [1825, -443, 87]],
    ]);
  });

  it('matches localized family aliases and PostScript names within a source', () => {
    const localized = findReferenceFontMetrics('ＭＳ 明朝', { source: 'office-mac' });
    expect(localized).toHaveLength(1);
    expect(localized[0]?.family).toBe('MS Mincho');

    const postscript = findReferenceFontMetrics('TimesNewRomanPSMT', {
      source: 'office-mac',
      weight: 400,
      style: 'normal',
    });
    expect(postscript).toHaveLength(1);
    expect(postscript[0]?.family).toBe('Times New Roman');
  });

  it.each(['BIZ UDMincho', 'BIZ UDGothic'])('retains source-verified normal-style tuples for %s', (family) => {
    const regular = findReferenceFontMetrics(family, {
      source: 'published-open-font', weight: 400, style: 'normal',
    });
    const bold = findReferenceFontMetrics(family.replace(' ', '') + '-Bold', {
      source: 'published-open-font', weight: 700, style: 'normal',
    });
    expect(regular).toHaveLength(1);
    expect(bold).toHaveLength(1);
    for (const profile of [regular[0], bold[0]]) {
      expect(profile).toMatchObject({
        unitsPerEm: 2048, hhea: [1802, -246, 0], win: [1802, 246], farEastCodePage: true,
      });
    }
    expect(findReferenceFontMetrics(family, {
      source: 'published-open-font', weight: 400, style: 'italic',
    })).toHaveLength(0);
  });

  it('resolves the Japanese BIZ UD family names to the same pinned profiles', () => {
    for (const [japanese, english] of [
      ['BIZ UDゴシック', 'BIZ UDGothic'],
      ['BIZ UD明朝', 'BIZ UDMincho'],
    ] as const) {
      for (const weight of [400, 700] as const) {
        const localized = findReferenceFontMetrics(japanese, { weight, style: 'normal' });
        expect(localized).toHaveLength(1);
        expect(localized).toEqual(findReferenceFontMetrics(english, { weight, style: 'normal' }));
      }
      expect(findReferenceFontMetrics(japanese, { style: 'italic' })).toHaveLength(0);
    }
  });

  it('returns a format-restricted profile only to lookups for that format', () => {
    for (const family of ['Noto Sans', 'Noto Sans JP']) {
      expect(findReferenceFontMetrics(family)).toHaveLength(0);
      expect(findReferenceFontMetrics(family, { format: 'pptx' })).toHaveLength(0);
      expect(findReferenceFontMetrics(family, { format: 'xlsx', weight: 400, style: 'normal' })).toHaveLength(0);
      const docx = findReferenceFontMetrics(family, { format: 'docx', weight: 400, style: 'normal' });
      expect(docx).toHaveLength(1);
      expect(docx[0]!.evidencedFormats).toEqual(['docx']);
    }
    // Unrestricted profiles answer every lookup.
    expect(findReferenceFontMetrics('Arial', { format: 'pptx', weight: 400, style: 'normal', source: 'office-mac' }))
      .toEqual(findReferenceFontMetrics('Arial', { weight: 400, style: 'normal', source: 'office-mac' }));
  });

  it('does not let callers mutate shared generated profiles', () => {
    const matches = findReferenceFontMetrics('Times New Roman');
    const filtered = findReferenceFontMetrics('Times New Roman', { source: 'office-mac' });
    const filteredAgain = findReferenceFontMetrics('Times New Roman', { source: 'office-mac' });
    const profile = matches[0];
    expect(profile).toBeDefined();
    expect(Object.isFrozen(matches)).toBe(true);
    expect(Object.isFrozen(filtered)).toBe(true);
    expect(filteredAgain).toBe(filtered);
    expect(Object.isFrozen(profile)).toBe(true);
    expect(Object.isFrozen(profile?.aliases)).toBe(true);
    expect(Object.isFrozen(profile?.hhea)).toBe(true);
    // OS/2 code-page evidence remains in the development provenance manifest;
    // the runtime lookup only retains fields consumed by layout.
    expect(profile).not.toHaveProperty('os2');
  });
});
