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

  it('retains only source-verified normal-style tuples from published open fonts', () => {
    const regular = findReferenceFontMetrics('BIZ UDMincho', {
      source: 'published-open-font', weight: 400, style: 'normal',
    });
    const bold = findReferenceFontMetrics('BIZUDMincho-Bold', {
      source: 'published-open-font', weight: 700, style: 'normal',
    });
    expect(regular).toHaveLength(1);
    expect(bold).toHaveLength(1);
    for (const profile of [regular[0], bold[0]]) {
      expect(profile).toMatchObject({
        unitsPerEm: 2048, hhea: [1802, -246, 0], farEastCodePage: true,
      });
    }
    expect(findReferenceFontMetrics('BIZ UDMincho', {
      source: 'published-open-font', weight: 400, style: 'italic',
    })).toHaveLength(0);
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
