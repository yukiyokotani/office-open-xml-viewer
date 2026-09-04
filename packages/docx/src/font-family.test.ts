import { describe, it, expect } from 'vitest';
import { normalizeFontFamily, normalizeFontFamilyUncached } from './line-layout.js';

describe('normalizeFontFamily — modern family respects w:pitch (§17.8.3.14)', () => {
  it('keeps a variable-pitch modern face on the Meiryo/CJK-sans path', () => {
    const chain = normalizeFontFamilyUncached(
      'Meiryo UI',
      { 'Meiryo UI': 'modern' },
      { 'Meiryo UI': 'variable' },
    );

    expect(chain).not.toContain('monospace');
    expect(chain).toContain('"Meiryo UI"');
  });

  it('maps a fixed-pitch modern face to monospace', () => {
    const chain = normalizeFontFamilyUncached(
      'MonoFace',
      { MonoFace: 'modern' },
      { MonoFace: 'fixed' },
    );

    expect(chain).toContain('"Courier New", monospace');
  });

  it('keeps a CJK fallback ahead of Latin monospace for fixed East Asian faces', () => {
    const chain = normalizeFontFamilyUncached(
      'ＭＳ ゴシック',
      { 'ＭＳ ゴシック': 'modern' },
      { 'ＭＳ ゴシック': 'fixed' },
    );

    // East-Asia-hinted ambiguous glyphs such as U+25A1 must retain the
    // full-width glyph of the authored CJK face when that face is unavailable;
    // Courier's narrow square must not capture them first.
    expect(chain).toContain('"Yu Gothic"');
    expect(chain.indexOf('"Yu Gothic"')).toBeLessThan(chain.indexOf('"Courier New"'));
    expect(chain).toContain('monospace');
  });

  it('does not assume an omitted pitch is fixed', () => {
    const chain = normalizeFontFamilyUncached('Meiryo UI', { 'Meiryo UI': 'modern' });

    expect(chain).not.toContain('monospace');
    expect(chain).toContain('"Meiryo UI"');
  });
});

describe('normalizeFontFamily — Arabic substitute fonts', () => {
  it('puts the Arabic substitute first so Latin/digits resolve from the same family as Arabic', () => {
    // Sakkal Majalla is family="auto" in fontTable; the run carries both
    // Arabic glyphs and Latin/digits. The Arabic substitute must lead the chain
    // (before any CJK sans face) so Latin/digits don't leak to Noto Sans JP.
    const chain = normalizeFontFamily('Sakkal Majalla');
    expect(chain.startsWith('"Sakkal Majalla", "Noto Naskh Arabic"')).toBe(true);
    // No CJK sans face before the Arabic substitute.
    const naskhIdx = chain.indexOf('Noto Naskh Arabic');
    const cjkIdx = chain.indexOf('Noto Sans JP');
    expect(naskhIdx).toBeGreaterThan(0);
    expect(naskhIdx).toBeLessThan(cjkIdx);
  });

  it('routes traditional Naskh faces to a serif Latin companion', () => {
    // Word's PDF export of sample-7 renders Sakkal Majalla's Latin with serifs,
    // so a serif Latin generic precedes the sans generics.
    const chain = normalizeFontFamily('Traditional Arabic');
    expect(chain).toContain('"Noto Serif"');
    expect(chain.endsWith('serif')).toBe(true);
    expect(chain.indexOf('Noto Serif')).toBeLessThan(chain.indexOf('Noto Sans JP'));
  });

  it('keys off the family, not a hardcoded string — case-insensitive', () => {
    expect(normalizeFontFamily('sakkal majalla').startsWith('"sakkal majalla", "Noto Naskh Arabic"')).toBe(true);
  });

  it('routes geometric Arabic faces (Univers Next Arabic) to a sans chain', () => {
    const chain = normalizeFontFamily('Univers Next Arabic');
    expect(chain.startsWith('"Univers Next Arabic", "Noto Sans Arabic"')).toBe(true);
    expect(chain.endsWith('sans-serif')).toBe(true);
  });
});

describe('normalizeFontFamily — Latin fonts lead with Latin faces, JP companion follows', () => {
  it('leads a Latin sans font with Latin sans faces; JP companion follows for stray CJK', () => {
    const chain = normalizeFontFamily('Arial');
    expect(chain.startsWith('"Arial"')).toBe(true);
    // Latin letters/digits must resolve to a Latin sans, NOT a Japanese Gothic's
    // wider Latin glyphs — so the Latin companions precede the JP companion.
    expect(chain.indexOf('"Helvetica"')).toBeLessThan(chain.indexOf('"Noto Sans JP"'));
    // …but the JP companion (and non-CJK script Notos) are still present so a
    // stray CJK / non-Latin glyph degrades to a real web font.
    expect(chain).toContain('"Noto Sans JP"');
    expect(chain).toContain('"Noto Naskh Arabic"');
    expect(chain).toContain('"Noto Sans Hebrew"');
    expect(chain).toContain('"Noto Sans Thai"');
    expect(chain).toContain('"Noto Sans Devanagari"');
    expect(chain.endsWith('sans-serif')).toBe(true);
  });

  it('leads a Latin serif font with Latin serif faces; mincho companion follows for stray CJK', () => {
    const chain = normalizeFontFamily('Times New Roman');
    expect(chain.startsWith('"Times New Roman"')).toBe(true);
    // Latin glyphs resolve to a Latin serif before the (wider-Latin) JP mincho.
    expect(chain.indexOf('"Cambria"')).toBeLessThan(chain.indexOf('"Yu Mincho"'));
    expect(chain).toContain('"Yu Mincho"');
    expect(chain).toContain('"Noto Serif JP"');
    expect(chain).toContain('"Noto Serif Hebrew"');
    expect(chain.endsWith('serif')).toBe(true);
  });
});

describe('normalizeFontFamily — CJK language-specific Noto ordering', () => {
  it('routes native Noto CJK names through the loaded Google family', () => {
    const sans = normalizeFontFamily('Noto Sans CJK SC');
    expect(sans.startsWith('"Noto Sans CJK SC", "Noto Sans SC", ')).toBe(true);

    const serif = normalizeFontFamily('Noto Serif CJK TC');
    expect(serif.startsWith('"Noto Serif CJK TC", "Noto Serif TC", ')).toBe(true);
    expect(serif.endsWith('serif')).toBe(true);
  });

  it('routes HK sans through Google Fonts without inventing an HK serif alias', () => {
    expect(normalizeFontFamily('Noto Sans CJK HK').startsWith(
      '"Noto Sans CJK HK", "Noto Sans HK", ',
    )).toBe(true);

    const serif = normalizeFontFamily('Noto Serif CJK HK');
    expect(serif).not.toContain('"Noto Serif HK"');
    expect(serif).not.toContain('"Noto Serif TC"');
  });

  it('puts Noto Sans KR first for Korean sans faces (Malgun Gothic, Gulim, Dotum, 돋움)', () => {
    for (const f of ['Malgun Gothic', 'Gulim', 'Dotum', '돋움']) {
      const chain = normalizeFontFamily(f);
      expect(chain, `${f}`).toContain('"Noto Sans KR"');
      // KR must precede JP so shared Han renders with Korean shapes.
      expect(chain.indexOf('Noto Sans KR')).toBeLessThan(
        chain.indexOf('Noto Sans JP') === -1 ? Infinity : chain.indexOf('Noto Sans JP'),
      );
    }
  });

  it('routes Korean serif faces (Batang) to Noto Serif KR', () => {
    const chain = normalizeFontFamily('Batang');
    expect(chain).toContain('"Noto Serif KR"');
    expect(chain.endsWith('serif')).toBe(true);
    expect(chain.indexOf('Noto Serif KR')).toBeLessThan(chain.indexOf('Noto Serif JP'));
  });

  it('puts Noto Sans SC first for Simplified Chinese faces (SimSun→serif, YaHei→sans)', () => {
    // SimSun is a song (serif) face → Noto Serif SC.
    const simsun = normalizeFontFamily('SimSun');
    expect(simsun).toContain('"Noto Serif SC"');
    expect(simsun.indexOf('Noto Serif SC')).toBeLessThan(
      simsun.indexOf('Noto Serif JP') === -1 ? Infinity : simsun.indexOf('Noto Serif JP'),
    );
    // Microsoft YaHei is a sans face → Noto Sans SC.
    const yahei = normalizeFontFamily('Microsoft YaHei');
    expect(yahei).toContain('"Noto Sans SC"');
    expect(yahei.indexOf('Noto Sans SC')).toBeLessThan(
      yahei.indexOf('Noto Sans JP') === -1 ? Infinity : yahei.indexOf('Noto Sans JP'),
    );
  });

  it('puts Noto Sans TC first for Traditional Chinese faces (PMingLiU→serif, JhengHei→sans)', () => {
    const pming = normalizeFontFamily('PMingLiU');
    expect(pming).toContain('"Noto Serif TC"');
    const jheng = normalizeFontFamily('Microsoft JhengHei');
    expect(jheng).toContain('"Noto Sans TC"');
    expect(jheng.indexOf('Noto Sans TC')).toBeLessThan(
      jheng.indexOf('Noto Sans SC') === -1 ? Infinity : jheng.indexOf('Noto Sans SC'),
    );
  });

  it('keeps Japanese faces on Noto JP (regression — Yu Gothic, Meiryo, MS Mincho)', () => {
    expect(normalizeFontFamily('Yu Gothic')).toContain('"Noto Sans JP"');
    expect(normalizeFontFamily('Meiryo')).toContain('"Noto Sans JP"');
    // MS Mincho is serif → Noto Serif JP.
    expect(normalizeFontFamily('MS Mincho')).toContain('"Noto Serif JP"');
  });
});

// B6: normalizeFontFamily is memoized per-document (keyed on the
// fontFamilyClasses object identity). These pin that memoization is transparent
// — identical inputs give identical results, and two different fontFamilyClasses
// objects never share a cache entry (so the fontTable §17.8.3.10 classification,
// which lives in fontFamilyClasses, still switches the chain).
describe('normalizeFontFamily — per-document memoization is transparent', () => {
  it('returns the identical string on repeated calls with the same classes object', () => {
    const classes = { Arial: 'roman' };
    const a = normalizeFontFamily('Arial', classes);
    const b = normalizeFontFamily('Arial', classes);
    expect(b).toBe(a);
    // roman ⇒ serif tail, so the cached result is the serif chain.
    expect(a).toContain('"Arial"');
    expect(a).toContain('serif');
  });

  it('does not mix results across different fontFamilyClasses objects', () => {
    // Same family, two different fontTable classifications: `swiss` (sans) vs
    // `roman` (serif). The memo must not serve one doc's result to the other.
    const asSwiss = normalizeFontFamily('Calibri', { Calibri: 'swiss' });
    const asRoman = normalizeFontFamily('Calibri', { Calibri: 'roman' });
    expect(asSwiss).not.toBe(asRoman);
    expect(asSwiss.endsWith('serif')).toBe(true); // sans tail ends "…, sans-serif"
    expect(asSwiss).not.toContain('Noto Serif');
    expect(asRoman).toContain('Noto Serif');
    // Re-querying the first object still yields the swiss (sans) chain, proving
    // the second object's entry did not overwrite it.
    expect(normalizeFontFamily('Calibri', { Calibri: 'swiss' })).not.toBe(asRoman);
  });

  it('handles a null family through the cache without colliding with a real "null" family', () => {
    const classes = {};
    const nullChain = normalizeFontFamily(null, classes);
    // Stable across calls.
    expect(normalizeFontFamily(null, classes)).toBe(nullChain);
    // A real face literally named "null" must not be served the null-family result.
    const namedNull = normalizeFontFamily('null', classes);
    expect(namedNull).toContain('"null"');
    expect(namedNull).not.toBe(nullChain);
  });
});
