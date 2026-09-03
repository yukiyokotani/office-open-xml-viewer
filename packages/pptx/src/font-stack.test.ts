import { describe, it, expect } from 'vitest';
import { buildFont, cssFontStack } from './renderer.js';

describe('cssFontStack — Arabic faces keep the Arabic chain (regression)', () => {
  it('leads with the Arabic Noto fallbacks for an Arabic-script face', () => {
    // OFFICE_FONT_SUBSTITUTE maps Sakkal Majalla → Noto Naskh Arabic.
    const stack = cssFontStack('Sakkal Majalla');
    expect(stack.startsWith('"Sakkal Majalla", "Noto Naskh Arabic"')).toBe(true);
    expect(stack).toContain('"Noto Sans Arabic"');
    // No CJK / non-CJK script tail injected before the generic for Arabic.
    expect(stack).not.toContain('Noto Sans KR');
    expect(stack).not.toContain('Noto Sans Thai');
  });
});

describe('cssFontStack — CJK language-specific Noto ordering', () => {
  it('Korean sans (Malgun Gothic) → Noto Sans KR leads the CJK tail', () => {
    const stack = cssFontStack('Malgun Gothic');
    expect(stack).toContain('"Noto Sans KR"');
    expect(stack.indexOf('Noto Sans KR')).toBeLessThan(stack.indexOf('Noto Sans JP'));
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('Simplified Chinese serif (SimSun) → Noto Serif SC leads', () => {
    const stack = cssFontStack('SimSun');
    expect(stack).toContain('"Noto Serif SC"');
    expect(stack.indexOf('Noto Serif SC')).toBeLessThan(stack.indexOf('Noto Serif JP'));
    expect(stack.endsWith('serif')).toBe(true);
  });

  it('Simplified Chinese sans (Microsoft YaHei) → Noto Sans SC leads', () => {
    const stack = cssFontStack('Microsoft YaHei');
    expect(stack).toContain('"Noto Sans SC"');
    expect(stack.indexOf('Noto Sans SC')).toBeLessThan(stack.indexOf('Noto Sans JP'));
  });

  it('Traditional Chinese (Microsoft JhengHei) → Noto Sans TC leads', () => {
    const stack = cssFontStack('Microsoft JhengHei');
    expect(stack).toContain('"Noto Sans TC"');
    expect(stack.indexOf('Noto Sans TC')).toBeLessThan(stack.indexOf('Noto Sans SC'));
  });

  it('Japanese faces stay on Noto JP (regression — Yu Gothic, Meiryo)', () => {
    expect(cssFontStack('Yu Gothic')).toContain('"Noto Sans JP"');
    expect(cssFontStack('Meiryo')).toContain('"Noto Sans JP"');
  });
});

describe('cssFontStack — non-CJK scripts appended to Latin faces', () => {
  it('adds Hebrew / Thai / Devanagari Notos to a plain Latin sans face', () => {
    const stack = cssFontStack('Arial');
    expect(stack).toContain('"Noto Sans Hebrew"');
    expect(stack).toContain('"Noto Sans Thai"');
    expect(stack).toContain('"Noto Sans Devanagari"');
    expect(stack).toContain('"Noto Sans"'); // Cyrillic coverage
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('adds Hebrew serif Noto to a serif face', () => {
    const stack = cssFontStack('Times New Roman');
    expect(stack).toContain('"Noto Serif Hebrew"');
    expect(stack).toContain('"Noto Serif"');
    expect(stack.endsWith('serif')).toBe(true);
  });
});

describe('cssFontStack — serif/sans generic classification (core classifier)', () => {
  it('Cambria degrades to a serif (latent pptx fix — was sans-serif)', () => {
    const stack = cssFontStack('Cambria');
    expect(stack.endsWith('serif')).toBe(true);
    expect(stack.endsWith('sans-serif')).toBe(false);
    // Cambria's metric-compatible substitute is appended (OFFICE_FONT_SUBSTITUTE).
    expect(stack).toContain('"Caladea"');
  });

  it('regression: Calibri stays sans, Times New Roman stays serif', () => {
    expect(cssFontStack('Calibri').endsWith('sans-serif')).toBe(true);
    expect(cssFontStack('Times New Roman').endsWith('serif')).toBe(true);
  });
});

describe('cssFontStack — private provider fallback', () => {
  it('keeps an installed authored face ahead of its isolated provider alias', () => {
    expect(cssFontStack('Calibri', 'Calibri', '__private_calibri'))
      .toMatch(/^"Calibri", "__private_calibri", "Carlito"/);
  });
});

describe('buildFont — style encoded in a face name', () => {
  it('uses the provider alias after the authored family', () => {
    const font = buildFont(true, false, 24, 'Deck Sans', {
      themeMajorFont: null,
      themeMinorFont: null,
      providerFontRoutes: { 'deck sans': '__private_deck_sans' },
      dpr: 1,
    });
    expect(font).toContain('"Deck Sans", "__private_deck_sans"');
  });

  it('selects a presentation-scoped embedded alias instead of the global authored family', () => {
    const font = buildFont(false, false, 24, 'Deck Sans', {
      themeMajorFont: null,
      themeMinorFont: null,
      embeddedFontAliases: new Map([['deck sans', '__ooxml_pptx_1_1']]),
      embeddedFontAuthoredFamilies: new Map([['__ooxml_pptx_1_1', 'deck sans']]),
      dpr: 1,
    });
    expect(font).toContain('"__ooxml_pptx_1_1"');
    expect(font).not.toContain('"Deck Sans"');
  });

  it('selects the embedded alias after resolving a theme font reference', () => {
    const font = buildFont(false, false, 24, '+mn-lt', {
      themeMajorFont: null,
      themeMinorFont: 'Deck Sans',
      embeddedFontAliases: new Map([['deck sans', '__ooxml_pptx_2_1']]),
      embeddedFontAuthoredFamilies: new Map([['__ooxml_pptx_2_1', 'deck sans']]),
      dpr: 1,
    });
    expect(font).toContain('"__ooxml_pptx_2_1"');
    expect(font).not.toContain('"Deck Sans"');
  });

  it('keeps the authored serif and substitute policy behind an embedded alias', () => {
    const font = buildFont(false, false, 24, 'Cambria', {
      themeMajorFont: null,
      themeMinorFont: null,
      embeddedFontAliases: new Map([['cambria', '__ooxml_pptx_3_1']]),
      embeddedFontAuthoredFamilies: new Map([['__ooxml_pptx_3_1', 'cambria']]),
      dpr: 1,
    });
    expect(font).toContain('"__ooxml_pptx_3_1"');
    expect(font).toContain('"Caladea"');
    expect(font.endsWith('serif')).toBe(true);
  });

  it('preserves a Medium theme face when the browser falls back', () => {
    const font = buildFont(false, false, 48, 'Franklin Gothic Medium', {
      themeMajorFont: null,
      themeMinorFont: null,
      dpr: 1,
    });
    expect(font).toMatch(/^600 48px "Franklin Gothic Medium"/);
    expect(font).toContain('"Libre Franklin"');
  });

  it('lets an explicit bold run override a named Medium face', () => {
    const font = buildFont(true, false, 48, 'Franklin Gothic Medium', {
      themeMajorFont: null,
      themeMinorFont: null,
      dpr: 1,
    });
    expect(font).toMatch(/^bold 48px "Franklin Gothic Medium"/);
  });
});
