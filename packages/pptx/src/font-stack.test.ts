import { describe, it, expect } from 'vitest';
import { buildFont, cssFontStack, renderTextBody } from './renderer.js';
import type { TextBody } from './types.js';

describe('cssFontStack — Arabic faces keep the Arabic chain (regression)', () => {
  it('uses Arabic fallbacks only for Arabic text and keeps the default class', () => {
    const stack = cssFontStack('Sakkal Majalla', 'Sakkal Majalla', undefined, 'العربية');
    expect(stack.startsWith('"Sakkal Majalla", "Noto Sans Arabic"')).toBe(true);
    expect(stack).not.toContain('"Noto Naskh Arabic"');
    expect(stack).toContain('"Noto Sans Arabic"');
    // No CJK / non-CJK script tail injected before the generic for Arabic.
    expect(stack).not.toContain('Noto Sans KR');
    expect(stack).not.toContain('Noto Sans Thai');
    expect(cssFontStack('Sakkal Majalla')).not.toContain('Noto Naskh Arabic');
    expect(cssFontStack('Sakkal Majalla', 'Sakkal Majalla', undefined, 'Latin', true))
      .not.toContain('Noto Naskh Arabic');
    expect(cssFontStack('Sakkal Majalla', 'Sakkal Majalla', undefined, 'العربية', true))
      .toContain('"Noto Naskh Arabic"');
  });
});

describe('cssFontStack — CJK language-specific Noto ordering', () => {
  it('routes a native Noto CJK name through the loaded Google family', () => {
    const sans = cssFontStack('Noto Sans CJK SC');
    expect(sans.startsWith('"Noto Sans CJK SC", "Noto Sans SC", ')).toBe(true);
    expect(sans.match(/"Noto Sans SC"/g)).toHaveLength(1);

    const serif = cssFontStack('Noto Serif CJK TC');
    expect(serif.startsWith('"Noto Serif CJK TC", "Noto Serif TC", ')).toBe(true);
    expect(serif.endsWith('serif')).toBe(true);
  });

  it('routes HK sans through Google Fonts without inventing an HK serif alias', () => {
    expect(cssFontStack('Noto Sans CJK HK').startsWith(
      '"Noto Sans CJK HK", "Noto Sans HK", ',
    )).toBe(true);

    const serif = cssFontStack('Noto Serif CJK HK');
    expect(serif).not.toContain('"Noto Serif HK"');
    expect(serif).not.toContain('"Noto Serif TC"');
    expect(serif).not.toMatch(/,\s*,/);
  });

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
    expect(stack).not.toContain('"Caladea"');
    expect(cssFontStack('Cambria', 'Cambria', undefined, '', true)).toContain('"Caladea"');
  });

  it('does not substitute distinct light and mathematical faces', () => {
    expect(cssFontStack('Calibri Light')).not.toContain('"Carlito"');
    expect(cssFontStack('Cambria Math')).not.toContain('"Caladea"');
  });

  it('regression: Calibri stays sans, Times New Roman stays serif', () => {
    expect(cssFontStack('Calibri').endsWith('sans-serif')).toBe(true);
    expect(cssFontStack('Calibri')).not.toContain('"Carlito"');
    expect(cssFontStack('Calibri', 'Calibri', undefined, '', true)).toContain('"Carlito"');
    expect(cssFontStack('Times New Roman').endsWith('serif')).toBe(true);
  });
});

describe('buildFont — style encoded in a face name', () => {
  it('keeps the existing CSS fallback for an unresolved theme-minor face', () => {
    const route = {
      requestedFamily: 'Calibri' as const, family: '__pinned_regular', source: 'substitute' as const,
      resourceIdentity: 'bundled:carlito:test', weight: 400 as const, style: 'normal' as const,
      metric: { family: '__pinned_regular' },
    };
    const rc = {
      themeMajorFont: 'Calibri Light', themeMinorFont: 'Calibri', dpr: 1,
      officeFontRoutes: { calibri: route },
    };
    expect(buildFont(false, false, 24, 'Calibri', rc, 'unstyled', false)).not.toContain('__pinned_regular');
    expect(buildFont(false, false, 24, 'Calibri', rc, 'explicit', true)).toContain('__pinned_regular');
  });

  it('uses only the matching retained Calibri tuple for Canvas selection', () => {
    const route = {
      requestedFamily: 'Calibri' as const,
      family: '__pinned_bold',
      source: 'substitute' as const,
      resourceIdentity: 'bundled:carlito:test',
      weight: 700 as const,
      style: 'normal' as const,
      metric: { family: '__pinned_bold' },
    };
    const rc = {
      themeMajorFont: null, themeMinorFont: 'Calibri', dpr: 1,
      officeFontRoutes: { 'calibri:700:normal': route },
    };
    expect(buildFont(true, false, 24, '+mn-lt', rc)).toContain('"__pinned_bold"');
    expect(buildFont(false, false, 24, '+mn-lt', rc)).not.toContain('"__pinned_bold"');
    expect(buildFont(true, false, 24, 'Arial', rc)).not.toContain('"__pinned_bold"');
  });
  it('keeps an embedded regular Calibri face while routing its uncovered bold style', () => {
    const route = {
      requestedFamily: 'Calibri' as const, family: '__pinned_bold', source: 'substitute' as const,
      resourceIdentity: 'bundled:carlito:test', weight: 700 as const, style: 'normal' as const,
      metric: { family: '__pinned_bold' },
    };
    const rc = {
      themeMajorFont: null, themeMinorFont: 'Calibri', dpr: 1,
      embeddedFontAliases: new Map([['calibri', '__deck_calibri']]),
      embeddedFontAuthoredFamilies: new Map([['__deck_calibri', 'calibri']]),
      embeddedFontTuples: new Set(['calibri:400:normal']),
      officeFontRoutes: { 'calibri:700:normal': route },
    };
    expect(buildFont(false, false, 24, '+mn-lt', rc)).toContain('"__deck_calibri"');
    expect(buildFont(true, false, 24, '+mn-lt', rc)).toContain('"__pinned_bold"');
  });
  it('builds the aliased family stack from a trimmed native name', () => {
    const font = buildFont(false, false, 24, ' Noto Sans CJK SC ', {
      themeMajorFont: null,
      themeMinorFont: null,
      dpr: 1,
    });
    expect(font).toMatch(/^24px "Noto Sans CJK SC", "Noto Sans SC", /);
  });

  it('uses the same resolved alias stack for measurement and paint', () => {
    let font = '';
    let fillStyle = '';
    let direction: CanvasDirection = 'ltr';
    const measuredFonts: string[] = [];
    const paintedFonts: string[] = [];
    const ctx = {
      get font() { return font; },
      set font(value: string) { font = value; },
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string) { fillStyle = value; },
      get direction() { return direction; },
      set direction(value: CanvasDirection) { direction = value; },
      measureText: () => {
        measuredFonts.push(font);
        return {
          width: 40,
          actualBoundingBoxAscent: 16,
          actualBoundingBoxDescent: 4,
          fontBoundingBoxAscent: 16,
          fontBoundingBoxDescent: 4,
        };
      },
      fillText: () => { paintedFonts.push(font); },
      fillRect: () => {},
      drawImage: () => {},
      save: () => {},
      restore: () => {},
      translate: () => {},
      rotate: () => {},
      scale: () => {},
      beginPath: () => {},
      moveTo: () => {},
      lineTo: () => {},
      stroke: () => {},
      clip: () => {},
      rect: () => {},
    } as unknown as CanvasRenderingContext2D;
    const body = {
      verticalAnchor: 't',
      paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0, lvl: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null,
        bullet: { type: 'none' },
        defFontSize: null, defColor: null, defBold: null, defItalic: null,
        defFontFamily: null, tabStops: [], eaLnBrk: true,
        runs: [{
          type: 'text', text: '中文', bold: null, italic: null,
          underline: false, strikethrough: false, fontSize: 20,
          color: '000000', fontFamily: 'Noto Sans CJK SC',
          fontFamilyEa: 'Noto Sans CJK SC',
        }],
      }],
      defaultFontSize: 20, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0,
      wrap: 'none', vert: 'horz', autoFit: 'none',
    } as TextBody;

    renderTextBody(ctx, body, 0, 0, 400, 100, 1 / 12700);

    const resolved = paintedFonts.find((value) => value.includes('"Noto Sans SC"'));
    expect(resolved).toBeDefined();
    expect(measuredFonts).toContain(resolved);
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

  it('keeps the authored serif class behind an embedded alias', () => {
    const font = buildFont(false, false, 24, 'Cambria', {
      themeMajorFont: null,
      themeMinorFont: null,
      embeddedFontAliases: new Map([['cambria', '__ooxml_pptx_3_1']]),
      embeddedFontAuthoredFamilies: new Map([['__ooxml_pptx_3_1', 'cambria']]),
      dpr: 1,
    });
    expect(font).toContain('"__ooxml_pptx_3_1"');
    expect(font).not.toContain('"Caladea"');
    expect(font.endsWith('serif')).toBe(true);
  });

  it('preserves a Medium theme face when the browser falls back', () => {
    const font = buildFont(false, false, 48, 'Franklin Gothic Medium', {
      themeMajorFont: null,
      themeMinorFont: null,
      dpr: 1,
    });
    expect(font).toMatch(/^600 48px "Franklin Gothic Medium"/);
    expect(font).not.toContain('"Libre Franklin"');
    expect(buildFont(false, false, 48, 'Franklin Gothic Medium', {
      themeMajorFont: null, themeMinorFont: null, dpr: 1, googleSubstitutes: true,
    })).toContain('"Libre Franklin"');
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


it('uses per-presentation fallback for neutral fonts while preserving named regions', () => {
  const sc = cssFontStack('Calibri', 'Calibri', 'sc');
  expect(sc).not.toContain('"Carlito"');
  const optedIn = cssFontStack('Calibri', 'Calibri', 'sc', '', true);
  expect(optedIn.indexOf('"Carlito"')).toBeLessThan(optedIn.indexOf('"Noto Sans SC"'));
  expect(sc.indexOf('"Noto Sans SC"')).toBeLessThan(sc.indexOf('"Noto Sans JP"'));
  const tc = cssFontStack('Calibri', 'Calibri', 'tc');
  expect(tc.indexOf('"Noto Sans TC"')).toBeLessThan(tc.indexOf('"Noto Sans SC"'));
  const jp = cssFontStack('Meiryo', 'Meiryo', 'sc');
  expect(jp.indexOf('"Noto Sans JP"')).toBeLessThan(jp.indexOf('"Noto Sans SC"'));
  expect(buildFont(false, false, 12, 'Arial', {
    themeMajorFont: 'Meiryo', themeMinorFont: 'Calibri', dpr: 1,
  }, '漢')).toContain('"Noto Sans JP", "Noto Sans SC"');
  expect(buildFont(false, false, 12, 'Arial', {
    themeMajorFont: 'Meiryo', themeMinorFont: 'Calibri', dpr: 1,
  }, '→')).not.toContain('Noto Sans JP');
});


it('keeps Arabic substitutes ahead of the configured CJK fallback', () => {
  const stack = cssFontStack('Amiri', 'Amiri', 'sc', 'العربية');
  expect(stack.indexOf('Noto Naskh Arabic')).toBeLessThan(stack.indexOf('Noto Sans SC'));
  expect(stack).toContain('"Noto Sans SC", "Noto Sans TC"');
});
