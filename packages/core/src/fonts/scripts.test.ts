import { describe, it, expect } from 'vitest';
import {
  classifyCjkFont,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  GOOGLE_CJK_FONT_ALIASES,
  googleCjkFontAlias,
  SCRIPT_GOOGLE_FONTS,
  SCRIPT_PRELOAD_NAMES,
  ScriptPreloadAccumulator,
  scriptPreloadNamesForText,
} from './scripts.js';

/** Lower-case every name and assert it is resolvable in the preload URL map. */
function assertResolvable(names: string[]): void {
  for (const n of names) {
    expect(
      SCRIPT_GOOGLE_FONTS[n.toLowerCase()],
      `${n} missing from SCRIPT_GOOGLE_FONTS`,
    ).toBeDefined();
  }
}

describe('classifyCjkFont — Office font name → CJK language', () => {
  it('classifies native Noto CJK names across case and surrounding whitespace', () => {
    expect(classifyCjkFont(' Noto Sans CJK SC ')).toBe('sc');
    expect(classifyCjkFont('nOtO SeRiF CjK Tc')).toBe('tc');
    expect(classifyCjkFont('Noto Sans CJK JP')).toBe('jp');
    expect(classifyCjkFont('Noto Serif CJK KR')).toBe('kr');
    expect(classifyCjkFont(' Noto Sans CJK HK ')).toBe('hk');
    expect(classifyCjkFont('Noto Serif CJK HK')).toBe('hk');
  });

  it('classifies Korean faces (Malgun Gothic, Batang, Gulim, Dotum, 돋움)', () => {
    expect(classifyCjkFont('Malgun Gothic')).toBe('kr');
    expect(classifyCjkFont('Batang')).toBe('kr');
    expect(classifyCjkFont('Gulim')).toBe('kr');
    expect(classifyCjkFont('Dotum')).toBe('kr');
    expect(classifyCjkFont('돋움')).toBe('kr');
    expect(classifyCjkFont('맑은 고딕')).toBe('kr');
  });

  it('classifies Simplified Chinese faces (SimSun, 宋体, Microsoft YaHei, DengXian, FangSong, KaiTi)', () => {
    expect(classifyCjkFont('SimSun')).toBe('sc');
    expect(classifyCjkFont('宋体')).toBe('sc');
    expect(classifyCjkFont('Microsoft YaHei')).toBe('sc');
    expect(classifyCjkFont('微软雅黑')).toBe('sc');
    expect(classifyCjkFont('DengXian')).toBe('sc');
    expect(classifyCjkFont('等线')).toBe('sc');
    expect(classifyCjkFont('FangSong')).toBe('sc');
    expect(classifyCjkFont('KaiTi')).toBe('sc');
    expect(classifyCjkFont('SimHei')).toBe('sc');
  });

  it('classifies Traditional Chinese faces (PMingLiU, 新細明體, Microsoft JhengHei, MingLiU, DFKai-SB)', () => {
    expect(classifyCjkFont('PMingLiU')).toBe('tc');
    expect(classifyCjkFont('新細明體')).toBe('tc');
    expect(classifyCjkFont('Microsoft JhengHei')).toBe('tc');
    expect(classifyCjkFont('微軟正黑體')).toBe('tc');
    expect(classifyCjkFont('MingLiU')).toBe('tc');
    expect(classifyCjkFont('DFKai-SB')).toBe('tc');
    expect(classifyCjkFont('標楷體')).toBe('tc');
  });

  it('classifies Japanese faces (Yu Gothic, Meiryo, MS Mincho, Hiragino) as jp', () => {
    expect(classifyCjkFont('Yu Gothic')).toBe('jp');
    expect(classifyCjkFont('游ゴシック')).toBe('jp');
    expect(classifyCjkFont('Meiryo')).toBe('jp');
    expect(classifyCjkFont('MS Mincho')).toBe('jp');
    expect(classifyCjkFont('ＭＳ 明朝')).toBe('jp');
    expect(classifyCjkFont('Hiragino Sans')).toBe('jp');
    expect(classifyCjkFont('ヒラギノ角ゴ')).toBe('jp');
  });

  it('returns null for non-CJK faces', () => {
    expect(classifyCjkFont('Arial')).toBeNull();
    expect(classifyCjkFont('Times New Roman')).toBeNull();
    expect(classifyCjkFont('Sakkal Majalla')).toBeNull();
    expect(classifyCjkFont('')).toBeNull();
  });

  it('is case-insensitive for Latin transliterations', () => {
    expect(classifyCjkFont('simsun')).toBe('sc');
    expect(classifyCjkFont('MALGUN GOTHIC')).toBe('kr');
  });
});

describe('classifyFontGeneric — font name → CSS generic class', () => {
  it('classifies Latin serif faces as serif', () => {
    expect(classifyFontGeneric('Century')).toBe('serif');
    expect(classifyFontGeneric('Garamond')).toBe('serif');
    expect(classifyFontGeneric('Times New Roman')).toBe('serif');
    expect(classifyFontGeneric('Cambria')).toBe('serif');
    expect(classifyFontGeneric('Georgia')).toBe('serif');
    expect(classifyFontGeneric('Palatino')).toBe('serif');
    expect(classifyFontGeneric('Didot')).toBe('serif');
    expect(classifyFontGeneric('Bodoni')).toBe('serif');
    expect(classifyFontGeneric('Playfair Display')).toBe('serif');
    expect(classifyFontGeneric('Source Serif Pro')).toBe('serif');
    expect(classifyFontGeneric('Noto Serif')).toBe('serif');
  });

  it('classifies Antiqua serif faces (Book Antiqua / Palatino clones) as serif', () => {
    // ECMA-376 documents authored in Central/Eastern Europe and Cyrillic locales
    // frequently use "*Antiqua" serif families (Book Antiqua — a Palatino clone —
    // plus locale variants like URW Antiqua, Antiqua). "Antiqua" is the German/
    // typographic term for a Roman/serif face, so any "*antiqua" name is serif and
    // must resolve to the serif Noto/Times fallback chain, not sans.
    expect(classifyFontGeneric('Book Antiqua')).toBe('serif');
    expect(classifyFontGeneric('Antiqua')).toBe('serif');
    expect(classifyFontGeneric('URW Antiqua')).toBe('serif');
    expect(classifyFontGeneric('BookAntiqua')).toBe('serif');
  });

  it('classifies CJK serif (song/ming/kai/fangsong) faces as serif', () => {
    expect(classifyFontGeneric('SimSun')).toBe('serif');
    expect(classifyFontGeneric('PMingLiU')).toBe('serif');
    expect(classifyFontGeneric('MS Mincho')).toBe('serif');
    expect(classifyFontGeneric('游明朝')).toBe('serif');
    expect(classifyFontGeneric('Batang')).toBe('serif');
    expect(classifyFontGeneric('KaiTi')).toBe('serif');
    expect(classifyFontGeneric('FangSong')).toBe('serif');
    expect(classifyFontGeneric('宋体')).toBe('serif');
  });

  it('classifies sans faces as sans', () => {
    expect(classifyFontGeneric('Calibri')).toBe('sans');
    expect(classifyFontGeneric('Arial')).toBe('sans');
    expect(classifyFontGeneric('Verdana')).toBe('sans');
    expect(classifyFontGeneric('Yu Gothic')).toBe('sans');
    expect(classifyFontGeneric('Meiryo')).toBe('sans');
    expect(classifyFontGeneric('Microsoft YaHei')).toBe('sans');
    expect(classifyFontGeneric('Malgun Gothic')).toBe('sans');
    expect(classifyFontGeneric('SimHei')).toBe('sans');
  });

  it('classifies monospace faces as mono', () => {
    expect(classifyFontGeneric('Consolas')).toBe('mono');
    expect(classifyFontGeneric('Courier New')).toBe('mono');
    expect(classifyFontGeneric('Cascadia Mono')).toBe('mono');
  });

  it('returns sans for nullish / empty input', () => {
    expect(classifyFontGeneric(null)).toBe('sans');
    expect(classifyFontGeneric(undefined)).toBe('sans');
    expect(classifyFontGeneric('')).toBe('sans');
  });

  it('regression guard: Century is serif (was misclassified as sans)', () => {
    expect(classifyFontGeneric('Century')).toBe('serif');
  });

  it('disambiguates substring collisions: Century Gothic / Source Sans are SANS', () => {
    // "Century Gothic" is a geometric SANS face that shares the "century" token
    // with the serif Century family — it must NOT be classified serif.
    expect(classifyFontGeneric('Century Gothic')).toBe('sans');
    expect(classifyFontGeneric('CenturyGothic')).toBe('sans');
    // …but the rest of the Century family stays serif.
    expect(classifyFontGeneric('Century Schoolbook')).toBe('serif');
    expect(classifyFontGeneric('New Century Schoolbook')).toBe('serif');
    // "Source Serif" leads with the explicit "source serif" token, so the sans
    // sibling does not leak into serif.
    expect(classifyFontGeneric('Source Sans Pro')).toBe('sans');
    expect(classifyFontGeneric('Source Serif Pro')).toBe('serif');
  });
});

describe('cjkFallbackChain — language-specific Noto CJK ordering', () => {
  it('places the matching Noto CJK first (sans)', () => {
    expect(cjkFallbackChain('kr', 'sans')[0]).toBe('Noto Sans KR');
    expect(cjkFallbackChain('sc', 'sans')[0]).toBe('Noto Sans SC');
    expect(cjkFallbackChain('tc', 'sans')[0]).toBe('Noto Sans TC');
    expect(cjkFallbackChain('jp', 'sans')[0]).toBe('Noto Sans JP');
    expect(cjkFallbackChain('hk', 'sans')[0]).toBe('Noto Sans HK');
  });

  it('does not add HK to non-HK fallback chains', () => {
    for (const lang of ['kr', 'sc', 'tc', 'jp'] as const) {
      expect(cjkFallbackChain(lang, 'sans')).not.toContain('Noto Sans HK');
      expect(cjkFallbackChain(lang, 'serif')).not.toContain('Noto Serif HK');
    }
  });

  it('places the matching Noto CJK first (serif)', () => {
    expect(cjkFallbackChain('kr', 'serif')[0]).toBe('Noto Serif KR');
    expect(cjkFallbackChain('sc', 'serif')[0]).toBe('Noto Serif SC');
    expect(cjkFallbackChain('tc', 'serif')[0]).toBe('Noto Serif TC');
    expect(cjkFallbackChain('jp', 'serif')[0]).toBe('Noto Serif JP');
  });

  it('does not invent or substitute a Google Fonts serif family for HK', () => {
    expect(cjkFallbackChain('hk', 'serif')).toEqual([]);
  });

  it('includes the other CJK languages after the matching one (so shared Han still resolves)', () => {
    const kr = cjkFallbackChain('kr', 'sans');
    expect(kr).toContain('Noto Sans JP');
    expect(kr).toContain('Noto Sans SC');
    expect(kr.indexOf('Noto Sans KR')).toBeLessThan(kr.indexOf('Noto Sans JP'));
  });
});

describe('non-CJK fallback constants', () => {
  it('sans chain covers Cyrillic (Noto Sans), Thai, Devanagari, Hebrew', () => {
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Thai');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Devanagari');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Hebrew');
  });

  it('serif chain covers Cyrillic (Noto Serif) and Hebrew', () => {
    expect(NON_CJK_SERIF_FALLBACKS).toContain('Noto Serif');
    expect(NON_CJK_SERIF_FALLBACKS).toContain('Noto Serif Hebrew');
  });
});

describe('SCRIPT_GOOGLE_FONTS / SCRIPT_PRELOAD_NAMES', () => {
  it('maps every native Noto CJK name to its Google Fonts family', () => {
    expect(GOOGLE_CJK_FONT_ALIASES).toEqual({
      'noto sans cjk sc': 'Noto Sans SC',
      'noto serif cjk sc': 'Noto Serif SC',
      'noto sans cjk tc': 'Noto Sans TC',
      'noto serif cjk tc': 'Noto Serif TC',
      'noto sans cjk jp': 'Noto Sans JP',
      'noto serif cjk jp': 'Noto Serif JP',
      'noto sans cjk kr': 'Noto Sans KR',
      'noto serif cjk kr': 'Noto Serif KR',
      'noto sans cjk hk': 'Noto Sans HK',
    });
    expect(googleCjkFontAlias('  NOTO SANS CJK SC ')).toBe('Noto Sans SC');
    expect(googleCjkFontAlias('Noto Serif CJK KR')).toBe('Noto Serif KR');
    expect(googleCjkFontAlias('Noto Sans CJK HK')).toBe('Noto Sans HK');
    expect(googleCjkFontAlias('Noto Serif CJK HK')).toBeNull();
    expect(googleCjkFontAlias('Noto Sans SC')).toBeNull();

    for (const [alias, family] of Object.entries(GOOGLE_CJK_FONT_ALIASES)) {
      expect(SCRIPT_GOOGLE_FONTS[alias]).toEqual({
        ...SCRIPT_GOOGLE_FONTS[family.toLowerCase()],
        loadFamily: family,
      });
    }
  });

  it('maps every script Noto family to a Google Fonts URL', () => {
    for (const name of SCRIPT_PRELOAD_NAMES) {
      const entry = SCRIPT_GOOGLE_FONTS[name.toLowerCase()];
      expect(entry, `missing Google Fonts entry for ${name}`).toBeDefined();
      expect(entry.url).toMatch(/fonts\.googleapis\.com/);
    }
  });

  it('includes CJK, Cyrillic-covering, Thai, Devanagari, Hebrew families', () => {
    expect(SCRIPT_GOOGLE_FONTS['noto sans kr']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans hk']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto serif hk']).toBeUndefined();
    expect(SCRIPT_GOOGLE_FONTS['noto serif sc']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans thai']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans devanagari']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans hebrew']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto serif hebrew']).toBeDefined();
  });

  it('preloads only the available HK sans family for shared Han text', () => {
    expect(scriptPreloadNamesForText(['香港'], 'hk')).toEqual(['Noto Sans HK']);
  });
});

describe('scriptPreloadNamesForText — script-aware preload set from document text', () => {
  it('returns nothing for empty / no text', () => {
    expect(scriptPreloadNamesForText([], null)).toEqual([]);
    expect(scriptPreloadNamesForText([''], null)).toEqual([]);
  });

  it('returns ZERO CJK/script names for pure Latin (ASCII + Latin-1)', () => {
    const out = scriptPreloadNamesForText(['Hello, World!', 'café résumé'], null);
    expect(out).toEqual([]);
  });

  it('detects Japanese (Hiragana/Katakana) → Noto Sans/Serif JP', () => {
    const out = scriptPreloadNamesForText(['こんにちは カタカナ'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
    assertResolvable(out);
  });

  it('detects Korean (Hangul) → Noto Sans/Serif KR', () => {
    const out = scriptPreloadNamesForText(['안녕하세요'], null);
    expect(out).toEqual(['Noto Sans KR', 'Noto Serif KR']);
    assertResolvable(out);
  });

  it('Han only with cjkLang hint sc → uses the hint', () => {
    const out = scriptPreloadNamesForText(['汉字测试'], 'sc');
    expect(out).toEqual(['Noto Sans SC', 'Noto Serif SC']);
  });

  it('Han only with no hint → defaults to jp', () => {
    const out = scriptPreloadNamesForText(['漢字'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('Han + Hangul → Hangul sets kr; Han resolves to the kr face (no JP emitted)', () => {
    const out = scriptPreloadNamesForText(['한국어 漢字'], 'jp');
    expect(out).toEqual(['Noto Sans KR', 'Noto Serif KR']);
  });

  it('Han + Kana → Kana sets jp even with sc hint', () => {
    const out = scriptPreloadNamesForText(['ひらがな 漢字'], 'sc');
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('Hangul + Kana → both KR and JP faces', () => {
    const out = scriptPreloadNamesForText(['한국 ひらがな'], null);
    expect(out).toContain('Noto Sans KR');
    expect(out).toContain('Noto Serif KR');
    expect(out).toContain('Noto Sans JP');
    expect(out).toContain('Noto Serif JP');
    assertResolvable(out);
  });

  it('detects Arabic → Noto Naskh Arabic + Noto Sans Arabic', () => {
    const out = scriptPreloadNamesForText(['مرحبا بالعالم'], null);
    expect(out).toContain('Noto Naskh Arabic');
    expect(out).toContain('Noto Sans Arabic');
  });

  it('detects Thai → Noto Sans Thai', () => {
    const out = scriptPreloadNamesForText(['สวัสดี'], null);
    expect(out).toEqual(['Noto Sans Thai']);
    assertResolvable(out);
  });

  it('detects Hebrew → Noto Sans/Serif Hebrew', () => {
    const out = scriptPreloadNamesForText(['שלום עולם'], null);
    expect(out).toContain('Noto Sans Hebrew');
    expect(out).toContain('Noto Serif Hebrew');
    assertResolvable(out);
  });

  it('detects Devanagari → Noto Sans Devanagari', () => {
    const out = scriptPreloadNamesForText(['नमस्ते'], null);
    expect(out).toEqual(['Noto Sans Devanagari']);
    assertResolvable(out);
  });

  it('detects Cyrillic → Noto Sans/Serif (which cover Cyrillic)', () => {
    const out = scriptPreloadNamesForText(['Привет мир'], null);
    expect(out).toEqual(['Noto Sans', 'Noto Serif']);
    assertResolvable(out);
  });

  it('detects Greek → Noto Sans/Serif', () => {
    const out = scriptPreloadNamesForText(['Γειά σου'], null);
    expect(out).toEqual(['Noto Sans', 'Noto Serif']);
  });

  it('mixed CJK (Japanese) + Arabic + Cyrillic', () => {
    const out = scriptPreloadNamesForText(['日本語 العربية Кириллица'], null);
    expect(out).toContain('Noto Sans JP');
    expect(out).toContain('Noto Serif JP');
    expect(out).toContain('Noto Naskh Arabic');
    expect(out).toContain('Noto Sans Arabic');
    expect(out).toContain('Noto Sans');
    expect(out).toContain('Noto Serif');
  });

  it('handles astral Han (U+20000+) via codePointAt', () => {
    const out = scriptPreloadNamesForText(['\u{20089}'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('is deterministic: same text yields the same set regardless of chunking', () => {
    const a = scriptPreloadNamesForText(['日本語 العربية'], null);
    const b = scriptPreloadNamesForText(['日本', '語 ', 'العربية'], null);
    expect([...a].sort()).toEqual([...b].sort());
  });

  it('every returned name is resolvable in SCRIPT_GOOGLE_FONTS (CJK + Cyrillic + scripts)', () => {
    const out = scriptPreloadNamesForText(['漢字 Привет สวัสดी नमस्ते שלום'], null);
    assertResolvable(out);
  });
});

describe('ScriptPreloadAccumulator — incremental script facts', () => {
  it('is exactly equivalent to one-shot collection across incremental feeds', () => {
    const chunks = ['漢字', '한국어', 'ひらがな', ' العربية', ' Привет'];
    const accumulator = new ScriptPreloadAccumulator('sc');
    for (const chunk of chunks) accumulator.addText([chunk]);
    expect(accumulator.names()).toEqual(scriptPreloadNamesForText(chunks, 'sc'));
  });

  it('preserves document-wide Han/Kana/Hangul interaction', () => {
    const accumulator = new ScriptPreloadAccumulator('sc');
    accumulator.addText(['漢']);
    expect(accumulator.names()).toEqual(['Noto Sans SC', 'Noto Serif SC']);
    accumulator.addText(['한']);
    expect(accumulator.names()).toEqual(['Noto Sans KR', 'Noto Serif KR']);
    accumulator.addText(['か']);
    expect(accumulator.names()).toEqual([
      'Noto Sans KR', 'Noto Serif KR', 'Noto Sans JP', 'Noto Serif JP',
    ]);
  });

  it('handles empty feeds and split chunks deterministically', () => {
    const accumulator = new ScriptPreloadAccumulator(null);
    accumulator.addText([]);
    accumulator.addText(['']);
    accumulator.addText(['\uD840']);
    accumulator.addText(['\uDC89']);
    // A surrogate pair split across chunks is not a code point in either chunk;
    // callers feed rendered strings, so chunk boundaries are semantically real.
    expect(accumulator.names()).toEqual([]);
    accumulator.addText(['\uD840\uDC89']);
    expect(accumulator.names()).toEqual(['Noto Sans JP', 'Noto Serif JP']);
    expect(accumulator.names()).toEqual(accumulator.names());
  });

  it('clones script state without coupling later incremental feeds', () => {
    const original = new ScriptPreloadAccumulator('sc');
    original.addText(['漢']);
    const candidate = original.clone();
    candidate.addText(['한']);
    expect(original.names()).toEqual(['Noto Sans SC', 'Noto Serif SC']);
    expect(candidate.names()).toEqual(['Noto Sans KR', 'Noto Serif KR']);
  });
});

describe('isComplexScriptCodePoint — §17.3.2.26 cs-axis blocks', () => {
  const cp = (s: string) => s.codePointAt(0) as number;
  it('classifies RTL/complex scripts as complex', () => {
    expect(isComplexScriptCodePoint(cp('א'))).toBe(true); // Hebrew U+05D0
    expect(isComplexScriptCodePoint(cp('ع'))).toBe(true); // Arabic U+0639
    expect(isComplexScriptCodePoint(cp('ܐ'))).toBe(true); // Syriac U+0710
    expect(isComplexScriptCodePoint(cp('ޱ'))).toBe(true); // Thaana U+07B1
    expect(isComplexScriptCodePoint(0xfdf2)).toBe(true); // Arabic Presentation Forms-A
    expect(isComplexScriptCodePoint(0xfe8e)).toBe(true); // Arabic Presentation Forms-B
    expect(isComplexScriptCodePoint(0x1ee00)).toBe(true); // Arabic Math (Plane-1)
  });
  it('classifies Latin / digits / punctuation / CJK as NOT complex', () => {
    expect(isComplexScriptCodePoint(cp('A'))).toBe(false);
    expect(isComplexScriptCodePoint(cp('5'))).toBe(false);
    expect(isComplexScriptCodePoint(cp('.'))).toBe(false);
    expect(isComplexScriptCodePoint(cp('あ'))).toBe(false); // Hiragana
    expect(isComplexScriptCodePoint(cp('漢'))).toBe(false); // Han
    expect(isComplexScriptCodePoint(cp('가'))).toBe(false); // Hangul
  });
});
