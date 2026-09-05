import { describe, it, expect } from 'vitest';
import { formatOrdinalNumber, type NumberFormat } from './number-format';

describe('formatOrdinalNumber — ECMA-376 §17.18.59 ST_NumberFormat', () => {
  describe('decimal (Arabic cardinal)', () => {
    it('renders positive integers verbatim', () => {
      expect(formatOrdinalNumber(1, 'decimal')).toBe('1');
      expect(formatOrdinalNumber(123, 'decimal')).toBe('123');
      expect(formatOrdinalNumber(0, 'decimal')).toBe('0');
    });
    it('keeps the sign for negatives', () => {
      expect(formatOrdinalNumber(-5, 'decimal')).toBe('-5');
    });
  });

  describe('lowerRoman / upperRoman — §17.16.4.3.1 roman / Roman', () => {
    it('converts standard values (spec example: 123 -> cxxiii / CXXIII)', () => {
      expect(formatOrdinalNumber(123, 'lowerRoman')).toBe('cxxiii');
      expect(formatOrdinalNumber(123, 'upperRoman')).toBe('CXXIII');
    });
    it('handles the four subtractive pairs', () => {
      expect(formatOrdinalNumber(4, 'upperRoman')).toBe('IV');
      expect(formatOrdinalNumber(9, 'upperRoman')).toBe('IX');
      expect(formatOrdinalNumber(40, 'upperRoman')).toBe('XL');
      expect(formatOrdinalNumber(90, 'upperRoman')).toBe('XC');
      expect(formatOrdinalNumber(400, 'upperRoman')).toBe('CD');
      expect(formatOrdinalNumber(900, 'upperRoman')).toBe('CM');
    });
    it('renders 1 and 3999 (max within the classic additive system)', () => {
      expect(formatOrdinalNumber(1, 'lowerRoman')).toBe('i');
      expect(formatOrdinalNumber(3999, 'upperRoman')).toBe('MMMCMXCIX');
    });
    it('renders values above 3999 with repeated M (no bars/overlines)', () => {
      // 4000 = MMMM (Word writes four Ms; it does NOT use the vinculum bar form).
      expect(formatOrdinalNumber(4000, 'upperRoman')).toBe('MMMM');
      expect(formatOrdinalNumber(4999, 'upperRoman')).toBe('MMMMCMXCIX');
    });
    it('falls back to decimal for zero and negatives (roman has no glyph for 0)', () => {
      expect(formatOrdinalNumber(0, 'lowerRoman')).toBe('0');
      expect(formatOrdinalNumber(-1, 'upperRoman')).toBe('-1');
    });
  });

  describe('lowerLetter / upperLetter — §17.16.4.3.1 alphabetic / ALPHABETIC', () => {
    it('maps 1..26 to a..z / A..Z', () => {
      expect(formatOrdinalNumber(1, 'lowerLetter')).toBe('a');
      expect(formatOrdinalNumber(26, 'lowerLetter')).toBe('z');
      expect(formatOrdinalNumber(1, 'upperLetter')).toBe('A');
      expect(formatOrdinalNumber(26, 'upperLetter')).toBe('Z');
    });
    it('repeats the same letter beyond 26 (spec: 27 -> aa, 52 -> zz, 54 -> BBB)', () => {
      expect(formatOrdinalNumber(27, 'lowerLetter')).toBe('aa');
      expect(formatOrdinalNumber(28, 'lowerLetter')).toBe('bb');
      expect(formatOrdinalNumber(52, 'lowerLetter')).toBe('zz');
      expect(formatOrdinalNumber(53, 'lowerLetter')).toBe('aaa');
      expect(formatOrdinalNumber(54, 'upperLetter')).toBe('BBB');
    });
    it('falls back to decimal for zero and negatives (no letter for 0)', () => {
      expect(formatOrdinalNumber(0, 'upperLetter')).toBe('0');
      expect(formatOrdinalNumber(-3, 'lowerLetter')).toBe('-3');
    });
  });

  // Table-driven per-format checks for the international systems. Each row is a
  // format and a list of [input, expected] pairs derived from the §17.18.59
  // enumeration definitions (character sets + construction steps) and, where the
  // spec gives them, the §17.16.4.3.1 field-switch examples (noted inline).
  describe('international ST_NumberFormat systems', () => {
    const cases: Array<[NumberFormat, Array<[number, string]>]> = [
      // ── Positional digit substitution (base-10, own zero glyph) ──────────
      // decimalFullWidth ０–９ (U+FF10–U+FF19).
      ['decimalFullWidth', [
        [1, '１'], [9, '９'], [10, '１０'], [123, '１２３'], [2024, '２０２４'],
      ]],
      // decimalHalfWidth == ASCII decimal.
      ['decimalHalfWidth', [
        [1, '1'], [10, '10'], [123, '123'],
      ]],
      // thaiNumbers ๐–๙ (spec field example 123 -> ๑๒๓).
      ['thaiNumbers', [
        [1, '๑'], [9, '๙'], [10, '๑๐'], [20, '๒๐'], [123, '๑๒๓'],
      ]],
      // hindiNumbers १–९, zero ० (spec field example 123 -> १२३).
      ['hindiNumbers', [
        [1, '१'], [9, '९'], [10, '१०'], [123, '१२३'],
      ]],
      // ideographDigital 〇一二…九, positional (spec: 10 -> 一〇, 20 -> 二〇).
      ['ideographDigital', [
        [1, '一'], [9, '九'], [10, '一〇'], [11, '一一'], [20, '二〇'],
        [100, '一〇〇'], [2024, '二〇二四'],
      ]],
      // japaneseDigitalTenThousand is identical to ideographDigital.
      ['japaneseDigitalTenThousand', [
        [10, '一〇'], [100, '一〇〇'],
      ]],
      // koreanDigital 영일이…구, positional (zero = 영).
      ['koreanDigital', [
        [1, '일'], [9, '구'], [10, '일영'], [11, '일일'], [20, '이영'],
        [100, '일영영'],
      ]],
      // koreanDigital2 CJK numerals, zero = 零.
      ['koreanDigital2', [
        [1, '一'], [10, '一零'], [11, '一一'],
      ]],
      // taiwaneseDigital CJK numerals, zero = ○.
      ['taiwaneseDigital', [
        [1, '一'], [10, '一○'], [20, '二○'],
      ]],

      // ── chineseCounting / taiwaneseCounting (十-prefix positional) ────────
      // §17.18.59 example: 10 -> 十, 11 -> 十一, 20 -> 二十, 99 -> 九十九,
      // 100 -> 一〇〇, 101 -> 一〇一.
      ['chineseCounting', [
        [1, '一'], [9, '九'], [10, '十'], [11, '十一'], [15, '十五'],
        [19, '十九'], [20, '二十'], [21, '二十一'], [99, '九十九'],
        [100, '一〇〇'], [101, '一〇一'], [111, '一一一'], [200, '二〇〇'],
        [2024, '二〇二四'],
      ]],
      ['taiwaneseCounting', [
        [10, '十'], [20, '二十'], [99, '九十九'], [100, '一○○'],
      ]],

      // ── Grouped counting / legal CJK ────────────────────────────────────
      // japaneseCounting: leading-one elision (10 -> 十, 100 -> 百), NO 零 fill.
      // §17.18.59 example: 十, 十一, …, 二十, 二十一.
      ['japaneseCounting', [
        [1, '一'], [10, '十'], [11, '十一'], [19, '十九'], [20, '二十'],
        [21, '二十一'], [100, '百'], [111, '百十一'], [200, '二百'],
        [1000, '千'], [2024, '二千二十四'], [10000, '一万'], [10005, '一万五'],
        [12345, '一万二千三百四十五'],
      ]],
      // chineseCountingThousand: NO leading-one elision (10 -> 一十), WITH 零 fill.
      ['chineseCountingThousand', [
        [1, '一'], [10, '一十'], [11, '一十一'], [20, '二十'], [100, '一百'],
        [111, '一百一十一'], [1000, '一千'], [1001, '一千零一'],
        [2024, '二千零二十四'], [10000, '一万'], [10005, '一万零五'],
        [12345, '一万二千三百四十五'],
      ]],
      // taiwaneseCountingThousand mirrors chineseCountingThousand's rules.
      ['taiwaneseCountingThousand', [
        [10, '一十'], [1001, '一千零一'],
      ]],
      // chineseLegalSimplified: legal digits 壹贰叁…, 拾佰仟万, WITH 零 fill.
      // §17.16.4.3.1 CHINESENUM2 example 123 -> 壹佰貳拾參 (traditional); the
      // simplified glyphs give 壹佰贰拾叁.
      ['chineseLegalSimplified', [
        [1, '壹'], [10, '壹拾'], [20, '贰拾'], [100, '壹佰'],
        [123, '壹佰贰拾叁'], [1001, '壹仟零壹'],
      ]],
      // ideographLegalTraditional: traditional legal digits 壹貳參…, NO 零 fill.
      // §17.18.59 example: 壹, …, 壹拾, 壹拾壹, …, 貳拾, 貳拾壹.
      ['ideographLegalTraditional', [
        [1, '壹'], [10, '壹拾'], [11, '壹拾壹'], [20, '貳拾'], [21, '貳拾壹'],
        [123, '壹佰貳拾參'],
      ]],
      // japaneseLegal: 壱弐参…, thousand = 阡, myriad = 萬, NO 零 fill.
      // §17.18.59 example: 壱, 弐, 参, …, 壱拾, 壱拾壱, …, 弐拾, 弐拾壱.
      ['japaneseLegal', [
        [1, '壱'], [10, '壱拾'], [11, '壱拾壱'], [20, '弐拾'], [21, '弐拾壱'],
        [100, '壱百'], [2024, '弐阡弐拾四'], [10000, '壱萬'],
      ]],
      // koreanCounting: 일이삼…, 십백천만, leading-one elision, NO 零 fill.
      // §17.18.59 example: 일, 이, 삼, …, 팔, 구, 십, 십일.
      ['koreanCounting', [
        [1, '일'], [10, '십'], [11, '십일'], [20, '이십'], [100, '백'],
        [2024, '이천이십사'],
      ]],
      // koreanLegal: native-Korean words, tens-word + ones-word (§17.18.59
      // example: 하나, 열, 열하나, 스물, 스물하나). ≥100 undefined → decimal.
      ['koreanLegal', [
        [1, '하나'], [9, '아홉'], [10, '열'], [11, '열하나'], [20, '스물'],
        [21, '스물하나'], [90, '아흔'], [99, '아흔아홉'], [100, '100'],
      ]],

      // ── Repeat-letter alphabets ─────────────────────────────────────────
      // arabicAlpha (28): §17.16.4.3.1 ARABICALPHA example 12 -> س.
      ['arabicAlpha', [
        [1, 'أ'], [12, 'س'], [28, 'ي'], [29, 'أأ'], [56, 'يي'], [57, 'أأأ'],
      ]],
      // arabicAbjad (28): §17.16.4.3.1 ARABICABJAD example 12 -> ل.
      ['arabicAbjad', [
        [1, 'أ'], [12, 'ل'], [28, 'ظ'], [29, 'أأ'],
      ]],
      // hebrew2 (22-letter alphabet + ת suffix — NOT the repeat scheme).
      // §17.18.59 steps: write the RESULT glyph once, then append ת once per
      // 22 subtracted: 23 -> את, 24 -> בת. §17.16.4.3.1 field example:
      // 123 \* HEBREW2 -> מ + 5×ת (123 − 5×22 = 13 -> מ). 44 − 22 = 22 stops
      // ("equal to or less than the size of the set") -> glyph ת + 1×ת = תת.
      ['hebrew2', [
        [1, 'א'], [10, 'י'], [22, 'ת'], [23, 'את'], [24, 'בת'], [44, 'תת'],
        [45, 'אתת'], [123, 'מתתתתת'],
      ]],
      // russianLower / russianUpper (29-letter alphabet, repeat).
      ['russianLower', [
        [1, 'а'], [29, 'я'], [30, 'аа'], [58, 'яя'], [59, 'ааа'],
      ]],
      ['russianUpper', [
        [1, 'А'], [29, 'Я'], [30, 'АА'],
      ]],
      // thaiLetters (41-letter set, repeat).
      ['thaiLetters', [
        [1, 'ก'], [41, 'ฮ'], [42, 'กก'],
      ]],
      // chosung / ganada (14-jamo/syllable sets, repeat).
      ['chosung', [
        [1, 'ㄱ'], [14, 'ㅎ'], [15, 'ㄱㄱ'],
      ]],
      ['ganada', [
        [1, '가'], [14, '하'], [15, '가가'],
      ]],
      // hindiVowels (37) / hindiConsonants (18) sets, repeat.
      ['hindiVowels', [
        [1, 'क'], [37, 'ह'], [38, 'कक'],
      ]],
      ['hindiConsonants', [
        [1, 'अ'], [16, 'औ'], [17, 'अं'], [18, 'अः'], [19, 'अअ'],
      ]],

      // ── Hebrew gematria (positional letters) ────────────────────────────
      // §17.16.4.3.1 HEBREW1 example 123 -> קכג. 15/16 special cases (טו/טז).
      ['hebrew1', [
        [1, 'א'], [9, 'ט'], [10, 'י'], [15, 'טו'], [16, 'טז'], [17, 'יז'],
        [20, 'כ'], [21, 'כא'], [100, 'ק'], [123, 'קכג'], [200, 'ר'],
        [400, 'ת'], [500, 'ך'],
      ]],

      // ── Other algorithmic systems ───────────────────────────────────────
      // hex: base-16 over 0–9, A–F, UPPERCASE (§17.18.59 example: …, E, F, 10,
      // 11, …, 1E, 1F, 20). NOTE the §17.16.4.3.1 field example "355 → FF" is
      // the spec's own arithmetic error (0xFF = 255); the algorithm is normative.
      ['hex', [
        [1, '1'], [9, '9'], [10, 'A'], [15, 'F'], [16, '10'], [30, '1E'],
        [31, '1F'], [32, '20'], [255, 'FF'], [4096, '1000'],
      ]],
      // numberInDash: decimal between two dashes (§17.18.59 example: - 1 -,
      // - 2 -, …, - 10 -; §17.16.4.3.1 ArabicDash: 123 -> - 123 -).
      ['numberInDash', [
        [1, '- 1 -'], [9, '- 9 -'], [10, '- 10 -'], [123, '- 123 -'],
      ]],
      // decimalZero: 1–9 zero-padded, everything else plain decimal (§17.18.59
      // example: 01, 02, …, 09, 10, 11, …, 99, 100, 101).
      ['decimalZero', [
        [1, '01'], [9, '09'], [10, '10'], [99, '99'], [100, '100'],
      ]],

      // ── Enclosed / CJK sequence formats ─────────────────────────────────
      // decimalEnclosedCircle: §17.18.59 tables 1–20 → U+2460–U+2473 (①..⑳);
      // "for values greater than the size of the set, the items fall back to
      // the decimal format" (spec example: …, ⑲, ⑳, 21, …). Boundary 20/21.
      ['decimalEnclosedCircle', [
        [1, '①'], [2, '②'], [3, '③'], [10, '⑩'],
        [20, '⑳'], [21, '21'], [100, '100'],
      ]],
      // aiueoFullWidth: §17.18.59 full-width katakana in a-i-u-e-o order. The
      // spec's ENUMERATED 48-code-point list includes the archaic ヰ/ヱ (so wo/n
      // land at 47/48); repeat scheme past 48. Field example: 1 \* AIUEO -> ア.
      ['aiueoFullWidth', [
        [1, 'ア'], [2, 'イ'], [5, 'オ'], [16, 'タ'],
        [44, 'ワ'], [45, 'ヰ'], [46, 'ヱ'], [47, 'ヲ'],
        [48, 'ン'], [49, 'アア'], [96, 'ンン'],
        [97, 'アアア'],
      ]],
      // aiueo: §17.18.59 half-width katakana, positions 1–46 (U+FF71–U+FF9C,
      // U+FF66 ｦ, U+FF9D ﾝ — no archaic forms). Spec example: ｱ, ｲ, …, ｦ, ﾝ, ｱｱ.
      ['aiueo', [
        [1, 'ｱ'], [5, 'ｵ'], [44, 'ﾜ'], [45, 'ｦ'],
        [46, 'ﾝ'], [47, 'ｱｱ'], [92, 'ﾝﾝ'],
      ]],
    ];

    for (const [fmt, rows] of cases) {
      describe(fmt, () => {
        for (const [input, expected] of rows) {
          it(`${input} -> ${expected}`, () => {
            expect(formatOrdinalNumber(input, fmt)).toBe(expected);
          });
        }
        it('falls back to decimal for zero and negatives', () => {
          expect(formatOrdinalNumber(0, fmt)).toBe('0');
          expect(formatOrdinalNumber(-2, fmt)).toBe('-2');
        });
      });
    }
  });

  describe('unsupported / text formats fall back to decimal', () => {
    it('renders the language spell-out formats as decimal (documented residual)', () => {
      // cardinalText/ordinalText/ordinal/thaiCounting/vietnameseCounting/… are
      // language-dependent spell-outs with no algorithmic definition in
      // §17.18.59; they degrade to Arabic decimal so a page number is never blank.
      expect(formatOrdinalNumber(5, 'cardinalText' as NumberFormat)).toBe('5');
      expect(formatOrdinalNumber(5, 'ordinalText' as NumberFormat)).toBe('5');
      expect(formatOrdinalNumber(5, 'thaiCounting' as NumberFormat)).toBe('5');
      expect(formatOrdinalNumber(5, 'vietnameseCounting' as NumberFormat)).toBe('5');
      expect(formatOrdinalNumber(5, 'custom' as NumberFormat)).toBe('5');
      expect(formatOrdinalNumber(5, undefined)).toBe('5');
    });
  });

  it('suppresses numbering only for the explicit none format', () => {
    for (const value of [0, 1, 5, -1, 2147483646]) {
      expect(formatOrdinalNumber(value, 'none')).toBe('');
      expect(formatOrdinalNumber(value, undefined)).toBe(String(value));
    }
  });

  it.each(['upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'hebrew2',
    'arabicAlpha', 'russianUpper', 'thaiLetters', 'aiueo'])('bounds output expansion for %s', fmt => {
    // A valid 32-bit DOC section start must not expand to millions of glyphs.
    expect(() => formatOrdinalNumber(2147483646, fmt)).toThrow(/number-format output budget/i);
    for (const value of [Infinity, NaN, 1.5, Number.MAX_VALUE]) {
      expect(formatOrdinalNumber(value, fmt)).toBe(String(value));
    }
  });

  it('checks exact repeated-output boundaries before allocating', () => {
    expect(formatOrdinalNumber(4096 * 26, 'upperLetter')).toHaveLength(4096);
    expect(() => formatOrdinalNumber(4096 * 26 + 1, 'upperLetter')).toThrow(RangeError);
    expect(formatOrdinalNumber(4096 * 1000, 'upperRoman')).toHaveLength(4096);
    expect(() => formatOrdinalNumber(4096 * 1000 + 1, 'upperRoman')).toThrow(RangeError);
    expect(formatOrdinalNumber(4096 * 22, 'hebrew2')).toHaveLength(4096);
    expect(() => formatOrdinalNumber(4096 * 22 + 1, 'hebrew2')).toThrow(RangeError);
    expect(formatOrdinalNumber(2147483646, 'decimal')).toBe('2147483646');
  });
});
