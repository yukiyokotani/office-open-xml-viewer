import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

it('maps every defined MSONFC to its OOXML number format', async () => {
  // MS-OSHARED 2.2.1.3 maps 0x00..0x3b and 0xff. The binary bullet
  // page-number format uses the decimal fallback allowed by MS-DOC 2.6.4.
  const formats = (`decimal upperRoman lowerRoman upperLetter lowerLetter ordinal cardinalText ordinalText hex chicago
    ideographDigital japaneseCounting aiueo iroha decimalFullWidth decimalHalfWidth japaneseLegal japaneseDigitalTenThousand
    decimalEnclosedCircle decimalFullWidth2 aiueoFullWidth irohaFullWidth decimalZero decimal ganada chosung
    decimalEnclosedFullstop decimalEnclosedParen decimalEnclosedCircleChinese ideographEnclosedCircle ideographTraditional
    ideographZodiac ideographZodiacTraditional taiwaneseCounting ideographLegalTraditional taiwaneseCountingThousand taiwaneseDigital
    chineseCounting chineseLegalSimplified chineseCountingThousand decimal koreanDigital koreanCounting koreanLegal koreanDigital2
    hebrew1 arabicAlpha hebrew2 arabicAbjad hindiVowels hindiConsonants hindiNumbers hindiCounting thaiLetters thaiNumbers thaiCounting
    vietnameseCounting numberInDash russianLower russianUpper`).split(/\s+/);
  expect(formats).toHaveLength(60);
  for (const value of [...Array.from({ length: 60 }, (_, i) => i), 255]) {
    const result = await converter.convert({ bytes: buildDocFixture({ text: 'Body\r',
      sectionProperties: concat(little16(0x300e), new Uint8Array([value]), little16(0x3011), new Uint8Array([1])),
    }), from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
    const archive = new DocxArchive(new Uint8Array(result.bytes));
    try {
      const xml = new TextDecoder().decode(archive.extract_image('word/document.xml'));
      const format = xml.match(/<w:pgNumType w:fmt="([^"]+)" w:start="0"\/>/)?.[1];
      expect(format, `MSONFC ${value}`).toBe(value === 255 ? 'none' : formats[value]);
    } finally { archive.free(); }
  }
});
