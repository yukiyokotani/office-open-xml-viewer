// Binary Word section page numbering (MS-DOC 2.6.4 sprmSNfcPgn,
// sprmSFPgnRestart, sprmSPgnStart97; MS-OSHARED 2.2.1.3 MSONFC).
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument, openDocxDocument, skia, skiaFactory } from './node-facade.js';

const format = (value: number) => concat(little16(0x300e), new Uint8Array([value]));
const restart = (value: number) => concat(little16(0x3011), new Uint8Array([1]), little16(0x7044), little32(value));

it('maps every defined MSONFC to its page number format', async () => {
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
    const model = await materializeDocxDocument(buildDocFixture({ text: 'Body\r', sectionProperties: concat(format(value), restart(0)) }),
      { modelSources: [testDocSource()] }) as unknown as { section: { pageNumType?: { fmt?: string; start?: number } } };
    expect(model.section.pageNumType, `MSONFC ${value}`).toEqual({ start: 0, fmt: value === 255 ? 'none' : formats[value] });
  }
});

const footer = 'Page \u0013PAGE\u001499\u0015 of \u0013NUMPAGES\u001499\u0015\r';

it.skipIf(!skia).each([
  { properties: [concat(format(2), restart(5)), new Uint8Array(), concat(format(3), restart(2))], expected: ['v', '6', 'B'] },
  { properties: [restart(0), concat(little16(0x501c), little16(99)), format(2)], expected: ['0', '1', 'ii'] },
  { properties: [restart(65536), new Uint8Array(), restart(2147483646)], expected: ['65536', '65537', '2147483646'] },
  { properties: [format(0xff), format(0x16), format(0x17)], expected: ['', '02', '3'] },
])('renders section-local page numbering: $expected', async ({ properties, expected }) => {
  const source = buildDocFixture({ text: 'One\fTwo\fThree\r', sectionEnds: [4, 8, 14],
    sectionProperties: properties, defaultTabTwips: 720,
    headers: ['', '', '', footer, '', '', ...Array<string>(12).fill('')],
  });
  const session = await openDocxDocument(source, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(3);
    for (let page = 0; page < 3; page++) {
      const runs: { text: string; y: number }[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      expect(runs.filter(r => r.y > 900).map(r => r.text).join('')).toBe(`Page ${expected[page]} of 3`);
    }
  } finally { await session.close(); }
});

it.skipIf(!skia)('propagates a bounded page-number expansion failure and permits a later open', async () => {
  const source = (value: number) => buildDocFixture({ text: 'Body\r', defaultTabTwips: 720,
    sectionProperties: concat(format(3), restart(value)),
    headers: ['', '', '', '\u0013PAGE\u001499\u0015\r', '', ''],
  });
  const render = async (value: number) => {
    const session = await openDocxDocument(source(value), { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
    try { await session.renderPage(0, { dpr: 1 }); }
    finally { await session.close(); }
  };
  await expect(render(2147483646)).rejects.toThrow(/number-format output budget/i);
  await expect(render(1)).resolves.toBeUndefined();
});
