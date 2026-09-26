// Binary Word list numbering (MS-DOC 2.9.147 LSTF/LVLF, 2.9.131 LFO) as the
// direct DOC reader projects it onto the DOCX model's paragraph markers.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument } from './node-facade.js';

const reference = (id: number, level = 0) => concat(little16(0x460b), little16(id), little16(0x260a), new Uint8Array([level]));
const distance = (code: number, value: number) => concat(little16(code), little16(value));
const toggle = (code: number, value: boolean) => concat(little16(code), new Uint8Array([Number(value)]));
const load = (bytes: Uint8Array) => materializeDocxDocument(bytes, { modelSources: [testDocSource()] });

interface Marker { text: string; format: string; fontFamily?: string; fontFacts?: { fontSize?: number; bold?: boolean } }
interface Paragraph { type: string; numbering: Marker | null; indentLeft: number; indentFirst: number; bidi?: boolean; runs: { text?: string; bold?: boolean }[] }
const paragraphs = (model: { body: unknown[] }) => model.body as Paragraph[];
const markers = (model: { body: unknown[] }) => paragraphs(model).map(p => p.numbering?.text ?? null);

function listData(options: {
  restart?: number;
  bullet?: string;
  multilevel?: boolean;
  papx?: Uint8Array;
  chpx?: Uint8Array;
  formats?: readonly number[];
  starts?: readonly number[];
  legalLevels?: readonly number[];
  listIds?: readonly number[];
  overrides?: readonly { listId?: number; level?: number; start?: number }[];
  restartLimits?: readonly (number | null)[];
} = {}) {
  const listIds = options.listIds ?? [42];
  const papx = options.papx ?? concat(distance(0x845e, 720), distance(0x8460, -360));
  const chpx = options.chpx ?? new Uint8Array();
  const text = options.bullet ? little16(options.bullet.charCodeAt(0)) : concat(little16(0), little16('.'.charCodeAt(0)));
  const listHeader = (id: number) => {
    const header = new Uint8Array(28); header.set(little32(id));
    for (let i = 0; i < 9; i++) header.set(little16(0xfff), 8 + i * 2);
    header[26] = options.multilevel ? 0 : 1;
    return header;
  };
  const level = (i: number) => {
    const header = new Uint8Array(28); header.set(little32(1));
    header[4] = options.bullet ? 0x17 : (options.formats?.[i] ?? (options.multilevel && i === 0 ? 1 : 0));
    if (!options.bullet) header[6] = 1;
    header.set(little32(options.starts?.[i] ?? 1));
    if (options.legalLevels?.includes(i)) header[5] |= 4;
    const restart = options.restartLimits?.[i];
    if (restart !== undefined && restart !== null) { header[5] |= 8; header[26] = restart; }
    if (i > 0) header[7] = 3;
    const value = i === 0 ? text : concat(little16(0), little16(46), little16(i), little16(46));
    header[24] = chpx.length; header[25] = papx.length;
    return concat(header, papx, chpx, little16(value.length / 2), value);
  };
  const levels = listIds.flatMap(() => options.multilevel ? Array.from({ length: 9 }, (_, i) => level(i)) : [level(0)]);
  const overrideOptions = options.overrides ?? [
    { listId: listIds[0] },
    { listId: listIds[0], start: options.restart },
  ];
  const headers = overrideOptions.map(item => {
    const header = new Uint8Array(16); header.set(little32(item.listId ?? listIds[0]));
    header[12] = item.start === undefined ? 0 : 1;
    return header;
  });
  const data = overrideOptions.map((item, i) => concat(
    little32(i === 0 ? 0 : 0xffffffff),
    ...(item.start === undefined ? [] : [little32(item.start), new Uint8Array([0x10 | (item.level ?? 0), 0, 0, 0])]),
  ));
  return { definitionHeaderBytes: 2 + listIds.length * 28,
    definitions: concat(little16(listIds.length), ...listIds.map(listHeader), ...levels),
    overrides: concat(little32(headers.length), ...headers, ...data) };
}
const levelRuns = (levels: readonly number[], id = 1) => levels.map((level, i) => ({ end: (i + 1) * 2, properties: reference(id, level) }));

it.each([undefined, 7])('shares list counters across LFO aliases and changing marker formatting, restart=%s', async restart => {
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData({ restart }),
    formattingRuns: [1, 2, 1, 2].map((id, i) => ({ end: (i + 1) * 2, properties: concat(reference(id), distance(0x4a43, i % 2 ? 28 : 20)) })),
  }));
  expect(markers(model)).toEqual(restart === undefined ? ['1.', '2.', '3.', '4.'] : ['1.', '7.', '8.', '9.']);
  // The marker follows its paragraph mark's size (sprmCHps 20/28 half-points).
  expect(paragraphs(model).map(p => p.numbering?.fontFacts?.fontSize)).toEqual([10, 14, 10, 14]);
});

it('keeps interleaved list definitions with distinct LSIDs independent', async () => {
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData({
    listIds: [42, 84], overrides: [{ listId: 42 }, { listId: 84 }],
  }), formattingRuns: [1, 2, 1, 2].map((id, i) => ({ end: (i + 1) * 2, properties: reference(id) })) }));
  expect(markers(model)).toEqual(['1.', '1.', '2.', '2.']);
});

// The start-once fact itself is Rust-covered
// (numbering::direct::lfo_aliases_share_lsid_sequence_but_each_override_start_applies_once).
it('direct rejects a start-overridden LFO whose ancestor level only another LFO used', async () => {
  await expect(load(buildDocFixture({ text: 'A\rB\rC\r', numbering: listData({
    multilevel: true, formats: [1, 0], overrides: [
      { listId: 42, level: 1, start: 7 }, { listId: 42, level: 1, start: 11 },
    ],
  }), formattingRuns: [[1, 0], [1, 1], [2, 1]].map(([id, level], i) => ({ end: (i + 1) * 2, properties: reference(id, level) })) }))).rejects.toThrow(new Error('UNSUPPORTED:missing prior Word list ancestor level'));
});

it.each([
  { limit: 0, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.3.', 'II.', 'II.4.'] },
  { limit: 1, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.3.', 'II.', 'II.1.'] },
  { limit: 2, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.1.', 'II.', 'II.1.'] },
])('honors binary level-2 restart boundary $limit', async ({ limit, expected }) => {
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\rE\rF\rG\rH\r',
    numbering: listData({ multilevel: true, formats: [1, 0, 0], restartLimits: [null, null, limit] }),
    formattingRuns: levelRuns([0, 1, 2, 2, 1, 2, 0, 2]),
  }));
  expect(markers(model)).toEqual(expected);
});

it.each([
  { name: 'normal multilevel restarts keep ancestor formats', formats: [1, 0], legal: [], expected: ['I.', 'I.1.', 'I.2.', 'II.', 'II.1.'] },
  // MS-DOC 2.9.150 fLegal / 2.4.6.3: ancestors display Arabic, but ArabicLZ
  // (0x16) is preserved; ECMA-376 w:isLgl alone would force decimal.
  { name: 'legal Roman parent', formats: [1, 0], legal: [1], expected: ['I.', '1.1.', '1.2.', 'II.', '2.1.'] },
  { name: 'legal decimalZero parent', formats: [22, 0], legal: [1], expected: ['01.', '01.1.', '01.2.', '02.', '02.1.'] },
  { name: 'legal decimalZero child', formats: [1, 22], legal: [1], expected: ['I.', '1.01.', '1.02.', 'II.', '2.01.'] },
  { name: 'legal decimalZero both', formats: [22, 22], legal: [1], expected: ['01.', '01.01.', '01.02.', '02.', '02.01.'] },
  { name: 'nonlegal child', formats: [1, 4], legal: [], expected: ['I.', 'I.a.', 'I.b.', 'II.', 'II.a.'] },
  { name: 'legal decimalZero keeps width and authored starts across 9 to 10', formats: [22, 22], legal: [1], starts: [9, 9], expected: ['09.', '09.09.', '09.10.', '10.', '10.09.'] },
])('formats multilevel markers: $name', async ({ formats, legal, starts, expected }) => {
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\rE\r',
    numbering: listData({ multilevel: true, formats, starts, legalLevels: legal }),
    formattingRuns: levelRuns([0, 1, 1, 0, 1]),
  }));
  expect(markers(model)).toEqual(expected);
});

// MS-DOC leaves a legal level whose own format is not Arabic ambiguous; the
// converter showed Arabic ('1.1.' and a legal Roman parent as '1.').
it.each([
  { formats: [1, 4], legal: [1] },
  { formats: [1, 4], legal: [0] },
])('direct rejects a legal level with a non-Arabic current format: $formats legal=$legal', async ({ formats, legal }) => {
  await expect(load(buildDocFixture({ text: 'A\rB\r',
    numbering: listData({ multilevel: true, formats, legalLevels: legal }), formattingRuns: levelRuns([0, 1]) }))).rejects.toThrow(new Error('UNSUPPORTED:legal Word list level with a non-Arabic current format'));
});

it('lets direct paragraph indentation override the list level', async () => {
  const direct = (left: number, first: number) => concat(distance(0x845e, left), distance(0x8460, first));
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData(),
    formattingRuns: [
      concat(reference(1), direct(1440, 120)),
      concat(reference(1), direct(0, 0)),
      reference(1),
      concat(reference(-1), direct(1440, 120)),
    ].map((properties, i) => ({ end: (i + 1) * 2, properties })),
  }));
  expect(paragraphs(model).map(p => [p.indentLeft, p.indentFirst])).toEqual([[72, 6], [0, 0], [36, -18], [72, 6]]);
});

it('keeps direct physical indentation when direct bidi opposes the list level', async () => {
  const level = concat(distance(0x845e, 720), distance(0x8460, -360), toggle(0x2441, true));
  const properties = concat(reference(1), distance(0x845e, 1440), distance(0x8460, 120), toggle(0x2441, false));
  const [paragraph] = paragraphs(await load(buildDocFixture({ text: 'A\r', numbering: listData({ papx: level }), paragraphProperties: properties })));
  expect(paragraph.bidi ?? false).toBe(false);
  expect([paragraph.indentLeft, paragraph.indentFirst]).toEqual([72, 6]);
});

it('omits skipped/removed references without consuming their sequence values', async () => {
  const refs = [reference(1), reference(1, 12), reference(0, 255), reference(-2047, 255), reference(1)];
  const model = await load(buildDocFixture({ text: 'A\rB\rC\rD\rE\r', numbering: listData(),
    formattingRuns: refs.map((properties, i) => ({ end: (i + 1) * 2, properties })),
  }));
  expect(markers(model)).toEqual(['1.', null, null, null, '2.']);
});

it('keeps authored bullet text and applies marker-only CHPX to the marker', async () => {
  const [paragraph] = paragraphs(await load(buildDocFixture({ text: 'Body\r',
    numbering: listData({ bullet: '·', chpx: concat(little16(0x0835), new Uint8Array([1]), distance(0x4a43, 32)) }),
    paragraphProperties: reference(1) })));
  expect(paragraph.numbering).toMatchObject({ format: 'bullet', text: '·', fontFacts: { bold: true, fontSize: 16 } });
  expect(paragraph.runs.map(r => [r.text, r.bold])).toEqual([['Body', false]]);
});

it.each([
  { bullet: '', font: 'Symbol', charset: 2 },
  { bullet: '', font: 'Wingdings', charset: 0 },
  { bullet: '•', font: 'Arial', charset: 238 },
  { bullet: '・', font: 'ＭＳ 明朝', charset: 128 },
  // FFN names are complete UTF-16 (astral and XML-sensitive characters).
  { bullet: '', font: 'A😀&"', charset: 2 },
])('keeps bullet $bullet with its resolved $font marker font unmapped', async ({ bullet, font, charset }) => {
  const [paragraph] = paragraphs(await load(buildDocFixture({
    text: 'Body\r',
    fonts: [{ name: font, charset }],
    numbering: listData({ bullet, chpx: concat(distance(0x4a4f, 0), distance(0x4a51, 0)) }),
    paragraphProperties: reference(1),
  })));
  expect(paragraph.numbering).toMatchObject({ text: bullet, fontFamily: font });
});

it('direct rejects an XML-control bullet instead of dropping the marker', async () => {
  await expect(load(buildDocFixture({ text: 'Body\r', numbering: listData({ bullet: '\u0001' }), paragraphProperties: reference(1) }))).rejects.toThrow(new Error('UNSUPPORTED:invalid Word numbering text Unicode'));
});

it('scopes numbering to each note/header without restarting the main story', async () => {
  const model = await load(buildDocFixture({ text: 'A\u0002\rB\u0002\r', numbering: listData(),
    paragraphProperties: reference(1), characterProperties: new Uint8Array([0x55, 0x08, 1]),
    headers: ['', 'Header\r', '', '', '', ''],
    footnotes: [{ cp: 1, text: '\u0002First note\r' }, { cp: 4, text: '\u0002Second note\r' }],
  })) as unknown as { body: Paragraph[]; footnotes: { content: Paragraph[] }[]; headers: { default: { body: Paragraph[] } } };
  expect(markers(model)).toEqual(['1.', '2.']);
  expect(model.footnotes.map(n => n.content[0].numbering?.text)).toEqual(['1.', '1.']);
  expect(model.headers.default.body[0].numbering?.text).toBe('1.');
});
