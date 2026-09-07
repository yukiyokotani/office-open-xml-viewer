import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16, little32 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const reference = (id: number, level = 0) => concat(little16(0x460b), little16(id), little16(0x260a), new Uint8Array([level]));
const distance = (code: number, value: number) => concat(little16(code), little16(value));
const toggle = (code: number, value: boolean) => concat(little16(code), new Uint8Array([Number(value)]));

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

async function convert(bytes: Uint8Array) {
  const result = await converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 4 * 1024 * 1024 });
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  const part = (name: string) => new TextDecoder().decode(archive.extract_image(name));
  try {
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    return { result, model, document: part('word/document.xml'), numbering: part('word/numbering.xml'),
      relationships: part('word/_rels/document.xml.rels'), types: part('[Content_Types].xml') };
  } finally { archive.free(); }
}

async function convertWithOptionalNumbering(bytes: Uint8Array) {
  const result = await converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 4 * 1024 * 1024 });
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    const document = new TextDecoder().decode(archive.extract_image('word/document.xml'));
    let numbering: string | undefined;
    try {
      numbering = new TextDecoder().decode(archive.extract_image('word/numbering.xml'));
    } catch {
      // A document whose only list marker is unsupported has no numbering part.
    }
    let relationships: string | undefined;
    try {
      relationships = new TextDecoder().decode(archive.extract_image('word/_rels/document.xml.rels'));
    } catch {
      // A package with no document relationships omits the relationships part.
    }
    const types = new TextDecoder().decode(archive.extract_image('[Content_Types].xml'));
    return { result, model, document, numbering, relationships, types };
  } finally { archive.free(); }
}

it.each([undefined, 7])('shares list counters across aliases and changing marker formatting, restart=%s', async restart => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData({ restart }),
    formattingRuns: [1, 2, 1, 2].map((id, i) => ({ end: (i + 1) * 2, properties: concat(reference(id), distance(0x4a43, i % 2 ? 28 : 20)) })),
  }));
  expect(output.relationships).toContain('relationships/numbering');
  expect(output.types).toContain('wordprocessingml.numbering+xml');
  expect(output.numbering.match(/<w:abstractNum /g)).toHaveLength(1);
  const numbered = output.model.body.filter((p: { numbering?: unknown }) => p.numbering);
  expect(numbered.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(restart === undefined ? ['1.', '2.', '3.', '4.'] : ['1.', '7.', '8.', '9.']);
  expect(output.numbering).toContain('<w:sz w:val="28"/>');
  expect(output.numbering).toContain('<w:sz w:val="20"/>');
});

it('keeps interleaved list definitions with distinct LSIDs independent', async () => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData({
    listIds: [42, 84], overrides: [{ listId: 42 }, { listId: 84 }],
  }), formattingRuns: [1, 2, 1, 2].map((id, i) => ({ end: (i + 1) * 2, properties: reference(id) })) }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(['1.', '1.', '2.', '2.']);
  expect(output.numbering.match(/<w:abstractNum /g)).toHaveLength(2);
});

it('applies a level start override once per active LFO across marker-format variants', async () => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r', numbering: listData({
    multilevel: true, formats: [1, 0], overrides: [
      { listId: 42, level: 1, start: 7 }, { listId: 42, level: 1, start: 11 },
    ],
  }), formattingRuns: [
    { id: 1, level: 0, size: 20 },
    { id: 1, level: 1, size: 20 },
    { id: 2, level: 1, size: 20 },
    { id: 1, level: 1, size: 28 },
    { id: 2, level: 1, size: 28 },
  ].map(({ id, level, size }, i) => ({ end: (i + 1) * 2, properties: concat(reference(id, level), distance(0x4a43, size)) })) }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual([
    'I.', 'I.7.', 'I.11.', 'I.12.', 'I.13.',
  ]);
  const numIds = [...output.document.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map(match => match[1]);
  expect(numIds).toHaveLength(5);
  expect(numIds.slice(1)).not.toContain(numIds[0]);
  expect(output.numbering.match(/<w:abstractNum /g)).toHaveLength(1);
});

it.each([
  { limit: 0, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.3.', 'II.', 'II.4.'] },
  { limit: 1, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.3.', 'II.', 'II.1.'] },
  { limit: 2, expected: ['I.', 'I.1.', 'I.1.', 'I.2.', 'I.2.', 'I.1.', 'II.', 'II.1.'] },
])('honors binary level-2 restart boundary $limit in emitted numbering', async ({ limit, expected }) => {
  const levels = [0, 1, 2, 2, 1, 2, 0, 2];
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\rF\rG\rH\r',
    numbering: listData({ multilevel: true, formats: [1, 0, 0], restartLimits: [null, null, limit] }),
    formattingRuns: levels.map((level, i) => ({ end: (i + 1) * 2, properties: reference(1, level) })),
  }));
  expect(output.numbering).toContain(`<w:lvlRestart w:val="${limit}"/>`);
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(expected);
});

it('preserves direct paragraph indentation after list formatting', async () => {
  const direct = (left: number, first: number) => concat(distance(0x845e, left), distance(0x8460, first));
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\r', numbering: listData(),
    formattingRuns: [
      concat(reference(1), direct(1440, 120)),
      concat(reference(1), direct(0, 0)),
      reference(1),
      concat(reference(-1), direct(1440, 120)),
    ].map((properties, i) => ({ end: (i + 1) * 2, properties })),
  }));
  const indents = [...output.document.matchAll(/<w:ind w:left="(\d+)" w:right="\d+" w:(hanging|firstLine)="(\d+)"\/>/g)].map(m => m.slice(1));
  expect(indents).toEqual([
    ['1440', 'firstLine', '120'],
    ['0', 'firstLine', '0'],
    ['720', 'hanging', '360'],
    ['1440', 'firstLine', '120'],
  ]);
});

it('preserves direct physical indentation when direct bidi opposes the list level', async () => {
  const level = concat(distance(0x845e, 720), distance(0x8460, -360), toggle(0x2441, true));
  const properties = concat(reference(1), distance(0x845e, 1440), distance(0x8460, 120), toggle(0x2441, false));
  const output = await convert(buildDocFixture({ text: 'A\r', numbering: listData({ papx: level }), paragraphProperties: properties }));
  expect(output.document).toContain('<w:bidi w:val="0"/>');
  expect(output.document).toContain('<w:ind w:left="1440" w:right="0" w:firstLine="120"/>');
});

it('omits skipped/removed references without consuming their sequence values', async () => {
  const refs = [reference(1), reference(1, 12), reference(0, 255), reference(-2047, 255), reference(1)];
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r', numbering: listData(),
    formattingRuns: refs.map((properties, i) => ({ end: (i + 1) * 2, properties })),
  }));
  expect(output.model.body.map((p: { numbering?: { text: string } }) => p.numbering?.text ?? null)).toEqual(['1.', null, null, null, '2.']);
});

it('preserves authored bullet text and applies marker-only CHPX', async () => {
  const output = await convert(buildDocFixture({ text: 'Body\r', numbering: listData({ bullet: '·', chpx: concat(little16(0x0835), new Uint8Array([1]), distance(0x4a43, 32)) }), paragraphProperties: reference(1) }));
  expect(output.numbering).toContain('w:numFmt w:val="bullet"');
  expect(output.numbering).toContain('w:lvlText w:val="·"');
  expect(output.numbering).toContain('<w:b w:val="1"/>');
  expect(output.document).not.toContain('<w:b w:val="1"/>');
  expect(output.model.body[0].numbering.text).toBe('·');
});

it.each([
  { bullet: '\uf0b7', font: 'Symbol', charset: 2 },
  { bullet: '\uf0fc', font: 'Wingdings', charset: 0 },
  { bullet: '\uf06c', font: 'Wingdings', charset: 0 },
  { bullet: '•', font: 'Arial', charset: 238 },
  { bullet: '•', font: 'Symbol', charset: 2 },
  { bullet: '・', font: 'ＭＳ 明朝', charset: 128 },
])('preserves $bullet with resolved $font references without mapping it', async ({ bullet, font, charset }) => {
  const output = await convert(buildDocFixture({
    text: 'Body\r',
    fonts: [{ name: font, charset }],
    numbering: listData({ bullet, chpx: concat(distance(0x4a4f, 0), distance(0x4a51, 0)) }),
    paragraphProperties: reference(1),
  }));
  expect(output.result.warnings).not.toContain('legacy-doc:unsupported-numbering-text-or-autonum-omitted');
  expect(output.numbering).toContain(`w:lvlText w:val="${bullet}"`);
  expect(output.numbering).toContain(`w:ascii="${font}"`);
  expect(output.numbering).toContain(`w:hAnsi="${font}"`);
  expect(output.model.body[0].numbering.text).toBe(bullet);
});

it('omits an XML-control bullet with a warning while preserving body text', async () => {
  const output = await convertWithOptionalNumbering(buildDocFixture({ text: 'Body\r',
    numbering: listData({ bullet: '\u0001' }), paragraphProperties: reference(1) }));
  expect(output.result.warnings).toContain('legacy-doc:unsupported-numbering-text-or-autonum-omitted');
  expect(output.numbering).toBeUndefined();
  expect(output.relationships ?? '').not.toContain('relationships/numbering');
  expect(output.types).not.toContain('wordprocessingml.numbering+xml');
  expect(output.document).toContain('>Body</w:t>');
  expect(output.model.body[0].numbering).toBeNull();
});

it('encodes astral and XML-sensitive synthetic FFN names as complete UTF-16 metadata', async () => {
  const font = 'A😀&"';
  const output = await convert(buildDocFixture({ text: 'Body\r', fonts: [{ name: font, charset: 2 }],
    numbering: listData({ bullet: '\uf0b7', chpx: concat(distance(0x4a4f, 0), distance(0x4a51, 0)) }), paragraphProperties: reference(1) }));
  expect(output.numbering).toContain('w:ascii="A😀&amp;&quot;"');
  expect(output.numbering).toContain('w:hAnsi="A😀&amp;&quot;"');
});

it.each([
  { name: '', charset: 0 },
  { name: 'bad\u0000name', charset: 0 },
  { name: 'Font', charset: -1 },
  { name: 'Font', charset: 256 },
  { name: 'Font', charset: 1.5 },
  { name: 'x'.repeat(108), charset: 0 },
])('rejects invalid synthetic FFN metadata: %#', font => {
  expect(() => buildDocFixture({ fonts: [font] })).toThrow(/synthetic font/i);
});

it('retains ancestor number formats and normal multilevel restarts', async () => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r', numbering: listData({ multilevel: true }),
    formattingRuns: [0, 1, 1, 0, 1].map((level, i) => ({ end: (i + 1) * 2, properties: reference(1, level) })),
  }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(['I.', 'I.1.', 'I.2.', 'II.', 'II.1.']);
  expect(output.numbering).toContain('<w:lvlRestart w:val="1"/>');
});

it.each([
  { formats: [1, 4], legal: true, expected: ['I.', '1.1.', '1.2.', 'II.', '2.1.'] },
  { formats: [22, 4], legal: true, expected: ['01.', '01.1.', '01.2.', '02.', '02.1.'] },
  { formats: [1, 22], legal: true, expected: ['I.', '1.01.', '1.02.', 'II.', '2.01.'] },
  { formats: [22, 22], legal: true, expected: ['01.', '01.01.', '01.02.', '02.', '02.01.'] },
  { formats: [1, 4], legal: false, expected: ['I.', 'I.a.', 'I.b.', 'II.', 'II.a.'] },
])('preserves binary legal-numbering formats: $formats, legal=$legal', async ({ formats, legal, expected }) => {
  // MS-DOC 2.9.150 fLegal / 2.4.6.3: ArabicLZ (0x16) is preserved.
  // Passing w:isLgl through is insufficient: ECMA-376 17.9.4 makes it decimal.
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r',
    numbering: listData({ multilevel: true, formats, legalLevels: legal ? [1] : [] }),
    formattingRuns: [0, 1, 1, 0, 1].map((level, i) => ({ end: (i + 1) * 2, properties: reference(1, level) })),
  }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(expected);
  expect(output.numbering).not.toContain('<w:isLgl');
  expect(output.numbering.match(/<w:abstractNum /g)).toHaveLength(1);
});

it('scopes numbering to each note/header without restarting the main story', async () => {
  const output = await convert(buildDocFixture({ text: 'A\u0002\rB\u0002\r', numbering: listData(),
    paragraphProperties: reference(1), characterProperties: new Uint8Array([0x55, 0x08, 1]),
    headers: ['', 'Header\r', '', '', '', ''],
    footnotes: [{ cp: 1, text: '\u0002First note\r' }, { cp: 4, text: '\u0002Second note\r' }],
  }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(['1.', '2.']);
  expect(output.model.footnotes.map((n: { content: { numbering: { text: string } }[] }) => n.content[0].numbering.text)).toEqual(['1.', '1.']);
  expect(output.model.headers.default.body[0].numbering.text).toBe('1.');
  expect(output.numbering.match(/<w:abstractNum /g)).toHaveLength(4);
});

it('does not inherit a parent paragraph\'s legal display into a nonlegal child', async () => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r',
    numbering: listData({ multilevel: true, formats: [1, 4], legalLevels: [0] }),
    formattingRuns: [0, 1, 1, 0, 1].map((level, i) => ({ end: (i + 1) * 2, properties: reference(1, level) })),
  }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(['1.', 'I.a.', 'I.b.', '2.', 'II.a.']);
});

it('keeps legal decimalZero width and authored restart values across 9 to 10', async () => {
  const output = await convert(buildDocFixture({ text: 'A\rB\rC\rD\rE\r',
    numbering: listData({ multilevel: true, formats: [22, 22], starts: [9, 9], legalLevels: [1] }),
    formattingRuns: [0, 1, 1, 0, 1].map((level, i) => ({ end: (i + 1) * 2, properties: reference(1, level) })),
  }));
  expect(output.model.body.map((p: { numbering: { text: string } }) => p.numbering.text)).toEqual(['09.', '09.09.', '09.10.', '10.', '10.09.']);
});
