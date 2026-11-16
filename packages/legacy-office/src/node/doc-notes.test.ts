// Binary Word footnotes and endnotes (MS-DOC 2.3.2/2.3.5 note documents,
// 2.8.16-2.8.20 reference/text PLCs) through the direct DOC reader.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument, openDocxDocument, skia, skiaFactory } from './node-facade.js';

const load = (bytes: Uint8Array) => materializeDocxDocument(bytes, { legacyDoc: testDocSource() });
interface Run { type: string; text?: string; bold?: boolean; imagePath?: string; noteRef?: { kind: string; id: string } }
interface Block { type: string; runs: Run[] }
interface Note { id: string; content: Block[] }
interface Model { body: Block[]; footnotes: Note[]; endnotes: Note[]; headers: { default: { body: Block[] } | null } }

it('projects formatted footnotes and endnotes with independent IDs, excluding comments and headers', async () => {
  const model = await load(buildDocFixture({ text: 'A\u0002B\u0002\r',
    footnotes: [{ cp: 1, text: '\u0002Footnote 😀\rSecond paragraph\r' }],
    endnotes: [{ cp: 3, text: '\u0002Endnote\r' }], comments: 'IGNORED COMMENT\r',
    headers: ['', 'Header\r', '', '', '', ''],
    characterProperties: new Uint8Array([0x55, 0x08, 1, 0x35, 0x08, 1]),
  })) as unknown as Model;
  expect(model.body[0].runs.map(r => r.noteRef ?? r.text)).toEqual(['A', { kind: 'footnote', id: '1' }, 'B', { kind: 'endnote', id: '1' }]);
  expect(model.footnotes).toHaveLength(1);
  expect(model.endnotes).toHaveLength(1);
  const [footnote] = model.footnotes;
  expect(footnote.id).toBe('1');
  // The in-note automatic mark is the note-reference run with an empty id.
  expect(footnote.content.map(p => p.runs.map(r => r.noteRef ?? r.text))).toEqual([
    [{ kind: 'footnote', id: '' }, 'Footnote 😀'], ['Second paragraph'],
  ]);
  expect(footnote.content.flatMap(p => p.runs).every(r => r.bold)).toBe(true);
  expect(model.endnotes[0].content.map(p => p.runs.map(r => r.noteRef ?? r.text))).toEqual([[{ kind: 'endnote', id: '' }, 'Endnote']]);
  expect(JSON.stringify(model)).not.toContain('IGNORED COMMENT');
});

// MS-DOC 2.3.3/2.8.22: six reserved header stories, each nonempty
// story ending in its own guard. U+0003's short kind is independent of role.
const shortFootnoteSeparators = ['\u0003\r\r', '\u0003\r\r', '\r\r', '', '', ''] as const;

it('projects native short/short footnotes through the opt-in public source', async () => {
  const model = await load(buildDocFixture({ text: 'Body\u0002\r',
    footnotes: [{ cp: 4, text: '\u0002Recovered note\r' }],
    noteSeparators: shortFootnoteSeparators,
    characterProperties: concat(little16(0x0855), new Uint8Array([1]), little16(0x4a43), little16(28)),
  }));
  const runs = (model.footnotes?.[0]?.content ?? []).flatMap(block => block.type === 'paragraph' ? block.runs : []);
  expect(runs.filter(run => run.type === 'text').map(run => run.text)).toContain('Recovered note');
  expect(runs.filter(run => run.type === 'text').every(run => run.fontSize === 14)).toBe(true);
});

it.each([
  { label: 'ordinary full rule', kind: 'footnotes', separators: ['\u0004\r\r', '\u0003\r\r', '', '', '', ''] },
  { label: 'partial continuation rule', kind: 'footnotes', separators: ['\u0003\r\r', '\u0003\r', '', '', '', ''] },
  { label: 'authored notice text', kind: 'footnotes', separators: ['\u0003\r\r', '\u0003\r\r', 'Notice\r\r', '', '', ''] },
  { label: 'endnote short continuation', kind: 'endnotes', separators: ['', '', '', '\u0003\r\r', '\u0003\r\r', ''] },
] as const)('draws the default separator for $label', async ({ kind, separators }) => {
  // A separator story outside the model's reserved-story shapes: the note
  // kind keeps the separator the layout draws without one.
  const model = await load(buildDocFixture({ text: 'Body\u0002\r',
    [kind]: [{ cp: 4, text: '\u0002Note\r' }], noteSeparators: separators,
    characterProperties: new Uint8Array([0x55, 0x08, 1]),
  }));
  expect(reservedSeparators(model)[kind === 'footnotes' ? 'footnote' : 'endnote']).toBeUndefined();
  expect((model as unknown as Model)[kind]).toHaveLength(1);
});

interface ReservedSeparators { footnote?: unknown; endnote?: unknown }
const reservedSeparators = (model: unknown): ReservedSeparators =>
  (model as { __noteLayoutSettings: { reservedSeparators: ReservedSeparators } }).__noteLayoutSettings.reservedSeparators;

const separatorSymbol = concat(little16(0x6a09), little16(0), little16(0x41));
it.each([
  { label: 'ordinary control glyph', offset: 0, properties: separatorSymbol },
  { label: 'continuation content-mark glyph', offset: 4, properties: separatorSymbol },
  { label: 'notice content-mark glyph', offset: 6, properties: separatorSymbol },
  { label: 'notice frame', offset: 6, properties: concat(little16(0x8418), little16(1440)) },
])('draws the default separator for a metric payload the model does not hold: $label', async ({ offset, properties }) => {
  const text = 'Body\u0002\r';
  const footnotes = [{ cp: 4, text: '\u0002Note\r' }];
  const headerStart = text.length + footnotes[0].text.length + 1;
  // Format only the selected reserved-story character. Body and note anchors
  // remain valid, so this exercises separator eligibility rather than masking
  // it behind malformed references or unrelated body formatting.
  const start = headerStart + offset;
  const model = await load(buildDocFixture({ text, footnotes, noteSeparators: shortFootnoteSeparators,
    fonts: [{ name: 'Arial', charset: 0 }],
    characterProperties: new Uint8Array([0x55, 0x08, 1]),
    formattingRuns: [
      { end: start, properties: new Uint8Array() },
      { end: start + 1, properties },
      { end: headerStart + 9, properties: new Uint8Array() },
    ],
  }));
  expect(reservedSeparators(model).footnote).toBeUndefined();
  expect((model as unknown as Model).footnotes).toHaveLength(1);
});

it('rejects a tracked insertion in a reserved separator story', async () => {
  const text = 'Body\u0002\r';
  const footnotes = [{ cp: 4, text: '\u0002Note\r' }];
  const start = text.length + footnotes[0].text.length + 1 + 6;
  await expect(load(buildDocFixture({ text, footnotes, noteSeparators: shortFootnoteSeparators,
    fonts: [{ name: 'Arial', charset: 0 }],
    characterProperties: new Uint8Array([0x55, 0x08, 1]),
    formattingRuns: [
      { end: start, properties: new Uint8Array() },
      { end: start + 1, properties: concat(little16(0x0801), new Uint8Array([1])) },
      { end: start + 3, properties: new Uint8Array() },
    ],
  }))).rejects.toMatchObject({ code: 'unsupported-content', details: { format: 'doc', reason: 'source-subset' } });
});

it.each([
  { text: 'Ax\r', characterProperties: new Uint8Array([0x55, 0x08, 1]), reason: 'invalid Word automatic note marker' },
  { text: 'A\u0002\r', characterProperties: new Uint8Array(), reason: 'Word note reference lacks special-character property' },
  { text: '😀\r', characterProperties: new Uint8Array([0x55, 0x08, 1]), reason: 'Word note reference splits Unicode character' },
])('rejects an invalid automatic note anchor: $reason', async ({ reason, ...properties }) => {
  await expect(load(buildDocFixture({ ...properties, footnotes: [{ cp: 1, text: '\u0002Note\r' }] }))).rejects.toMatchObject({ code: 'unsupported-content', details: { format: 'doc', reason: 'source-subset' } });
});

it('direct rejects an automatic reference hidden inside field instructions', async () => {
  // The converter dropped the hidden reference and its note; the direct
  // reader requires every note reference to be displayed.
  await expect(load(buildDocFixture({ text: 'A\u0013\u0002HIDDEN\u0014cached\u0015\r',
    footnotes: [{ cp: 2, text: '\u0002Note\r' }], characterProperties: concat(little16(0x0855), new Uint8Array([1])) }))).rejects.toMatchObject({ code: 'unsupported-content', details: { format: 'doc', reason: 'source-subset' } });
});

it('shares one picture resource across body, header, footnote and endnote stories', async () => {
  // MS-DOC PICFAndOfficeArtData and MS-ODRAW OfficeArtBlipPNG: one passive
  // inline picture shared by all story kinds, never an external resource.
  const record = (kind: number, options: number, data: Uint8Array) => concat(little16(options), little16(kind), little32(data.length), data);
  const png = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=', 'base64'));
  const picf = new Uint8Array(68);
  const view = new DataView(picf.buffer);
  for (const [at, value] of [[4, 68], [6, 100], [28, 1440], [30, 720], [32, 1000], [34, 1000]]) view.setUint16(at, value, true);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1), little32(0x800))),
    record(0xf00b, 0x13, concat(little16(0x0104), little32(1))),
  ));
  const data = concat(picf, shape, record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), png)));
  new DataView(data.buffer).setUint32(0, data.length, true);
  const model = await load(buildDocFixture({
    text: 'Body\u0001\u0002\u0002\r', data,
    headers: ['', 'Header\u0001\r', '', '', '', ''],
    footnotes: [{ cp: 5, text: '\u0002Footnote\u0001\r' }],
    endnotes: [{ cp: 6, text: '\u0002Endnote\u0001\r' }],
    characterProperties: concat(little16(0x0855), new Uint8Array([1]), little16(0x6a03), little32(0)),
  })) as unknown as Model;
  const images = (blocks: Block[] | undefined) => (blocks ?? []).flatMap(b => b.runs).filter(r => r.type === 'image');
  const stories = [model.body, model.headers.default?.body, model.footnotes[0].content, model.endnotes[0].content];
  expect(stories.map(story => images(story).map(r => [r.imagePath, (r as { widthPt?: number }).widthPt, (r as { heightPt?: number }).heightPt])))
    .toEqual(Array(4).fill([[images(model.body)[0]?.imagePath, 72, 36]]));
  expect(new Set(stories.flatMap(story => images(story).map(r => r.imagePath))).size).toBe(1);
});

it.skipIf(!skia).each([
  { kind: 'footnotes', short: false },
  { kind: 'endnotes', short: false },
  { kind: 'footnotes', short: true },
] as const)('renders $kind (short continuation: $short) below the body with matching reference and note numbers', async ({ kind, short }) => {
  const bytes = buildDocFixture({ text: 'Body\u0002\r', [kind]: [{ cp: 4, text: '\u0002Recovered note\r' }],
    characterProperties: new Uint8Array([0x55, 0x08, 1]), defaultTabTwips: 720,
    ...(short ? { noteSeparators: shortFootnoteSeparators } : {}) });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, legacyDoc: testDocSource() });
  try {
    const runs: { text: string; y: number }[] = [];
    for (let page = 0; page < session.pageCount; page++) await session.renderPage(page, { dpr: 1, onTextRun: r => runs.push({ text: r.text, y: r.y + page * 2000 }) });
    const body = runs.find(r => r.text.includes('Body'));
    const note = runs.find(r => r.text.includes('Recovered'));
    expect(body).toBeDefined();
    expect(note?.y).toBeGreaterThan(body?.y ?? Infinity);
    // Arabic footnotes and lowercase-Roman endnotes (the fixture's DOP).
    const mark = kind === 'footnotes' ? '1' : 'i';
    expect(runs.filter(r => r.text === mark).length).toBeGreaterThanOrEqual(2);
  } finally { await session.close(); }
});

// MS-DOC 2.8.19 aFtnIdx == 0 (custom mark) and a 2.3.3 paragraph-only
// ordinary separator story: the literal mark stays text, consumes no number,
// and the separator reserves its empty paragraph without a rule.
it.skipIf(!skia)('renders a custom-mark footnote without consuming the automatic numbering', async () => {
  const bytes = buildDocFixture({ text: 'A*B\u0002\r',
    footnotes: [{ cp: 1, automatic: false, text: '*Custom note\r' }, { cp: 3, text: '\u0002Automatic note\r' }],
    noteSeparators: ['\r\r', '\u0004\r\r', '', '', '', ''],
    characterProperties: new Uint8Array([0x55, 0x08, 1]), defaultTabTwips: 720 });
  const model = await load(bytes) as unknown as Model;
  const marks = model.body[0].runs.filter(r => r.noteRef);
  expect(marks.map(r => r.noteRef)).toEqual([
    { kind: 'footnote', id: '1', customMarkFollows: true }, { kind: 'footnote', id: '2' },
  ]);
  const separator = (model as unknown as { __noteLayoutSettings: { reservedSeparators: { footnote: { separator: { rule?: unknown; paragraph?: unknown } } } } })
    .__noteLayoutSettings.reservedSeparators.footnote.separator;
  expect(separator.rule).toBeUndefined();
  expect(separator.paragraph).toBeDefined();
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, legacyDoc: testDocSource() });
  try {
    const runs: string[] = [];
    for (let page = 0; page < session.pageCount; page++) await session.renderPage(page, { dpr: 1, onTextRun: r => runs.push(r.text) });
    // Body: A, the literal *, then the automatic reference numbered 1; the
    // custom note keeps its literal mark and the automatic note shows 1.
    expect(runs.join('')).toBe('A*B1*Custom note1Automatic note');
  } finally { await session.close(); }
});
