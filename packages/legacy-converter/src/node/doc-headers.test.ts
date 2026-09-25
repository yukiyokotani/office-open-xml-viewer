// Binary Word header/footer stories (MS-DOC 2.8.25 PlcfHdd, 2.8.25 PlcFld)
// through the direct DOC reader: model slots and the Node page render.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument, openDocxDocument, skia, skiaFactory } from './node-facade.js';

const load = (bytes: Uint8Array) => materializeDocxDocument(bytes, { modelSources: [testDocSource()] });
interface Run { type: string; text?: string; bold?: boolean; fieldType?: string; instruction?: string }
interface Story { body: { type: string; runs: Run[] }[] }
type Slots = Record<'default' | 'first' | 'even', Story | null>;
const storyRuns = (story: Story | null) => story?.body.map(p => p.runs) ?? null;

it('attaches all six variants after footnotes with piece formatting and no guard paragraphs', async () => {
  const model = await load(buildDocFixture({
    text: 'Body\u0002\r', footnotes: [{ cp: 4, text: '\u0002NOT A HEADER\r' }],
    headers: ['EH\r', 'OH😀\r', 'EF\r', 'OF\r', 'FH\r', 'FF\r'],
    characterProperties: concat(little16(0x0835), new Uint8Array([1, 0x55, 0x08, 1])),
    defaultTabTwips: 720, facingPages: true,
    sectionProperties: concat(little16(0x300a), new Uint8Array([1])),
  })) as unknown as { body: Story['body']; headers: Slots; footers: Slots; section: { titlePage: boolean; evenAndOddHeaders: boolean } };
  expect(model.body.flatMap(p => p.runs.map(r => r.text)).join('')).not.toContain('NOT A HEADER');
  for (const [kind, variant, text] of [
    ['headers', 'even', 'EH'], ['headers', 'default', 'OH😀'], ['footers', 'even', 'EF'],
    ['footers', 'default', 'OF'], ['headers', 'first', 'FH'], ['footers', 'first', 'FF'],
  ] as const) {
    // One paragraph per story: the guard paragraph mark is not projected.
    expect(storyRuns(model[kind][variant]), `${kind} ${variant}`).toEqual([[expect.objectContaining({ text, bold: true })]]);
  }
  expect(model.section).toMatchObject({ titlePage: true, evenAndOddHeaders: true });
});

it('keeps an explicit blank story while absent variants stay absent', async () => {
  const model = await load(buildDocFixture({ text: 'Body\r', headers: ['', '\r', '', 'Footer\r', '', ''] })) as unknown as { headers: Slots; footers: Slots };
  expect(storyRuns(model.headers.default)).toEqual([[]]);
  expect(storyRuns(model.footers.default)).toEqual([[expect.objectContaining({ text: 'Footer' })]]);
  expect([model.headers.first, model.headers.even, model.footers.first, model.footers.even]).toEqual([null, null, null, null]);
});

const footerFields = 'Page \u0013 PAGE \\* MERGEFORMAT \u001499\u0015 of \u0013NUMPAGES\u001499\u0015; \u0013INCLUDETEXT "file://secret"\u0014cached\u0015\r';

it('evaluates passive page fields and shows only the stored result of other fields', async () => {
  const model = await load(buildDocFixture({ text: 'Body\r', headers: ['', '', '', footerFields, '', ''] })) as unknown as { footers: Slots };
  const runs = storyRuns(model.footers.default)?.[0] ?? [];
  expect(runs.map(r => r.type === 'field' ? `{${r.fieldType}:${r.instruction}}` : r.text).join('')).toBe('Page {page:PAGE \\* MERGEFORMAT} of {numPages:NUMPAGES}; cached');
  expect(JSON.stringify(model)).not.toContain('file://secret');
});

it('direct rejects locked page fields instead of showing their stored result', async () => {
  // The converter kept a locked field's stored "99"; the direct reader has no
  // Office control for fLocked evaluated fields yet.
  await expect(load(buildDocFixture({ text: 'Body\r', headers: ['', '', '', footerFields, '', ''], lockedHeaderFields: true }))).rejects.toThrow(new Error('UNSUPPORTED:Word evaluated field with edited, locked or private result is not supported'));
});

it.skipIf(!skia)('renders first/even/odd header variants and evaluated page fields per page', async () => {
  const footer = 'Page \u0013PAGE\u001499\u0015 of \u0013NUMPAGES\u001499\u0015\r';
  const source = buildDocFixture({ text: 'One\fTwo\fThree\r',
    headers: ['EVEN\r', 'ODD\r', footer, footer, 'FIRST\r', footer],
    facingPages: true, defaultTabTwips: 720,
    sectionProperties: concat(little16(0x300a), new Uint8Array([1])),
  });
  const session = await openDocxDocument(source, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(3);
    for (let page = 0; page < 3; page++) {
      const runs: { text: string; y: number }[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      expect(runs.filter(r => r.y < 96).map(r => r.text).join('')).toBe(['FIRST', 'EVEN', 'ODD'][page]);
      expect(runs.filter(r => r.y > 900).map(r => r.text).join('')).toBe(`Page ${page + 1} of 3`);
    }
  } finally { await session.close(); }
});
