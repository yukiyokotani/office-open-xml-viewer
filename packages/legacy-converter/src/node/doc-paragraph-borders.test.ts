// Binary Word paragraph borders (MS-DOC 2.9.16 Brc, 2.9.17 Brc80,
// 2.9.18 BrcOperand) through the direct DOC reader.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument, openDocxDocument, skia, skiaFactory } from './node-facade.js';

const load = (bytes: Uint8Array) => materializeDocxDocument(bytes, { modelSources: [testDocSource()] });

it.each([false, true])('lets a no-border sentinel clear an earlier bottom border, old=%s', async old => {
  const initial = concat(little16(0x6426), new Uint8Array([8, 1, 2, 0]));
  const clear = old ? concat(little16(0x6426), new Uint8Array(4).fill(255))
    : concat(little16(0xc650), new Uint8Array([8, 0x12, 0x34, 0x56, 0, 255, 255, 255, 255]));
  const model = await load(buildDocFixture({ text: 'Body\r', paragraphProperties: concat(initial, clear) })) as unknown as { body: { borders?: unknown }[] };
  expect(model.body[0].borders).toEqual({ bottom: { style: 'none', color: null, width: 0, space: 0 } });
});

it('rejects a malformed BrcOperand length without consuming the next property', async () => {
  const bytes = buildDocFixture({ text: 'Body\r', paragraphProperties: concat(
    little16(0xc650), new Uint8Array([7, 0, 0, 0, 0, 8, 1, 0]), little16(0x2406), new Uint8Array([1]),
  ) });
  expect(await load(bytes).then(() => 'resolved', (e: unknown) => String(e instanceof Error ? e.message : e)))
    .toBe('UNSUPPORTED:invalid Word paragraph border operand');
});

it.skipIf(!skia).each([false, true])('renders one grouped paragraph rule in the body and one in the header, old=%s', async old => {
  // Red two-point bottom border; identical adjacent paragraphs form one group.
  const border = old ? concat(little16(0x6426), new Uint8Array([16, 1, 6, 2]))
    : concat(little16(0xc650), new Uint8Array([8, 255, 0, 0, 0, 16, 1, 2, 0]));
  const source = buildDocFixture({ text: 'One\rTwo\r', paragraphProperties: border,
    headers: ['', 'Running header\r', '', '', '', ''], defaultTabTwips: 720,
  });
  const session = await openDocxDocument(source, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };
    const { data } = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    const rows: number[] = [];
    for (let y = 0; y < canvas.height; y++) {
      let red = 0;
      for (let x = 0; x < canvas.width; x++) {
        const i = (y * canvas.width + x) * 4;
        if (data[i] > 200 && data[i + 1] < 80 && data[i + 2] < 80 && data[i + 3] > 200) red++;
      }
      if (red > canvas.width / 2) rows.push(y);
    }
    const groups = rows.filter((y, i) => i === 0 || y !== rows[i - 1] + 1);
    // One header rule and one rule below the pair, not one per body paragraph.
    expect(groups).toHaveLength(2);
    expect(groups[0]).toBeLessThan(96);
    expect(groups[1]).toBeGreaterThan(96);
  } finally { await session.close(); }
});
