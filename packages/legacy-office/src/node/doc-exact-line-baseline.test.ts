// Binary Word exact line spacing (MS-DOC 2.9.150 LSPD, sprmPDyaLine) through
// the ordinary DOCX layout: a DOC is a Word document, so its exact lines take
// WORD_EXACT_LINE_BASELINE, the baseline four fifths down the authored line.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument, openDocxDocument, skia, skiaFactory, type NodeCanvasFactory } from './node-facade.js';

type Recorded2D = {
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  getTransform(): { a: number; b: number; c: number; d: number; e: number; f: number };
};

function recordingFactory(draws: { text: string; y: number }[]): NodeCanvasFactory {
  const base = skiaFactory();
  return {
    loadImage: base.loadImage,
    createCanvas(width, height) {
      const canvas = base.createCanvas(width, height);
      const getContext = canvas.getContext;
      canvas.getContext = function (this: unknown, ...args: unknown[]) {
        const context = Reflect.apply(getContext, this, args) as Recorded2D | null;
        if (args[0] === '2d' && context) {
          const fillText = context.fillText;
          context.fillText = function (this: Recorded2D, ...call: Parameters<Recorded2D['fillText']>) {
            const m = this.getTransform();
            draws.push({ text: String(call[0]), y: m.b * call[1] + m.d * call[2] + m.f });
            return Reflect.apply(fillText, this, call);
          };
        }
        return context;
      } as unknown as typeof canvas.getContext;
      return canvas;
    },
  };
}

it.skipIf(!skia)('puts the baseline of an exact DOC line four fifths down the line', async () => {
  // LSPD dyaLine -600: exactly 600 twips (30pt), fMultLinespace 0.
  const exact30 = concat(little16(0x6412), little16(0x10000 - 600), little16(0));
  const bytes = buildDocFixture({ text: 'One\rTwo\r', paragraphProperties: exact30 });
  const model = await materializeDocxDocument(bytes, { legacyDoc: testDocSource() }) as unknown as {
    section: { marginTop: number }; body: { lineSpacing?: { rule: string; value: number } }[];
  };
  expect(model.body[0]!.lineSpacing).toMatchObject({ rule: 'exact', value: 30 });
  const draws: { text: string; y: number }[] = [];
  const session = await openDocxDocument(bytes, { factory: recordingFactory(draws), currentDate: 0, legacyDoc: testDocSource() });
  try {
    await session.renderPage(0, { dpr: 1 });
  } finally { await session.close(); }
  const one = draws.find(draw => draw.text.includes('One'))!;
  const two = draws.find(draw => draw.text.includes('Two'))!;
  // Device pixels per point, from the 30pt exact advance between the lines.
  const scale = (two.y - one.y) / 30;
  expect(one.y / scale - model.section.marginTop).toBeCloseTo(24, 4);
});
