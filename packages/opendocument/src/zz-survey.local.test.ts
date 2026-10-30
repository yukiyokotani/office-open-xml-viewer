// Local-only probe (never committed): render admitted private ODT controls and
// dump each text run's page, top (pt) and text.
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { it } from 'vitest';
import { openDocxDocument } from '../../node/src/index.ts';
import { loadSkiaForTests } from '../../node/src/test-imports.ts';
import { testOdtSource } from './test-sources.js';

const dir = process.env.ODT_DIR as string;
const out = process.env.SURVEY_OUT as string;
it('survey', async () => {
  const skia = await loadSkiaForTests();
  if (!skia) throw new Error('no skia');
  const factory = {
    createCanvas: (w: number, h: number) => new skia.Canvas(w, h),
    loadImage: async (b: Uint8Array) => skia.loadImage(Buffer.from(b)),
  } as never;
  const lines: string[] = [];
  const only = process.env.ODT_ONLY;
  for (const name of (await readdir(dir)).filter(n => n.endsWith('.odt') && (!only || n.includes(only))).sort()) {
    try {
      const session = await openDocxDocument(new Uint8Array(await readFile(join(dir, name))), { factory, currentDate: 0, odt: testOdtSource() });
      for (let page = 0; page < session.pageCount; page++) {
        const runs: { text: string; x: number; y: number; w: number; h: number }[] = [];
        const canvas = await session.renderPage(page, { dpr: 1, onTextRun: (run: { text: string; x: number; y: number; w: number; h: number }) => runs.push(run) }) as unknown as { toBuffer(f: string): Promise<Buffer> };
        if (process.env.PNG_DIR) await writeFile(join(process.env.PNG_DIR, `${name}-${page + 1}.png`), await canvas.toBuffer('png'));
        for (const run of runs) lines.push(`${name} p${page + 1} x=${(run.x * 0.75).toFixed(2)} y=${(run.y * 0.75).toFixed(2)} h=${(run.h * 0.75).toFixed(2)} ${run.text}`);
      }
      await session.close();
    } catch (error) {
      lines.push(`${name} REJECTED ${(error as Error).message.slice(0, 80)}`);
    }
  }
  await writeFile(out, lines.join('\n') + '\n');
}, 600_000);
