import { afterEach, expect, it, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { googleFonts } from '@silurus/ooxml-core/internal/font-sources';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { storedZip } from '../../../tests/fixtures/chart-ex-packages.mjs';
import { XlsxWorkbook } from './workbook.js';

/** A load with font sources imports the font-source module while the worker
 * parses, before the (stubbed) parse settles; a load without sources does
 * not. Module evaluation is observed once per test file, so the load that
 * must not import it runs first. */

const { evaluations, SilentWorker } = vi.hoisted(() => ({
  evaluations: { count: 0 },
  SilentWorker: class implements WorkerLike {
    postMessage(): void {}
    addEventListener(): void {}
    removeEventListener(): void {}
    terminate(): void {}
  },
}));
vi.mock('./font-source-load.js', () => { evaluations.count++; return {}; });
vi.mock('./worker.ts?worker&inline', () => ({ default: SilentWorker }));

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

it('imports the font-source module during the parse of a load with sources only', async () => {
  vi.stubGlobal('location', { href: 'http://localhost/' });
  const stop = new Error('request stub');
  const seen: number[] = [];
  vi.spyOn(WorkerBridge.prototype, 'request').mockImplementation((async () => {
    await vi.dynamicImportSettled();
    seen.push(evaluations.count);
    throw stop;
  }) as never);
  const zip = () => (storedZip([['a.txt', 'x']]) as Uint8Array).slice().buffer as ArrayBuffer;
  await expect(XlsxWorkbook.load(zip(), {})).rejects.toBe(stop);
  await expect(XlsxWorkbook.load(zip(), { fonts: [googleFonts()] })).rejects.toBe(stop);
  expect(seen).toEqual([0, 1]);
});
