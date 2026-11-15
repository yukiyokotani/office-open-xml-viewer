import { afterEach, expect, it, vi } from 'vitest';
import type { WorkerLike } from '@silurus/ooxml-core';
import { googleFonts } from '@silurus/ooxml-core/internal/font-sources';
import { PptxPresentation } from './presentation';

/** A main-mode load with font sources imports the font-source module while
 * the worker parses, before the (stubbed) parse settles; no other load does.
 * Module evaluation is observed once per test file, so the loads that must
 * not import it run first. */

const { evaluations } = vi.hoisted(() => ({ evaluations: { count: 0 } }));
vi.mock('./font-source-load.js', () => { evaluations.count++; return {}; });

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

it('imports the font-source module during a main-mode parse with sources only', async () => {
  vi.stubGlobal('Worker', SilentWorker);
  vi.stubGlobal('OffscreenCanvas', class {});
  vi.stubGlobal('location', { href: 'http://localhost/' });
  const stop = new Error('parse stub');
  const seen: number[] = [];
  vi.spyOn(PptxPresentation.prototype as unknown as { _parse(...args: unknown[]): Promise<void> }, '_parse')
    .mockImplementation(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      seen.push(evaluations.count);
      throw stop;
    });
  const fonts = [googleFonts()];
  await expect(PptxPresentation.load(new ArrayBuffer(0), {})).rejects.toBe(stop);
  await expect(PptxPresentation.load(new ArrayBuffer(0), { mode: 'worker', fonts })).rejects.toBe(stop);
  await expect(PptxPresentation.load(new ArrayBuffer(0), { fonts })).rejects.toBe(stop);
  expect(seen).toEqual([0, 0, 1]);
});
