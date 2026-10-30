import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import {
  USE_GOOGLE_FONTS_DEPRECATION_MESSAGE,
  googleFonts,
  resetFontDeprecationWarningsForTests,
} from '@silurus/ooxml-core/internal/font-sources';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { storedZip } from '../../../tests/fixtures/chart-ex-packages.mjs';
import { XlsxWorkbook, loadXlsxSheetSource } from './workbook.js';

/** `XlsxWorkbook.load` (and the delimited-text source) normalise the
 * deprecated `useGoogleFonts` before reading any option (design §9.1); a
 * render-worker load posts the normalised descriptors, never the key, and a
 * main-mode parse worker receives no font sources at all. The worker request
 * is captured and the load ends there. */

const { SilentWorker } = vi.hoisted(() => ({
  SilentWorker: class implements WorkerLike {
    postMessage(): void {}
    addEventListener(): void {}
    removeEventListener(): void {}
    terminate(): void {}
  },
}));

vi.mock('./worker.ts?worker&inline', () => ({ default: SilentWorker }));
vi.mock('./render-worker-host', () => ({ createRenderWorker: () => new SilentWorker() }));
vi.mock('./delimited-text-worker-host', () => ({ createDelimitedTextWorker: () => new SilentWorker() }));

const G = globalThis as Record<string, unknown>;
const ORIGINAL = { OffscreenCanvas: G.OffscreenCanvas, Worker: G.Worker, location: G.location };
afterEach(() => {
  G.OffscreenCanvas = ORIGINAL.OffscreenCanvas;
  G.Worker = ORIGINAL.Worker;
  G.location = ORIGINAL.location;
  resetFontDeprecationWarningsForTests();
  vi.restoreAllMocks();
});

const stop = new Error('request stub');
function captureRequests() {
  G.Worker = SilentWorker;
  G.OffscreenCanvas = class {};
  G.location = { href: 'http://localhost/' };
  const messages: Record<string, unknown>[] = [];
  vi.spyOn(WorkerBridge.prototype, 'request').mockImplementation(((build: (id: number) => Record<string, unknown>) => {
    messages.push(build(1));
    return Promise.reject(stop);
  }) as never);
  return messages;
}
const zip = () => (storedZip([['a.txt', 'x']]) as Uint8Array).slice().buffer as ArrayBuffer;
const csv = () => new TextEncoder().encode('a,b\n1,2\n').buffer as ArrayBuffer;

describe('XlsxWorkbook.load font options', () => {
  it('posts useGoogleFonts: true as fonts: [googleFonts()] to the render worker and warns once', async () => {
    const messages = captureRequests();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    await expect(XlsxWorkbook.load(zip(), { mode: 'worker', useGoogleFonts: true })).rejects.toBe(stop);
    await expect(XlsxWorkbook[loadXlsxSheetSource](csv(), { mode: 'worker', useGoogleFonts: true }, { format: 'csv' })).rejects.toBe(stop);
    const parses = messages.filter((message) => message.type === 'parse' || message.type === 'parseDelimitedText');
    expect(parses.map((message) => message.fonts)).toEqual([[googleFonts()], [googleFonts()]]);
    expect(parses.every((message) => !('useGoogleFonts' in message))).toBe(true);
    expect(warn.mock.calls.filter(([message]) => message === USE_GOOGLE_FONTS_DEPRECATION_MESSAGE)).toHaveLength(1);
  });

  it('sends no font sources to a main-mode parse worker, and none without sources', async () => {
    const messages = captureRequests();
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const fonts = [googleFonts({ cssOrigin: 'https://fonts.internal.example:8443' })];
    await expect(XlsxWorkbook.load(zip(), { fonts })).rejects.toBe(stop);
    await expect(XlsxWorkbook[loadXlsxSheetSource](csv(), { fonts }, { format: 'csv' })).rejects.toBe(stop);
    await expect(XlsxWorkbook.load(zip(), { mode: 'worker' })).rejects.toBe(stop);
    await expect(XlsxWorkbook.load(zip(), { mode: 'worker', fonts })).rejects.toBe(stop);
    const parses = messages.filter((message) => message.type === 'parse' || message.type === 'parseDelimitedText');
    expect(parses.map((message) => message.fonts)).toEqual([undefined, undefined, undefined, fonts]);
  });

  it('rejects useGoogleFonts: true with fonts, and malformed sources, before any worker starts', async () => {
    const messages = captureRequests();
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    await expect(XlsxWorkbook.load(zip(), { useGoogleFonts: true, fonts: [] })).rejects.toThrow(TypeError);
    await expect(XlsxWorkbook.load(zip(), { fonts: [{ kind: 'font-css-api' }] as never })).rejects.toThrow(TypeError);
    await expect(XlsxWorkbook[loadXlsxSheetSource](csv(), { fonts: [{ kind: 'font-css-api' }] as never }, { format: 'csv' }))
      .rejects.toThrow(TypeError);
    expect(messages).toEqual([]);
  });
});
