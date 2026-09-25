import { describe, expect, it, vi } from 'vitest';
import type {
  ModelSourceModuleDescriptor,
  OpenedModelSourceModule,
  WasmParserHost,
} from '@silurus/ooxml-core';
import {
  WorkerWorksheetSourceOwner,
  type OoxmlWorksheetArchive,
  type XlsxModelSourceArchive,
} from './worker-worksheet-source.js';
import { respondToHostLayoutRequest, XLSX_HOST_LAYOUT_REQUEST, XLSX_HOST_LAYOUT_RESULT } from './host-layout.js';

const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'xlsx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};
const CALIBRI_11 = { family: 'Calibri', sizePt: 11, bold: false, italic: false };
const encode = (value: unknown) => new TextEncoder().encode(JSON.stringify(value));

function sourceArchive(request?: unknown) {
  const archive = {
    parse: vi.fn(() => new Uint8Array()),
    open_sheet_cursor: vi.fn(),
    pull_sheet_cursor: vi.fn(() => new Uint8Array()),
    sheet_cursor_pull_finished: vi.fn(() => false),
    sheet_cursor_resource_usage: vi.fn(() => new Uint8Array()),
    acknowledge_sheet_cursor_terminal: vi.fn(),
    cancel_sheet_cursor: vi.fn(),
    close_sheet_cursor: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([5])),
    assert_healthy: vi.fn(),
    configure_host_layout: vi.fn(),
    ...(request === undefined ? {} : { host_layout_request: vi.fn(() => encode(request)) }),
  };
  return archive as typeof archive & XlsxModelSourceArchive & { host_layout_request?: ReturnType<typeof vi.fn> };
}

function owner(
  archive: XlsxModelSourceArchive,
  options: { viewDefaults?: Record<string, boolean>; close?: () => void; host?: object } = {},
) {
  const close = vi.fn(options.close ?? (() => undefined));
  const open = vi.fn(async (): Promise<OpenedModelSourceModule<XlsxModelSourceArchive>> => ({
    archive, viewDefaults: options.viewDefaults ?? {}, close,
  }));
  const host = (options.host ?? { archive: null, run: vi.fn(), ensureReady: vi.fn() }) as
    unknown as WasmParserHost<OoxmlWorksheetArchive>;
  return { owner: new WorkerWorksheetSourceOwner(host, open), open, close, host };
}

describe('WorkerWorksheetSourceOwner', () => {
  it('opens a model source without the OOXML host and configures the measured Normal-font width', async () => {
    const archive = sourceArchive(CALIBRI_11);
    const { owner: sources, open, close, host } = owner(archive);
    const measure = vi.fn(() => 7);
    const transfer = [new ArrayBuffer(1)];

    await sources.openModelSource(new Uint8Array([1, 2, 3]), descriptor, measure, transfer);
    expect(open).toHaveBeenCalledExactlyOnceWith(descriptor, new Uint8Array([1, 2, 3]), transfer);
    expect(measure).toHaveBeenCalledExactlyOnceWith(CALIBRI_11);
    expect(archive.configure_host_layout).toHaveBeenCalledExactlyOnceWith(7);
    expect(archive.parse).not.toHaveBeenCalled();
    expect(sources.maximumDigitWidth).toBe(7);
    expect(sources.kind).toBe('model-source');
    expect(sources.execute((current) => current.extract_image('xl/media/1'))).toEqual(new Uint8Array([5]));
    expect((host as unknown as { run: ReturnType<typeof vi.fn> }).run).not.toHaveBeenCalled();
    expect((host as unknown as { ensureReady: ReturnType<typeof vi.fn> }).ensureReady).not.toHaveBeenCalled();

    sources.closeModelSource();
    sources.closeModelSource();
    expect(close).toHaveBeenCalledOnce();
    expect(sources.maximumDigitWidth).toBeUndefined();
  });

  it('configures no width when the source names no font or the host cannot measure it', async () => {
    const unnamed = sourceArchive(null);
    const measure = vi.fn(() => 7);
    const first = owner(unnamed).owner;
    await first.openModelSource(new Uint8Array(), descriptor, measure);
    expect(measure).not.toHaveBeenCalled();
    expect(unnamed.configure_host_layout).toHaveBeenCalledExactlyOnceWith(undefined);
    expect(first.maximumDigitWidth).toBeUndefined();

    // A non-integer or out-of-range width is not admitted.
    const unmeasured = sourceArchive(CALIBRI_11);
    const second = owner(unmeasured).owner;
    await second.openModelSource(new Uint8Array(), descriptor, () => 7.5);
    expect(unmeasured.configure_host_layout).toHaveBeenCalledExactlyOnceWith(undefined);
  });

  it('does not configure host layout when the source does not ask for it', async () => {
    const archive = sourceArchive();
    const measure = vi.fn(() => 7);
    const { owner: sources } = owner(archive);
    await sources.openModelSource(new Uint8Array(), descriptor, measure);
    expect(measure).not.toHaveBeenCalled();
    expect(archive.configure_host_layout).not.toHaveBeenCalled();
    expect(sources.maximumDigitWidth).toBeUndefined();
  });

  it('closes the source when its host layout request is invalid or its measurement fails', async () => {
    for (const request of [{ family: 'Calibri', sizePt: 11 }, { ...CALIBRI_11, sizePt: -1 }]) {
      const archive = sourceArchive(request);
      const { owner: sources, close } = owner(archive);
      await expect(sources.openModelSource(new Uint8Array(), descriptor, () => 7))
        .rejects.toThrow('invalid XLSX host layout font');
      expect(close).toHaveBeenCalledOnce();
      expect(sources.cursor()).toBeNull();
      expect(archive.configure_host_layout).not.toHaveBeenCalled();
    }

    const archive = sourceArchive(CALIBRI_11);
    const { owner: sources, close } = owner(archive);
    await expect(sources.openModelSource(new Uint8Array(), descriptor, async () => {
      throw new Error('measurement failed');
    })).rejects.toThrow('measurement failed');
    expect(close).toHaveBeenCalledOnce();
    expect(sources.cursor()).toBeNull();
  });

  it('rejects any view default because XLSX defines none', async () => {
    const archive = sourceArchive(CALIBRI_11);
    const { owner: sources, close } = owner(archive, { viewDefaults: { showGridLines: true } });
    await expect(sources.openModelSource(new Uint8Array(), descriptor, () => 7))
      .rejects.toThrow('unsupported XLSX model source view default: showGridLines');
    expect(close).toHaveBeenCalledOnce();
    expect(archive.host_layout_request).not.toHaveBeenCalled();
    expect(sources.cursor()).toBeNull();
  });

  it('degrades missing optional capabilities and closes the source on a trap', async () => {
    const { owner: sources, close } = owner(sourceArchive());
    await sources.openModelSource(new Uint8Array(), descriptor, () => undefined);
    expect(sources.resourceUsage()).toBeUndefined();
    expect(() => sources.toMarkdown()).toThrow('Markdown conversion is unsupported for this source');
    const trap = Object.assign(new Error('unreachable'), { name: 'RuntimeError' });
    expect(() => sources.execute(() => { throw trap; })).toThrow(trap);
    expect(close).toHaveBeenCalledOnce();
    expect(sources.cursor()).toBeNull();
  });

  it('rejects a source while another workbook is loaded', async () => {
    const loaded = owner(sourceArchive(), { host: { archive: {}, run: vi.fn() } });
    await expect(loaded.owner.openModelSource(new Uint8Array(), descriptor, () => 7))
      .rejects.toThrow('already loaded');
    expect(loaded.open).not.toHaveBeenCalled();
  });
});

describe('respondToHostLayoutRequest', () => {
  it('answers a host layout request with the measured width and nothing for a failed measurement', () => {
    const post = vi.fn();
    expect(respondToHostLayoutRequest(post, { type: 'other' }, () => 7)).toBe(false);
    expect(respondToHostLayoutRequest(
      post, { type: XLSX_HOST_LAYOUT_REQUEST, requestId: 3, font: CALIBRI_11 }, () => 7,
    )).toBe(true);
    expect(respondToHostLayoutRequest(
      post, { type: XLSX_HOST_LAYOUT_REQUEST, requestId: 4, font: { family: 'Calibri' } }, () => 7,
    )).toBe(true);
    expect(respondToHostLayoutRequest(
      post, { type: XLSX_HOST_LAYOUT_REQUEST, requestId: 5, font: CALIBRI_11 },
      () => { throw new Error('no canvas'); },
    )).toBe(true);
    expect(post.mock.calls.map(([message]) => message)).toEqual([
      { type: XLSX_HOST_LAYOUT_RESULT, requestId: 3, maximumDigitWidth: 7 },
      { type: XLSX_HOST_LAYOUT_RESULT, requestId: 4 },
      { type: XLSX_HOST_LAYOUT_RESULT, requestId: 5 },
    ]);
  });
});
