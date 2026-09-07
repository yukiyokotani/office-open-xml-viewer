import { describe, expect, it, vi } from 'vitest';
import type { WasmParserHost } from '@silurus/ooxml-core';
import type { LegacyXlsDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-xls-source';
import type { OoxmlWorksheetArchive } from './worker-worksheet-source.js';
import { WorkerWorksheetSourceOwner } from './worker-worksheet-source.js';

const descriptor: LegacyXlsDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-xls-source/v1',
  builtin: 'xls',
  wasmUrl: 'https://example.test/direct-xls.wasm',
};

function directArchive(required = true) {
  return {
    measurement_request: vi.fn(() => new TextEncoder().encode(JSON.stringify({ required, font: null }))),
    configure_mdw: vi.fn(),
    parse: vi.fn(() => new Uint8Array()),
    open_sheet_cursor: vi.fn(),
    pull_sheet_cursor: vi.fn(() => new Uint8Array()),
    sheet_cursor_pull_finished: vi.fn(() => false),
    sheet_cursor_resource_usage: vi.fn(() => new Uint8Array()),
    acknowledge_sheet_cursor_terminal: vi.fn(),
    cancel_sheet_cursor: vi.fn(),
    close_sheet_cursor: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array()),
    resource_usage: vi.fn(() => new Uint8Array()),
    to_markdown: vi.fn(() => ''),
    assert_healthy: vi.fn(),
    close_workbook_session: vi.fn(),
    free: vi.fn(),
  };
}

describe('WorkerWorksheetSourceOwner', () => {
  it('closes malformed native measurement requests before exposing a source', async () => {
    const host = { archive: null } as unknown as WasmParserHost<OoxmlWorksheetArchive>;
    const archive = directArchive();
    archive.measurement_request.mockReturnValue(new TextEncoder().encode('{"required":true}'));
    const closeArchive = vi.fn();
    const owner = new WorkerWorksheetSourceOwner(host, async () => ({ archive, sourceByteLength: 0, closeArchive }));
    await expect(owner.openLegacy(new Uint8Array(), descriptor)).rejects.toThrow('invalid legacy XLS measurement request');
    expect(closeArchive).toHaveBeenCalledOnce();
    expect(owner.cursor()).toBeNull();
    expect(archive.configure_mdw).not.toHaveBeenCalled();
  });

  it('opens direct XLS without initializing or running the OOXML host', async () => {
    const host = {
      archive: null,
      run: vi.fn(),
      ensureReady: vi.fn(),
    } as unknown as WasmParserHost<OoxmlWorksheetArchive>;
    const archive = directArchive();
    const closeArchive = vi.fn();
    const open = vi.fn(async () => ({ archive, sourceByteLength: 3, closeArchive }));
    const owner = new WorkerWorksheetSourceOwner(host, open);

    expect(await owner.openLegacy(new Uint8Array([1, 2, 3]), descriptor)).toBe(archive);
    expect(host.run).not.toHaveBeenCalled();
    expect(host.ensureReady).not.toHaveBeenCalled();
    expect(archive.configure_mdw).toHaveBeenCalledOnce();
    expect(archive.configure_mdw).toHaveBeenCalledWith(undefined);
    expect(owner.execute((current) => current.parse())).toEqual(new Uint8Array());
    expect(host.run).not.toHaveBeenCalled();

    owner.closeLegacy();
    owner.closeLegacy();
    expect(closeArchive).toHaveBeenCalledOnce();
  });

  it('does not configure measurement when the native session does not request it', async () => {
    const host = { archive: null } as unknown as WasmParserHost<OoxmlWorksheetArchive>;
    const archive = directArchive(false);
    const owner = new WorkerWorksheetSourceOwner(host, async () => ({
      archive, sourceByteLength: 0, closeArchive: vi.fn(),
    }));
    await owner.openLegacy(new Uint8Array(), descriptor);
    expect(archive.configure_mdw).not.toHaveBeenCalled();
  });
});
