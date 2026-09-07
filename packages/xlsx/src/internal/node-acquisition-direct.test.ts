import { describe, expect, it, vi } from 'vitest';
import { acquireXlsxSessionFromArchive, type XlsxNodeArchive } from './node-acquisition.js';

const { initialize } = vi.hoisted(() => ({ initialize: vi.fn(() => { throw new Error('OOXML initialization forbidden'); }) }));
vi.mock('../wasm/xlsx_parser.js', () => ({ default: initialize, initSync: initialize, reinit: initialize }));

function fixture() {
  const workbook = { workbook: { sheets: [], date1904: false }, styles: {}, sharedStrings: [] };
  const archive = {
    parse: vi.fn(() => new TextEncoder().encode(JSON.stringify(workbook))),
    resource_usage: vi.fn(() => { throw new Error('xlsx resource usage is unavailable'); }),
  } as unknown as XlsxNodeArchive;
  return { archive, workbook, sourceByteLength: 1024, closeArchive: vi.fn() };
}

describe('Node native worksheet source adoption', () => {
  it('adopts bootstrap without OOXML WASM initialization or fake archive metrics', () => {
    const owned = fixture();
    const acquired = acquireXlsxSessionFromArchive(owned);
    expect(acquired.workbookIndex).toEqual(owned.workbook);
    expect(acquired.usage).toBeUndefined();
    expect(acquired.archive).toBe(owned.archive);
    expect(initialize).not.toHaveBeenCalled();
    expect(owned.closeArchive).not.toHaveBeenCalled();
    acquired.closeArchive();
    expect(owned.closeArchive).toHaveBeenCalledOnce();
  });

  it('closes ownership when bootstrap fails', () => {
    const owned = fixture();
    vi.mocked(owned.archive.parse).mockImplementation(() => { throw new Error('bad bootstrap'); });
    expect(() => acquireXlsxSessionFromArchive(owned)).toThrow('bad bootstrap');
    expect(owned.closeArchive).toHaveBeenCalledOnce();
  });

  it('closes without parsing when already aborted', () => {
    const owned = fixture();
    const controller = new AbortController();
    controller.abort();
    expect(() => acquireXlsxSessionFromArchive(owned, { signal: controller.signal })).toThrow(/aborted/);
    expect(owned.archive.parse).not.toHaveBeenCalled();
    expect(owned.closeArchive).toHaveBeenCalledOnce();
  });
});
