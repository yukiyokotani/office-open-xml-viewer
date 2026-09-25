import { describe, expect, it, vi } from 'vitest';
import { acquireXlsxSessionFromArchive, type XlsxNodeSessionArchive } from './node-acquisition.js';

const { initialize } = vi.hoisted(() => ({ initialize: vi.fn(() => { throw new Error('OOXML initialization forbidden'); }) }));
vi.mock('../wasm/xlsx_parser.js', () => ({ default: initialize, initSync: initialize, reinit: initialize }));

/** A model-source archive without the optional `resource_usage` capability. */
function fixture() {
  const workbook = { workbook: { sheets: [], date1904: false }, styles: {}, sharedStrings: [] };
  const archive = {
    parse: vi.fn(() => new TextEncoder().encode(JSON.stringify(workbook))),
  } as unknown as XlsxNodeSessionArchive;
  return { archive, workbook, sourceByteLength: 1024, closeArchive: vi.fn() };
}

describe('acquireXlsxSessionFromArchive', () => {
  it('adopts the bootstrap and host layout without OOXML WASM or invented usage', () => {
    const owned = fixture();
    const acquired = acquireXlsxSessionFromArchive({ ...owned, layoutMetrics: { maximumDigitWidth: 7 } });
    expect(acquired.workbookIndex).toEqual({ ...owned.workbook, layoutMetrics: { maximumDigitWidth: 7 } });
    expect(acquired.usage).toBeUndefined();
    expect(acquired.archive).toBe(owned.archive);
    expect(initialize).not.toHaveBeenCalled();
    expect(owned.closeArchive).not.toHaveBeenCalled();
    acquired.closeArchive();
    expect(owned.closeArchive).toHaveBeenCalledOnce();

    expect(acquireXlsxSessionFromArchive(fixture()).workbookIndex).not.toHaveProperty('layoutMetrics');
  });

  it('closes the archive when the bootstrap fails', () => {
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
