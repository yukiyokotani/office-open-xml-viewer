import { describe, expect, it, vi } from 'vitest';
import { readXlsxArchiveBootstrap } from './archive-bootstrap.js';

const usageBytes = new TextEncoder().encode(JSON.stringify({
  archiveEntryCount: 1,
  declaredInflatedBytes: 2,
  largestInflatedEntryBytes: 2,
  distinctInflatedBytes: 2,
  operationInflatedBytes: 2,
}));

describe('readXlsxArchiveBootstrap', () => {
  it('preserves a parsed degraded workbook when archive usage is unavailable', () => {
    const workbook = { workbook: { parseError: '(zip container): invalid archive' } };
    const calls: string[] = [];

    const result = readXlsxArchiveBootstrap(
      () => {
        calls.push('parse');
        return workbook;
      },
      () => {
        calls.push('usage');
        throw 'xlsx resource usage is unavailable';
      },
    );

    expect(result).toEqual({ workbook, usage: undefined });
    expect(calls).toEqual(['parse', 'usage']);
  });

  it('returns decoded usage for a healthy archive', () => {
    const workbook = new Uint8Array([1, 2, 3]);

    expect(readXlsxArchiveBootstrap(() => workbook, () => usageBytes)).toEqual({
      workbook,
      usage: {
        archiveEntryCount: 1,
        declaredInflatedBytes: 2,
        largestInflatedEntryBytes: 2,
        distinctInflatedBytes: 2,
        operationInflatedBytes: 2,
      },
    });
  });

  it('does not downgrade parse or unrelated usage failures', () => {
    const parseError = new Error('parse failed');
    const readUsage = vi.fn(() => usageBytes);
    expect(() => readXlsxArchiveBootstrap(
      () => { throw parseError; },
      readUsage,
    )).toThrow(parseError);
    expect(readUsage).not.toHaveBeenCalled();

    const usageError = new TypeError('usage checkpoint is malformed');
    expect(() => readXlsxArchiveBootstrap(
      () => ({ workbook: {} }),
      () => { throw usageError; },
    )).toThrow(usageError);
  });
});
