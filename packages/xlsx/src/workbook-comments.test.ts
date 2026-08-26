import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import type { Worksheet } from './types.js';

describe('XlsxWorkbook.getComments', () => {
  it('returns a detached snapshot for the requested sheet', async () => {
    const workbook = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
    const worksheet = {
      comments: [{ cellRef: 'B2', author: 'Ada', text: 'Review this' }],
    } as Worksheet;
    workbook.getWorksheet = vi.fn(async () => worksheet);

    const comments = await workbook.getComments(3);

    expect(workbook.getWorksheet).toHaveBeenCalledWith(3);
    expect(comments).toEqual(worksheet.comments);
    expect(comments).not.toBe(worksheet.comments);
    expect(comments[0]).not.toBe(worksheet.comments?.[0]);
  });

  it('returns an empty list when the sheet has no comments', async () => {
    const workbook = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
    workbook.getWorksheet = vi.fn(async () => ({ comments: undefined }) as Worksheet);

    await expect(workbook.getComments(0)).resolves.toEqual([]);
  });
});
