import { describe, expect, it } from 'vitest';
import { ConditionalFormattingNotice, CF_NOTICE_TEXT } from './cf-notice.js';
import { createCfReport } from '../../cf-diagnostics.js';

const VIEWPORT = { row: 1, col: 1, rows: 1, cols: 1 };

describe.skipIf(typeof document === 'undefined')('ConditionalFormattingNotice', () => {
  it('shows only for a committed report with diagnostics, and clears', () => {
    const parent = document.createElement('div');
    const notice = new ConditionalFormattingNotice(document, parent);
    expect(notice.element.hidden).toBe(true);
    expect(notice.element.getAttribute('role')).toBe('status');
    notice.update(createCfReport(0, VIEWPORT, [
      { kind: 'unsupported', phase: 'expression', blockIndex: 0, ruleIndex: 0, row: 1, col: 1 },
    ]));
    expect(notice.element.hidden).toBe(false);
    expect(notice.element.textContent).toBe(CF_NOTICE_TEXT);
    // A supported committed frame clears the warning.
    notice.update(createCfReport(0, VIEWPORT, []));
    expect(notice.element.hidden).toBe(true);
    // No report for a committed frame never keeps a stale claim.
    notice.update(createCfReport(0, VIEWPORT, [
      { kind: 'invalid', phase: 'cellIs', blockIndex: 0, ruleIndex: 0, row: 1, col: 1 },
    ]));
    notice.update(null);
    expect(notice.element.hidden).toBe(true);
    notice.destroy();
    expect(parent.children.length).toBe(0);
  });
});
