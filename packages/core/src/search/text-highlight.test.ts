import { describe, expect, it } from 'vitest';
import { loadTermSearch, validateTextHighlightGroups } from './text-highlight.js';

describe('text highlight set', () => {
  it('is empty when no group has a non-empty term', () => {
    expect(validateTextHighlightGroups([])).toBe(true);
    expect(validateTextHighlightGroups([{ id: 'a', terms: [] }, { id: 'b', terms: [''] }])).toBe(true);
    expect(validateTextHighlightGroups([{ id: 'a', terms: [] }, { id: 'b', terms: [' '] }])).toBe(false);
  });

  it('rejects a repeated group id, also when the set is otherwise empty', () => {
    const repeated = new TypeError('highlightText: group ids must be unique');
    expect(() => validateTextHighlightGroups([{ id: 'a', terms: ['x'] }, { id: 'a', terms: ['y'] }])).toThrow(repeated);
    expect(() => validateTextHighlightGroups([{ id: 'a', terms: [] }, { id: 'a', terms: [] }])).toThrow(repeated);
  });

  it('loads the term search on demand', async () => {
    expect((await loadTermSearch()).findTextHighlightRanges).toBeTypeOf('function');
  });
});
