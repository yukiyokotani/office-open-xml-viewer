import { describe, expect, it } from 'vitest';
import { loadTermSearch, textHighlightGroupsAreEmpty } from './text-highlight.js';

describe('text highlight set', () => {
  it('is empty when no group has a non-empty term', () => {
    expect(textHighlightGroupsAreEmpty([])).toBe(true);
    expect(textHighlightGroupsAreEmpty([{ id: 'a', terms: [] }, { id: 'b', terms: [''] }])).toBe(true);
    expect(textHighlightGroupsAreEmpty([{ id: 'a', terms: [] }, { id: 'b', terms: [' '] }])).toBe(false);
  });

  it('loads the term search on demand', async () => {
    expect((await loadTermSearch()).findTextHighlightRanges).toBeTypeOf('function');
  });
});
