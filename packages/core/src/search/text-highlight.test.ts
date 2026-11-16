import { describe, expect, it } from 'vitest';
import { loadTermSearch, adoptTextHighlightGroups } from './text-highlight.js';

describe('text highlight set', () => {
  it('is adopted without empty terms and without groups that have no term left', () => {
    expect(adoptTextHighlightGroups([])).toEqual([]);
    expect(adoptTextHighlightGroups([{ id: 'a', terms: [] }, { id: 'b', terms: [''] }])).toEqual([]);
    expect(adoptTextHighlightGroups([
      { id: 'a', terms: [] }, { id: 'b', terms: ['', ' ', 'x'], color: 'red', wholeWord: true },
    ])).toEqual([{ id: 'b', terms: [' ', 'x'], color: 'red', wholeWord: true }]);
  });

  it('is detached from the objects the caller passed', () => {
    const terms = ['x'];
    const group = { id: 'a', terms, color: 'red' };
    const adopted = adoptTextHighlightGroups([group]);
    terms[0] = 'y';
    group.color = 'blue';
    expect(adopted).toEqual([{ id: 'a', terms: ['x'], color: 'red' }]);
  });

  it('rejects a repeated group id, also when the set is otherwise empty', () => {
    const repeated = new TypeError('duplicate group id');
    expect(() => adoptTextHighlightGroups([{ id: 'a', terms: ['x'] }, { id: 'a', terms: ['y'] }])).toThrow(repeated);
    expect(() => adoptTextHighlightGroups([{ id: 'a', terms: [] }, { id: 'a', terms: [] }])).toThrow(repeated);
  });

  it('loads the term search on demand', async () => {
    expect((await loadTermSearch()).findTextHighlightRanges).toBeTypeOf('function');
  });
});
