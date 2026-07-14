import { describe, expect, it } from 'vitest';
import type { TextSourceRef } from '../types/common';
import { offsetTextSourceRefs, sliceTextSourceRefs } from './source-ref';

const source: TextSourceRef = {
  partName: 'word/document.xml',
  path: [{ namespaceUri: 'urn:w', localName: 't', index: 0 }],
  textStart: 2,
  textEnd: 8,
  sourceStart: 10,
  sourceEnd: 16,
};

describe('text source references', () => {
  it('clips and rebases UTF-16 source intervals', () => {
    expect(sliceTextSourceRefs([source], 4, 7)).toEqual([{
      ...source,
      textStart: 0,
      textEnd: 3,
      sourceStart: 12,
      sourceEnd: 15,
    }]);
  });

  it('rebases mappings when text is appended', () => {
    expect(offsetTextSourceRefs([source], 5)).toEqual([{
      ...source,
      textStart: 7,
      textEnd: 13,
    }]);
  });
});
