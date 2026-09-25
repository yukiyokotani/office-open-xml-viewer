import { describe, expect, it } from 'vitest';
import { generalHorizontalAlignment } from './renderer';

describe('generalHorizontalAlignment (ECMA-376 §18.18.40)', () => {
  it('aligns by value type', () => {
    expect(generalHorizontalAlignment('number')).toBe('right');
    expect(generalHorizontalAlignment('bool')).toBe('center');
    expect(generalHorizontalAlignment('text')).toBe('left');
  });
});
