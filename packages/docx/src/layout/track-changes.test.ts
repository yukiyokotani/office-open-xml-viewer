import { describe, expect, it } from 'vitest';
import {
  WORD_TRACK_CHANGE_AUTHOR_COLORS,
  wordTrackChangeAuthorColor,
  wordTrackChangeDecoration,
} from './track-changes.js';

describe('track-changes paint policy (ECMA-376 §17.13.5)', () => {
  it('pins the eight track-change author colors independently of author indexing', () => {
    expect(WORD_TRACK_CHANGE_AUTHOR_COLORS).toEqual([
      '#C00000', '#0070C0', '#00B050', '#7030A0',
      '#E97132', '#196B24', '#9E480E', '#525252',
    ]);
    expect(Object.isFrozen(WORD_TRACK_CHANGE_AUTHOR_COLORS)).toBe(true);
  });

  it('maps visible track-change kinds to their revision decorations', () => {
    expect(wordTrackChangeDecoration('insertion')).toEqual({
      underline: true,
      strike: false,
    });
    expect(wordTrackChangeDecoration('deletion')).toEqual({
      underline: false,
      strike: true,
    });
    expect(wordTrackChangeDecoration(null)).toEqual({
      underline: false,
      strike: false,
    });
    expect(Object.isFrozen(wordTrackChangeDecoration('insertion'))).toBe(true);
  });

  it('assigns track-change author colors deterministically from the palette', () => {
    expect(wordTrackChangeAuthorColor(undefined)).toBe(WORD_TRACK_CHANGE_AUTHOR_COLORS[0]);
    expect(wordTrackChangeAuthorColor(null)).toBe(WORD_TRACK_CHANGE_AUTHOR_COLORS[0]);
    expect(wordTrackChangeAuthorColor('')).toBe(WORD_TRACK_CHANGE_AUTHOR_COLORS[0]);
    for (const author of ['Alice', 'Carol', 'Bob', 'Heidi']) {
      expect(WORD_TRACK_CHANGE_AUTHOR_COLORS).toContain(wordTrackChangeAuthorColor(author));
      expect(wordTrackChangeAuthorColor(author)).toBe(wordTrackChangeAuthorColor(author));
    }
    // Distinct authors hash to distinct palette slots where the eight slots allow.
    expect(wordTrackChangeAuthorColor('Alice')).not.toBe(wordTrackChangeAuthorColor('Carol'));
  });
});
