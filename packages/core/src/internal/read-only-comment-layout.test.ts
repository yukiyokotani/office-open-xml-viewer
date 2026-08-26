import { describe, expect, it } from 'vitest';
import {
  layoutReadOnlyCommentCards,
  readOnlyCommentAuthorAccent,
} from './read-only-comment-margin.js';

describe('layoutReadOnlyCommentCards', () => {
  it('keeps sparse cards close to their anchors', () => {
    expect(layoutReadOnlyCommentCards([
      { occurrenceKey: 'upper', preferredTop: 180, height: 80 },
      { occurrenceKey: 'lower', preferredTop: 620, height: 100 },
    ], 800, 8)).toEqual(new Map([
      ['upper', 180],
      ['lower', 620],
    ]));
  });

  it('separates overlapping cards without moving an unrelated upper card', () => {
    expect(layoutReadOnlyCommentCards([
      { occurrenceKey: 'first', preferredTop: 200, height: 100 },
      { occurrenceKey: 'second', preferredTop: 220, height: 100 },
    ], 800, 10)).toEqual(new Map([
      ['first', 200],
      ['second', 310],
    ]));
  });

  it('packs from the top when the cards cannot fit in the page height', () => {
    expect(layoutReadOnlyCommentCards([
      { occurrenceKey: 'first', preferredTop: 300, height: 220 },
      { occurrenceKey: 'second', preferredTop: 500, height: 220 },
    ], 400, 10)).toEqual(new Map([
      ['first', 0],
      ['second', 230],
    ]));
  });
});

describe('readOnlyCommentAuthorAccent', () => {
  it('is stable for one normalized author without collapsing common authors into a small palette', () => {
    expect(readOnlyCommentAuthorAccent('Ada')).toBe(readOnlyCommentAuthorAccent('Ａｄａ'));
    expect(readOnlyCommentAuthorAccent('Noah Williams')).not.toBe(
      readOnlyCommentAuthorAccent('Olivia Bennett'),
    );
    expect(new Set([
      'Ada',
      'Grace',
      'Linus',
      'Maya',
      'Noah',
      'Priya',
    ].map(readOnlyCommentAuthorAccent)).size).toBe(6);
  });
});
