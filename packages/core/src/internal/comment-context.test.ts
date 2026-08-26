import { describe, expect, it } from 'vitest';
import { boundedViewerCommentThreadContext } from './comment-context.js';

describe('boundedViewerCommentThreadContext', () => {
  it('bounds a whole thread without splitting a surrogate pair', () => {
    const result = boundedViewerCommentThreadContext(
      { id: 'root', text: 'ab😀c' },
      [{ id: 'reply', text: 'reply' }],
      3,
    );
    expect(result.thread.root.text).toBe('ab');
    expect(result.thread.replies[0]?.text).toBe('r');
    expect(result.textCharacters).toBe(3);
    expect(result.truncated).toBe(true);
  });

  it('rejects invalid public bounds', () => {
    expect(() => boundedViewerCommentThreadContext({ text: 'x' }, [], -1)).toThrow(RangeError);
  });
});
