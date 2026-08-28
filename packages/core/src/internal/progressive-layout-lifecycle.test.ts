import { describe, expect, it } from 'vitest';
import { ProgressiveLayoutLifecycle } from './progressive-layout-lifecycle.js';

describe('ProgressiveLayoutLifecycle', () => {
  it('distinguishes successful completion from terminal failure', () => {
    const lifecycle = new ProgressiveLayoutLifecycle();
    expect({ complete: lifecycle.complete, settled: lifecycle.settled }).toEqual({
      complete: true,
      settled: true,
    });

    lifecycle.begin();
    expect({ complete: lifecycle.complete, settled: lifecycle.settled }).toEqual({
      complete: false,
      settled: false,
    });

    const error = lifecycle.fail('background layout failed');
    expect(error).toBeInstanceOf(Error);
    expect({ complete: lifecycle.complete, settled: lifecycle.settled }).toEqual({
      complete: false,
      settled: true,
    });
    expect(() => lifecycle.throwIfFailed()).toThrow('background layout failed');

    lifecycle.begin();
    lifecycle.succeed();
    expect({ complete: lifecycle.complete, settled: lifecycle.settled }).toEqual({
      complete: true,
      settled: true,
    });
    expect(() => lifecycle.throwIfFailed()).not.toThrow();
  });
});
