import { describe, expect, it, vi } from 'vitest';
import { EmfPath } from './emf-path.js';

const context = () => ({
  beginPath: vi.fn(), moveTo: vi.fn(), lineTo: vi.fn(),
  closePath: vi.fn(), bezierCurveTo: vi.fn(),
});

describe('EMF retained path resource ownership', () => {
  it('accepts the command budget boundary and discards the entire path at overflow', () => {
    const path = new EmfPath();
    path.moveTo(0, 0);
    for (let i = 1; i < 0x10000; i++) path.lineTo(i, 0);
    const ctx = context();
    expect(path.replay(ctx)).toBe(true);
    expect(ctx.lineTo).toHaveBeenCalledTimes(0xffff);
    path.lineTo(0, 1);
    const after = context();
    expect(path.replay(after)).toBe(false);
    expect(after.beginPath).not.toHaveBeenCalled();
    path.moveTo(1, 1);
    expect(path.replay(after)).toBe(false);
  });

  it.each([Infinity, -Infinity, NaN])('invalidates the whole path for non-finite coordinates: %s', value => {
    const path = new EmfPath();
    path.moveTo(0, 0);
    path.lineTo(value, 1);
    expect(path.replay(context())).toBe(false);
  });

  it('starts a new figure after CloseFigure, not after each continuation record', () => {
    const path = new EmfPath();
    path.continueFrom(1, 2);
    path.lineTo(3, 4);
    path.continueFrom(3, 4);
    path.lineTo(5, 6);
    path.closePath();
    path.continueFrom(5, 6);
    path.lineTo(7, 8);
    const ctx = context();
    path.replay(ctx);
    expect(ctx.moveTo.mock.calls).toEqual([[1, 2], [5, 6]]);
    expect(ctx.closePath).toHaveBeenCalledTimes(1);
  });

  it('shares the allocation budget across snapshots, without charging for snapshots themselves', () => {
    const budget = { remaining: 4, replayRemaining: 100 };
    const path = new EmfPath(budget);
    path.moveTo(0, 0);
    path.lineTo(1, 1);
    const saved = path.snapshot();
    expect(budget.remaining).toBe(2);
    path.lineTo(2, 2);
    saved.lineTo(3, 3);
    expect(budget.remaining).toBe(0);
    path.lineTo(4, 4);
    expect(path.replay(context())).toBe(false);
    const ctx = context();
    expect(saved.replay(ctx)).toBe(true);
    expect(ctx.lineTo.mock.calls).toEqual([[1, 1], [3, 3]]);
  });

  it('bounds repeated replay of shared snapshots before issuing any partial drawing', () => {
    const budget = { remaining: 100, replayRemaining: 3 };
    const path = new EmfPath(budget);
    path.moveTo(0, 0);
    path.lineTo(1, 1);
    const saved = path.snapshot();
    expect(path.replay(context())).toBe(true);
    const second = context();
    expect(saved.replay(second)).toBe(false);
    expect(second.beginPath).not.toHaveBeenCalled();
  });
});
