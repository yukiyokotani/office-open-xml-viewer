import { describe, expect, it, vi } from 'vitest';
import { ProgressiveLayoutObserverNotifier } from './progressive-layout-observers.js';

describe('ProgressiveLayoutObserverNotifier', () => {
  it('isolates a throwing observer and reports it only once', () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const callback = vi.fn((_progress: { committedUnits: number }) => {
      throw new Error('observer failed');
    });
    const notifier = new ProgressiveLayoutObserverNotifier();

    expect(() => notifier.notify('onLayoutProgress', callback, { committedUnits: 1 })).not.toThrow();
    notifier.notify('onLayoutProgress', callback, { committedUnits: 2 });

    expect(callback).toHaveBeenCalledTimes(1);
    expect(consoleError).toHaveBeenCalledTimes(1);
    consoleError.mockRestore();
  });

  it('contains asynchronous observer rejection', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const callback = vi.fn(async () => { throw new Error('async observer failed'); });
    const notifier = new ProgressiveLayoutObserverNotifier();

    notifier.notify('onLayoutComplete', callback);
    await Promise.resolve();
    await Promise.resolve();

    expect(consoleError).toHaveBeenCalledTimes(1);
    consoleError.mockRestore();
  });

  it('isolates registrations when one function serves multiple observer slots', () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const callback = vi.fn((progress?: { committedUnits: number }) => {
      if (progress) throw new Error('progress observer failed');
    });
    const notifier = new ProgressiveLayoutObserverNotifier();

    notifier.notify('onLayoutProgress', callback, { committedUnits: 1 });
    notifier.notify('onLayoutProgress', callback, { committedUnits: 2 });
    notifier.notify('onLayoutComplete', callback);

    expect(callback).toHaveBeenCalledTimes(2);
    expect(callback).toHaveBeenLastCalledWith();
    expect(consoleError).toHaveBeenCalledTimes(1);
    consoleError.mockRestore();
  });
});
