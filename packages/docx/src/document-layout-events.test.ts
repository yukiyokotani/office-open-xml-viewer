import { describe, expect, it, vi } from 'vitest';
import { publishDocxLayout, subscribeDocxLayout } from './document-layout-events.js';

describe('document layout publications', () => {
  it('keeps the subscription installed when the initial notification fails', () => {
    const document = {};
    const report = vi.fn();
    const listener = vi.fn()
      .mockImplementationOnce(() => { throw new Error('initial relayout failed'); });

    expect(() => subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      listener,
      report,
    )).not.toThrow();
    expect(report).toHaveBeenCalledWith(expect.objectContaining({
      message: 'initial relayout failed',
    }));

    publishDocxLayout(document, { pageCount: 2, exact: false, complete: false });
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('isolates viewer failures so one subscriber cannot abort layout or starve another', () => {
    const document = {};
    const report = vi.fn();
    let first = true;
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      () => {
        if (first) first = false;
        else throw new Error('viewer relayout failed');
      },
      report,
    );
    const healthy = vi.fn();
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      healthy,
      report,
    );
    healthy.mockClear();

    expect(() => publishDocxLayout(document, {
      pageCount: 2,
      exact: false,
      complete: false,
    })).not.toThrow();
    expect(report).toHaveBeenCalledWith(expect.objectContaining({
      message: 'viewer relayout failed',
    }));
    expect(healthy).toHaveBeenCalledOnce();
  });

  it('publishes one immutable snapshot to every subscriber', () => {
    const document = {};
    const report = vi.fn();
    let first = true;
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      (publication) => {
        if (first) first = false;
        else (publication as { pageCount: number }).pageCount = 99;
      },
      report,
    );
    const healthy = vi.fn();
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      healthy,
      report,
    );
    healthy.mockClear();

    publishDocxLayout(document, { pageCount: 2, exact: false, complete: false });

    expect(report).toHaveBeenCalledOnce();
    expect(healthy).toHaveBeenCalledWith({ pageCount: 2, exact: false, complete: false });
  });

  it('does not let an error reporter starve later subscribers', () => {
    const document = {};
    let first = true;
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      () => {
        if (first) first = false;
        else throw new Error('listener failed');
      },
      () => { throw new Error('reporter failed'); },
    );
    const healthy = vi.fn();
    subscribeDocxLayout(
      document,
      () => ({ pageCount: 1, exact: false, complete: false }),
      healthy,
      vi.fn(),
    );
    healthy.mockClear();

    expect(() => publishDocxLayout(document, {
      pageCount: 2,
      exact: false,
      complete: false,
    })).not.toThrow();
    expect(healthy).toHaveBeenCalledOnce();
  });
});
