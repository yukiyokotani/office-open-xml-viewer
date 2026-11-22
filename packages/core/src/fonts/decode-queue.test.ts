import { describe, expect, it } from 'vitest';
import { DecodeQueue } from './decode-queue.js';

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => { resolve = done; });
  return { promise, resolve };
}

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('DecodeQueue', () => {
  it('drops a queued job whose holder is released before it runs', async () => {
    const queue = new DecodeQueue(1);
    const running = deferred();
    const ran: string[] = [];
    const live = {}, released = {};
    void queue.run({}, live, async () => { ran.push('first'); await running.promise; });
    const dropped = queue.run({}, released, async () => { ran.push('released'); });
    queue.release(released);
    await dropped;
    running.resolve();
    await tick();
    expect(ran).toEqual(['first']);
    expect(queue.size).toBe(0);
  });

  it('rejects a cancelled waiter promptly, while the job keeps its other holders', async () => {
    const queue = new DecodeQueue(1);
    const running = deferred();
    const key = {};
    const controller = new AbortController();
    const other = queue.run(key, {}, () => running.promise);
    const cancelled = queue.run(key, {}, () => running.promise, controller.signal);
    controller.abort();
    await expect(cancelled).rejects.toMatchObject({ name: 'AbortError' });
    running.resolve();
    await expect(other).resolves.toBeUndefined();
  });

  it('runs a live job ahead of jobs whose holders left', async () => {
    const queue = new DecodeQueue(1);
    const running = deferred();
    const order: string[] = [];
    const gone = {}, live = {};
    void queue.run({}, live, async () => { order.push('running'); await running.promise; });
    const abandoned = Array.from({ length: 3 }, () => queue.run({}, gone, async () => { order.push('gone'); }));
    const wanted = queue.run({}, live, async () => { order.push('live'); });
    queue.release(gone);
    await Promise.all(abandoned);
    running.resolve();
    await wanted;
    expect(order).toEqual(['running', 'live']);
  });

  it('shares one job per key and bounds concurrency', async () => {
    const queue = new DecodeQueue(2);
    const gates = [deferred(), deferred(), deferred()];
    let active = 0, peak = 0, runs = 0;
    const job = (gate: ReturnType<typeof deferred>) => async () => {
      runs++; active++; peak = Math.max(peak, active);
      await gate.promise;
      active--;
    };
    const key = {};
    const waits = [queue.run(key, {}, job(gates[0]!)), queue.run(key, {}, job(gates[0]!)),
      queue.run({}, {}, job(gates[1]!)), queue.run({}, {}, job(gates[2]!))];
    await tick();
    expect(peak).toBe(2);
    for (const gate of gates) gate.resolve();
    await Promise.all(waits);
    expect(runs).toBe(3);
  });
});
