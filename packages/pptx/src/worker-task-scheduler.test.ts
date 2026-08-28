import { describe, expect, it } from 'vitest';
import { yieldToHostTaskQueue } from './worker-task-scheduler.js';

describe('yieldToHostTaskQueue', () => {
  it('crosses a worker task boundary instead of only draining microtasks', async () => {
    const order: string[] = [];
    const yielded = yieldToHostTaskQueue().then(() => order.push('continued'));
    queueMicrotask(() => order.push('host-message-dispatched'));

    await yielded;
    expect(order).toEqual(['host-message-dispatched', 'continued']);
  });
});
