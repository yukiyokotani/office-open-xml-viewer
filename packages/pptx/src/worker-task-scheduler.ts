/**
 * Cross a host task boundary before acknowledging a progressive worker prefix.
 *
 * `await Promise.resolve()` only drains microtasks. A MessageChannel task lets
 * the `load()` continuation enqueue its opening-slide render first; Worker
 * message ordering then puts that render ahead of the acknowledgement that
 * releases the next preflight unit. The channel is one-shot and promptly closed.
 */
export function yieldToHostTaskQueue(): Promise<void> {
  return new Promise<void>((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      channel.port2.close();
      resolve();
    };
    channel.port2.postMessage(undefined);
  });
}
