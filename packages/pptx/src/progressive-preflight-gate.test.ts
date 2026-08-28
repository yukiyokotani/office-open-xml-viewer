import { describe, expect, it } from 'vitest';
import { ProgressivePreflightGate } from './progressive-preflight-gate.js';

describe('ProgressivePreflightGate', () => {
  it('cannot continue the next slide until the matching host acknowledgement', async () => {
    const gate = new ProgressivePreflightGate();
    let continued = false;
    const checkpoint = gate.wait(7, 1).then(() => { continued = true; });

    expect(gate.continue(8, 1)).toBe(false);
    expect(gate.continue(7, 2)).toBe(false);
    await Promise.resolve();
    expect(continued).toBe(false);

    expect(gate.continue(7, 1)).toBe(true);
    await checkpoint;
    expect(continued).toBe(true);
  });

  it('releases an obsolete checkpoint when a new parse resets the worker', async () => {
    const gate = new ProgressivePreflightGate();
    const obsolete = gate.wait(7, 1);
    gate.reset();
    await expect(obsolete).resolves.toBeUndefined();
  });
});
