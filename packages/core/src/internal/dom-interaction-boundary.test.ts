import { describe, expect, it } from 'vitest';
import { eventTargetsDataAttributeWithin } from './dom-interaction-boundary.js';

describe('eventTargetsDataAttributeWithin', () => {
  it('accepts a marked target only inside the owning Viewer root', () => {
    const ownRoot = {} as Node;
    const otherRoot = {} as Node;
    const ownCard = { dataset: { ooxmlCommentId: 'own' } } as unknown as Node;
    const otherCard = { dataset: { ooxmlCommentId: 'other' } } as unknown as Node;
    Object.assign(ownRoot, { contains: (candidate: Node) => candidate === ownCard });

    expect(eventTargetsDataAttributeWithin({
      target: ownCard,
      composedPath: () => [ownCard, ownRoot],
    } as unknown as Event, ownRoot, 'ooxmlCommentId')).toBe(true);
    expect(eventTargetsDataAttributeWithin({
      target: otherCard,
      composedPath: () => [otherCard, otherRoot],
    } as unknown as Event, ownRoot, 'ooxmlCommentId')).toBe(false);
  });
});
