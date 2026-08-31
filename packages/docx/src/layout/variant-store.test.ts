import { describe, expect, it } from 'vitest';
import type { DocumentLayout, LayoutServices } from './types.js';
import { LayoutVariantStore } from './variant-store.js';

function services(textFingerprint: string): LayoutServices {
  return {
    text: { fingerprint: textFingerprint } as LayoutServices['text'],
    images: { fingerprint: 'images:test' } as LayoutServices['images'],
    math: { fingerprint: 'math:test' } as LayoutServices['math'],
  };
}

function emptyLayout(): DocumentLayout {
  return { pages: [], diagnostics: [] };
}

describe('LayoutVariantStore', () => {
  it('memoizes immutable layouts by the actual service and option fingerprints', () => {
    const builds: number[] = [];
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => {
        builds.push(options.currentDateMs);
        return emptyLayout();
      },
    );

    const defaultLayout = store.defaultLayout;
    expect(store.layoutFor({ currentDateMs: 100 })).toBe(defaultLayout);
    expect(store.layoutFor({ currentDateMs: 101 })).not.toBe(defaultLayout);
    expect(store.layoutFor({ currentDateMs: 101 })).toBe(store.layoutFor({ currentDateMs: 101 }));
    expect(builds).toEqual([100, 101]);
    expect(Object.isFrozen(defaultLayout)).toBe(true);
  });

  it('does not let a non-default variant replace load-time default metadata ownership', () => {
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => ({
        pages: [],
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: String(options.currentDateMs),
        }],
      }),
    );

    const before = store.defaultLayout;
    store.layoutFor({ currentDateMs: 200 });

    expect(store.defaultLayout).toBe(before);
    expect(store.defaultLayout.diagnostics[0]?.message).toBe('100');
  });

  it('preserves the one normalized options identity at the selection boundary', () => {
    const options = Object.freeze({ currentDateMs: 100 });
    const store = new LayoutVariantStore(
      services('text:a'),
      options,
      () => emptyLayout(),
    );

    expect(store.select(options).options).toBe(options);
  });

  it('retains the default plus exactly one reusable non-default variant', () => {
    const builds: number[] = [];
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => {
        builds.push(options.currentDateMs);
        return emptyLayout();
      },
    );

    const defaultLayout = store.defaultLayout;
    const firstExplicit = store.layoutFor({ currentDateMs: 101 });
    expect(store.layoutFor({ currentDateMs: 101 })).toBe(firstExplicit);

    const secondExplicit = store.layoutFor({ currentDateMs: 102 });
    expect(secondExplicit).not.toBe(firstExplicit);
    expect(store.defaultLayout).toBe(defaultLayout);

    const rebuiltFirst = store.layoutFor({ currentDateMs: 101 });
    expect(rebuiltFirst).not.toBe(firstExplicit);
    expect(rebuiltFirst).not.toBe(secondExplicit);
    expect(store.layoutFor({ currentDateMs: 101 })).toBe(rebuiltFirst);
    expect(builds).toEqual([100, 101, 102, 101]);
  });

  it('retains the final and markup pair for one explicit field date', () => {
    const builds: string[] = [];
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => {
        builds.push(`${options.currentDateMs}:${options.showTrackedChanges === true}`);
        return emptyLayout();
      },
    );

    const datedFinal = store.layoutFor({ currentDateMs: 101 });
    const datedMarkup = store.layoutFor({ currentDateMs: 101, showTrackedChanges: true });
    expect(store.layoutFor({ currentDateMs: 101 })).toBe(datedFinal);
    expect(store.layoutFor({ currentDateMs: 101, showTrackedChanges: true })).toBe(datedMarkup);
    expect(builds).toEqual(['101:false', '101:true']);

    // Selecting another explicit date replaces the previous date's bounded
    // pair instead of retaining unbounded field-date variants.
    store.layoutFor({ currentDateMs: 102 });
    expect(store.layoutFor({ currentDateMs: 101 })).not.toBe(datedFinal);
    expect(builds).toEqual(['101:false', '101:true', '102:false', '101:false']);
  });

  it('evicts the old explicit-date pair before building the next date', () => {
    let oldPairDuringBuild: [boolean, boolean] | undefined;
    let store!: LayoutVariantStore;
    store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => {
        if (options.currentDateMs === 102) {
          oldPairDuringBuild = [
            store.hasLayoutFor({ currentDateMs: 101 }),
            store.hasLayoutFor({ currentDateMs: 101, showTrackedChanges: true }),
          ];
        }
        return emptyLayout();
      },
    );
    store.layoutFor({ currentDateMs: 101 });
    store.layoutFor({ currentDateMs: 101, showTrackedChanges: true });

    store.layoutFor({ currentDateMs: 102 });

    expect(oldPairDuringBuild).toEqual([false, false]);
  });
});
