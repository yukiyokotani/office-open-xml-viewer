import { afterEach, describe, expect, it } from 'vitest';
import {
  assertAndDeepFreezeDocumentLayout,
  deepFreezeDocumentLayout,
} from './invariants.js';
import { LayoutInvariantError } from './diagnostics.js';
import { sealPlainData, snapshotPlainData } from './plain-data.js';
import {
  documentLayoutValidationEnabled,
  setDocumentLayoutValidation,
} from './validation-policy.js';
import type { DocumentLayout } from './types.js';

// ─────────────────────────────────────────────────────────────────────────────
// The validation policy gates ONLY the path-precise `assertPlainData` pre-pass
// (a second walk whose sole extra value is the exact property path in the
// error). Fatal state — non-finite geometry, invariant violations — is
// detected unconditionally, per the layout engine's error contract: it must
// surface as a typed error in production too, never as a paint defect. This
// suite pins that split: what always throws, what always freezes, and the one
// thing the switch actually controls.
// ─────────────────────────────────────────────────────────────────────────────

/** A layout carrying a non-finite number: fatal state that must be rejected
 *  whether or not the development-only pre-pass runs. */
function layoutWithNonFiniteGeometry(): DocumentLayout {
  return {
    pages: [{ geometry: { widthPt: Number.NaN, heightPt: 792 } }],
    diagnostics: [],
  } as unknown as DocumentLayout;
}

afterEach(() => {
  // Restore the suite-wide default so ordering can never leak policy.
  setDocumentLayoutValidation(true);
});

describe('document layout validation policy', () => {
  it('defaults to enabled under a test runner', () => {
    expect(documentLayoutValidationEnabled()).toBe(true);
  });

  it('reports the configured state', () => {
    setDocumentLayoutValidation(false);
    expect(documentLayoutValidationEnabled()).toBe(false);
    setDocumentLayoutValidation(true);
    expect(documentLayoutValidationEnabled()).toBe(true);
  });

  it('rejects non-finite geometry while enabled', () => {
    setDocumentLayoutValidation(true);
    expect(() => assertAndDeepFreezeDocumentLayout(layoutWithNonFiniteGeometry()))
      .toThrow(LayoutInvariantError);
  });

  it('still rejects non-finite geometry while disabled', () => {
    // The contract: fatal state is not gated. The message loses its precise
    // property path without the pre-pass, nothing more.
    setDocumentLayoutValidation(false);
    expect(() => assertAndDeepFreezeDocumentLayout(layoutWithNonFiniteGeometry()))
      .toThrow(LayoutInvariantError);
    expect(() => deepFreezeDocumentLayout(layoutWithNonFiniteGeometry()))
      .toThrow(LayoutInvariantError);
    expect(() => snapshotPlainData({ widthPt: Number.NaN }, 'test'))
      .toThrow(/finite numbers/);
    expect(() => sealPlainData({ widthPt: Number.POSITIVE_INFINITY }, 'test'))
      .toThrow(/finite numbers/);
  });

  it('freezes via deepFreezeDocumentLayout whether or not the pre-pass runs', () => {
    setDocumentLayoutValidation(false);
    const layout = {
      pages: [{ geometry: { widthPt: 612, heightPt: 792 } }],
      diagnostics: [],
    } as unknown as DocumentLayout;
    const frozen = deepFreezeDocumentLayout(layout);
    expect(Object.isFrozen(frozen)).toBe(true);
    expect(Object.isFrozen(frozen.pages)).toBe(true);
    expect(Object.isFrozen(frozen.pages[0])).toBe(true);
    expect(Object.isFrozen(frozen.pages[0].geometry)).toBe(true);
  });

  it('keeps snapshotPlainData sealing and cloning while disabled', () => {
    setDocumentLayoutValidation(false);
    const source = { a: 1, nested: { b: [2, 3] } };
    const snapshot = snapshotPlainData(source, 'test');
    expect(snapshot).toEqual(source);
    expect(snapshot).not.toBe(source);
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.nested)).toBe(true);
    expect(Object.isFrozen(snapshot.nested.b)).toBe(true);
  });

  it('still reports genuinely non-cloneable data while disabled', () => {
    // The clone remains the backstop: only the precise property path in the
    // message is development-only.
    setDocumentLayoutValidation(false);
    expect(() => snapshotPlainData({ fn: () => 1 }, 'test')).toThrow(TypeError);
  });

  it('gates only the path-precise pre-pass', () => {
    // An accessor property is invisible to the clone (it reads the produced
    // value) but rejected by the descriptor-checking pre-pass — the one check
    // that is genuinely development-only.
    const accessor = Object.defineProperty({}, 'a', {
      enumerable: true,
      get: () => 1,
    });
    setDocumentLayoutValidation(true);
    expect(() => snapshotPlainData(accessor, 'test')).toThrow(TypeError);
    setDocumentLayoutValidation(false);
    expect(snapshotPlainData(accessor, 'test')).toEqual({ a: 1 });
  });
});
