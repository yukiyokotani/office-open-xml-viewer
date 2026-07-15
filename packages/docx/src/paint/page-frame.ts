import type { Matrix2DData } from '../layout/types.js';
import { composeAffine, inverseAffine } from './affine.js';

export interface PageFrameAdapter {
  readonly currentToPage: Matrix2DData;
}

export type PageAxisOwnership = Readonly<{
  coordinateSpace: 'physical-page-points';
  horizontal: 'page' | 'host';
  vertical: 'page' | 'host';
}>;

export function descendPageFrame(
  frame: PageFrameAdapter | undefined,
  localToParent: Matrix2DData,
): PageFrameAdapter | undefined {
  return frame ? { currentToPage: composeAffine(frame.currentToPage, localToParent) } : undefined;
}

export function pageFrameReentry(
  frame: PageFrameAdapter,
  ownership: PageAxisOwnership,
): Readonly<{ currentToTarget: Matrix2DData; targetToPage: Matrix2DData }> {
  if (ownership.coordinateSpace !== 'physical-page-points') {
    throw new Error('Anchored retained geometry must declare physical page coordinates');
  }
  const current = frame.currentToPage;
  // Anchor acquisition has already applied page/host reference-frame offsets
  // independently per axis. The ownership flags govern later relocation only;
  // the retained box itself remains one non-singular physical page frame.
  const targetToPage = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const pageToCurrent = inverseAffine(current);
  if (!pageToCurrent) throw new Error('Current retained coordinate frame is not invertible');
  return {
    currentToTarget: composeAffine(pageToCurrent, targetToPage),
    targetToPage,
  };
}
