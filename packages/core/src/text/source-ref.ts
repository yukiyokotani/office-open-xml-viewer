import type { TextSourceRef } from '../types/common';

/**
 * Clip source mappings to `[start, end)` in their containing UTF-16 string and
 * rebase the returned `textStart` / `textEnd` offsets to the clipped string.
 */
export function sliceTextSourceRefs(
  refs: readonly TextSourceRef[] | undefined,
  start: number,
  end: number,
): TextSourceRef[] {
  if (!refs || end <= start) return [];

  const result: TextSourceRef[] = [];
  for (const ref of refs) {
    const overlapStart = Math.max(start, ref.textStart);
    const overlapEnd = Math.min(end, ref.textEnd);
    if (overlapEnd <= overlapStart) continue;

    const textLength = ref.textEnd - ref.textStart;
    const sourceLength = ref.sourceEnd - ref.sourceStart;
    if (textLength !== sourceLength) continue;

    result.push({
      ...ref,
      textStart: overlapStart - start,
      textEnd: overlapEnd - start,
      sourceStart: ref.sourceStart + overlapStart - ref.textStart,
      sourceEnd: ref.sourceStart + overlapEnd - ref.textStart,
    });
  }
  return result;
}

/** Rebase mappings when their containing text is appended after `offset` UTF-16 units. */
export function offsetTextSourceRefs(
  refs: readonly TextSourceRef[] | undefined,
  offset: number,
): TextSourceRef[] {
  if (!refs || refs.length === 0) return [];
  return refs.map((ref) => ({
    ...ref,
    textStart: ref.textStart + offset,
    textEnd: ref.textEnd + offset,
  }));
}
