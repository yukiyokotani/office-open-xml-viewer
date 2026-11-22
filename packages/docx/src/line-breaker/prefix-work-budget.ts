/** Library resource policy, independent of Office line-fitting behavior.
 * Exact prefix searches, joined admission and their grapheme scans (especially negative pitch) may revisit every candidate
 * on every queued tail. Bound their cumulative requested UTF-16 work per pass,
 * across semantic and emergency searches, before entering shaping. 16 Mi units
 * is a fixed work quota, not a width tolerance or a candidate-count truncation.
 * Grapheme scans are charged before allocation. Cache hits count too: constructing their text/context keys also scans text.
 * Exhaustion aborts this layout pass; no incomplete lines or alternate fit rule
 * are returned. The pass owns this counter, so retries and documents share no
 * mutable budget/cache and failure retains no input content. */
export const PREFIX_WORK_LIMIT_UTF16 = 16 * 1024 * 1024;

/** Pure preflight for optional acquisition. Refusal changes no layout state;
 * mandatory line breaking retains its throwing contract below. */
export function checkedPrefixWorkTotal(used: number, utf16Units: number, limit = PREFIX_WORK_LIMIT_UTF16): number | undefined {
  const next = used + utf16Units;
  return Number.isSafeInteger(used) && used >= 0 && Number.isSafeInteger(utf16Units)
    && utf16Units >= 0 && Number.isSafeInteger(next) && next <= limit ? next : undefined;
}

export function createPrefixWorkBudget(limit = PREFIX_WORK_LIMIT_UTF16): (utf16Units: number) => void {
  let used = 0;
  return utf16Units => {
    const next = checkedPrefixWorkTotal(used, utf16Units, limit);
    if (next === undefined) {
      throw new RangeError('DOCX line-break prefix search exceeded its UTF-16 work quota');
    }
    used = next;
  };
}

/** Conservative JSON-key input work without allocating escaped strings. JSON
 * can encode one UTF-16 unit with six units (\\uXXXX). Callers supply closed
 * request/route metadata records, never text/context or arbitrary object graphs. */
export function checkedMetadataUtf16Work(value: unknown): number | undefined {
  if (value === undefined) return 0;
  if (value === null) return 4;
  if (typeof value === 'string') return checkedPrefixWorkTotal(2, value.length * 6);
  if (typeof value === 'number' || typeof value === 'boolean') return JSON.stringify(value).length;
  if (typeof value !== 'object') return undefined;
  let work = 2;
  for (const [key, item] of Object.entries(value)) {
    if (item === undefined) continue;
    const keyWork = checkedMetadataUtf16Work(key), itemWork = checkedMetadataUtf16Work(item);
    const next = keyWork === undefined || itemWork === undefined ? undefined
      : checkedPrefixWorkTotal(work, keyWork + itemWork + 2);
    if (next === undefined) return undefined;
    work = next;
  }
  return work;
}
