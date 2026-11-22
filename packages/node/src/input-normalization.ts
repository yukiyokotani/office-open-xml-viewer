// Internal helpers shared by the Node format entries. Not part of the public
// API: `src/index.ts` and the `./docx` / `./pptx` / `./xlsx` exports do not
// re-export this module.

/** Accept the public `ArrayBuffer | Uint8Array` input without copying a
 *  `Uint8Array` (a Node `Buffer` is one). */
export function toUint8(buffer: ArrayBuffer | Uint8Array): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}

/** Resolve DOCX `currentDate` (field evaluation clock) to epoch milliseconds. */
export function normalizeCurrentDate(value: Date | number | undefined): number {
  const current = value instanceof Date ? value.getTime() : (value ?? Date.now());
  if (!Number.isFinite(current)) {
    throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  }
  return current;
}
