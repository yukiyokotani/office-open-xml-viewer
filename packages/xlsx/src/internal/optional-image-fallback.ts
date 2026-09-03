import type { OptionalImageCodec } from '@silurus/ooxml-core';

type ImageLookup = Map<string, CanvasImageSource | null>;

const unavailableByLookup = new WeakMap<ImageLookup, Map<string, OptionalImageCodec>>();

export function clearOptionalImageUnavailable(lookup: ImageLookup): void {
  unavailableByLookup.delete(lookup);
}

export function markOptionalImageUnavailable(
  lookup: ImageLookup,
  key: string,
  codec: OptionalImageCodec,
): void {
  let entries = unavailableByLookup.get(lookup);
  if (!entries) {
    entries = new Map();
    unavailableByLookup.set(lookup, entries);
  }
  entries.set(key, codec);
}

export function isOptionalImageUnavailable(
  lookup: ImageLookup | undefined,
  key: string,
  codec: OptionalImageCodec,
): boolean {
  return lookup !== undefined && unavailableByLookup.get(lookup)?.get(key) === codec;
}
