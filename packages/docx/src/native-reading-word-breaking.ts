/** A native parser-owned reading capability. Raw bytes are retained, never
 * mapped to a valid Word word-breaking enum or replacement dictionary. */
export interface NativeReadingWordBreaking {
  readonly rawHres: number;
  readonly rawChHres: number;
}

export function acquireNativeReadingWordBreaking(value: unknown): Readonly<NativeReadingWordBreaking> | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new TypeError('Invalid native reading word-breaking owner');
  const item = value as Record<string, unknown>;
  if (Object.keys(item).length !== 2 || !Object.hasOwn(item, 'rawHres') || !Object.hasOwn(item, 'rawChHres'))
    throw new TypeError('Invalid native reading word-breaking owner');
  const byte = (v: unknown): v is number => typeof v === 'number' && Number.isInteger(v) && v >= 0 && v <= 255;
  if (!byte(item.rawHres) || !byte(item.rawChHres) || item.rawHres === 1 && item.rawChHres === 0)
    throw new TypeError('Invalid native reading word-breaking owner');
  return Object.freeze({ rawHres: item.rawHres, rawChHres: item.rawChHres });
}
