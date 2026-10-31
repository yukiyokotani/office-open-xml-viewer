/**
 * Lookup tables keyed by document-supplied strings.
 *
 * Font names, preset names, style values and chart labels come straight from
 * untrusted documents. A plain object literal answers `table["constructor"]`
 * with `Object` and `table["__proto__"]` with `Object.prototype`, and an `in`
 * check accepts inherited names too, so such a key would be treated as a table
 * entry and crash or misdirect the caller (for example a slide whose font is
 * named "constructor" failed to render at all). Tables are often spread into
 * derived maps, which restores `Object.prototype`, so the guard belongs at the
 * lookup: every constant table indexed by a document string is read through
 * {@link ownValue} or {@link hasOwnKey}.
 */

/** `table[key]` restricted to the table's own properties. */
export function ownValue<T>(
  table: Readonly<Record<string, T>>,
  key: string | null | undefined,
): T | undefined {
  return key != null && Object.hasOwn(table, key) ? table[key] : undefined;
}

/** `key in table`, restricted to the table's own properties. */
export function hasOwnKey(table: object, key: string | null | undefined): boolean {
  return key != null && Object.hasOwn(table, key);
}

/**
 * An empty record without a prototype, for maps whose keys are document- or
 * caller-supplied. Assigning `__proto__` to a plain `{}` invokes the inherited
 * setter (the entry is dropped, or the prototype replaced), while here every
 * key, `__proto__` included, becomes an own entry. Structured clone and JSON
 * produce an ordinary object on the receiving side, so readers there still go
 * through {@link ownValue}.
 */
export function nullRecord<T>(): Record<string, T> {
  return Object.create(null) as Record<string, T>;
}
