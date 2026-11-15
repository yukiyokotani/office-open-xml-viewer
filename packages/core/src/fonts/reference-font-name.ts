const ASCII_TEXT = /^[\u0000-\u007f]*$/;

/**
 * The key `findReferenceFontMetrics` matches a family name or alias under.
 * Kept in its own module, free of the catalogue import, so a derived index of
 * the catalogue (for example PPTX's generated face-weight index) can share the
 * exact normalization without bundling the catalogue itself.
 */
export function normalizeReferenceFontFamilyName(value: string): string {
  // The alias index normalizes every catalogue alias on its first lookup in
  // each realm, so the common ASCII names take a cheap path: NFKC leaves ASCII
  // unchanged, and on ASCII the en-US case mapping equals the locale-independent
  // toLowerCase (Unicode SpecialCasing tailors only lt, tr and az).
  if (ASCII_TEXT.test(value)) return value.trim().replace(/\s+/g, ' ').toLowerCase();
  return value.normalize('NFKC').trim().replace(/\s+/gu, ' ').toLocaleLowerCase('en-US');
}
