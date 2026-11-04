/** Canvas text drawing styles `lang` (HTML §4.12.5.1.4) makes the shaper use
 * a language's OpenType forms (`locl`), e.g. Serbian italic Cyrillic. A
 * context without the attribute cannot select them: measurement and paint
 * then both ignore the language and stay equal. Callers restore the returned
 * previous value; `undefined` means nothing was changed. */
export function selectCanvasLanguage(context: object, lang: string | undefined): string | undefined {
  if (lang === undefined || !('lang' in context)) return undefined;
  const target = context as { lang: string };
  const previous = target.lang;
  target.lang = lang;
  return previous;
}

export function restoreCanvasLanguage(context: object, previous: string | undefined): void {
  if (previous !== undefined) (context as { lang: string }).lang = previous;
}
