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

/** `save()` that keeps the language in effect. The drawing state includes
 * `lang` (HTML §4.12.5.1.1), and `save()` must leave the current state as it
 * was, but Chrome resets `lang` to "inherit" on the state that stays active
 * (`restore()` brings the saved value back). Without this, text drawn inside
 * a save/restore pair loses the language it was measured with. */
export function saveCanvasState(context: { save(): void; lang?: string }): void {
  const lang = 'lang' in context ? context.lang : undefined;
  context.save();
  if (lang !== undefined && context.lang !== lang) context.lang = lang;
}
