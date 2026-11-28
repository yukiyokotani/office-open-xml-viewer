/** Slides this package decoded from a PPTX package. A parsed slide already
 * carries its whole master text chain, so text the chain leaves uncoloured
 * renders black (the browser presentation and render worker do the same),
 * while a caller-built slide keeps the public `defaultTextColor` fallback.
 * Identity, not a model field, so the public Slide shape does not change. */
const parsedSlides = new WeakSet<object>();

export function markParsedPptxSlide(slide: object): void {
  parsedSlides.add(slide);
}

export function isParsedPptxSlide(slide: object): boolean {
  return parsedSlides.has(slide);
}

/** Record on a slide decoded by the `odp` source option that its tables
 * keep the ODF row-height rule, as data (TableElement.odfRowHeights), so a
 * copy of the model (structuredClone, JSON) keeps it and a provider that
 * omits the field gets it. Covers tables in equation-branch fallbacks too. */
export function markOdfPptxTables(slide: { elements: readonly object[] }): void {
  const visit = (elements: readonly object[]): void => {
    for (const element of elements as readonly { type?: string; odfRowHeights?: boolean; mathFallback?: { elements: readonly object[] } }[]) {
      if (element.type === 'table') element.odfRowHeights = true;
      if (element.mathFallback) visit(element.mathFallback.elements);
    }
  };
  visit(slide.elements);
}
