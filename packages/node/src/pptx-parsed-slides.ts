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
