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

/** Slides an ODP model source (`odp`) decoded. Their tables keep the ODF
 * row-height rule whether or not the provider marks them (the renderer's
 * `odfTables`). Identity, like the parsed marker. */
const odfSlides = new WeakSet<object>();

export function markOdfPptxSlide(slide: object): void {
  odfSlides.add(slide);
}

export function isOdfPptxSlide(slide: object): boolean {
  return odfSlides.has(slide);
}
