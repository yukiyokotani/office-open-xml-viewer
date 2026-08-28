export interface SectionContentPage {
  readonly pageIndex: number;
  readonly sectionRegions: readonly Readonly<{
    sectionOccurrenceId: string;
    flowDomainIds: readonly string[];
  }>[];
  readonly contentFlowDomainIds: readonly string[];
}

export interface SectionOwnedPage {
  readonly sectionOccurrenceId: string;
}

/** The first physical page containing a retained body occurrence from each section.
 * A same-page section transition can leave an empty region behind when the prior
 * section exhausted the page or the first incoming block forces a new page. Such
 * capacity is not a content appearance and must not consume the incoming section's
 * page-number restart. */
export function sectionContentFirstAppearancePageIndices(
  pages: readonly SectionContentPage[],
): ReadonlyMap<string, number> {
  const firstAppearance = new Map<string, number>();
  for (const page of pages) {
    const contentDomains = new Set(page.contentFlowDomainIds);
    for (const region of page.sectionRegions) {
      if (!firstAppearance.has(region.sectionOccurrenceId)
        && region.flowDomainIds.some((domainId) => contentDomains.has(domainId))) {
        firstAppearance.set(region.sectionOccurrenceId, page.pageIndex);
      }
    }
  }
  return firstAppearance;
}

/** Whether this is the first physical page whose page-level owner is its section. */
export function isFirstSectionOwnedPage(
  pages: readonly SectionOwnedPage[],
  pageIndex: number,
): boolean {
  const page = pages[pageIndex];
  if (!page) return false;
  return pageIndex === 0
    || pages[pageIndex - 1]?.sectionOccurrenceId !== page.sectionOccurrenceId;
}
