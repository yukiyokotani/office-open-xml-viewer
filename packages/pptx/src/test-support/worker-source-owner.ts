// Test-only stand-in for `internal/worker-presentation-source.ts`: the render
// worker imports its source owner from the URL it is given, so this module
// opens a "model source" over the same fake archive the parser path uses.
const encode = (value: unknown) => new TextEncoder().encode(JSON.stringify(value));

const archive = {
  presentation_bootstrap: () => encode({
    slideCount: 1, slideWidth: 914400, slideHeight: 914400,
    defaultTextColor: 'FFFFFF', majorFont: null, minorFont: null,
    hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
    slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
  }),
  pull_slide: () => encode({
    index: 0, slideNumber: 1, partName: 'ppt/slides/slide1.xml',
    background: null, elements: [], notes: null, hidden: false,
  }),
  slide_cursor_resource_usage: () => encode({
    archiveEntryCount: 1, declaredInflatedBytes: 1, distinctInflatedBytes: 1, operationInflatedBytes: 1,
  }),
  acknowledge_slide: () => undefined,
  cancel_slide: () => undefined,
  close_presentation_session: () => undefined,
  assert_healthy: () => undefined,
  free: () => undefined,
};

export class WorkerPresentationSourceOwner {
  constructor(_host: unknown) {}
  get kind(): 'model-source' { return 'model-source'; }
  async openModelSource(): Promise<typeof archive> { return archive; }
  cursor(): typeof archive { return archive; }
  execute<T>(operation: (value: typeof archive) => T): T { return operation(archive); }
  closeModelSource(): void {}
}
