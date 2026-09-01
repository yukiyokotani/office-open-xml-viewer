export interface AnnouncementSection {
  readonly title: string;
  /** Published packages or integration modules affected by the section. */
  readonly modules?: readonly string[];
  /** Product/API reason for the change, separate from migration mechanics. */
  readonly rationale?: string;
  readonly paragraphs: readonly string[];
  readonly bullets?: readonly string[];
  readonly kind?: 'summary';
  readonly examples?: readonly AnnouncementExample[];
}

export interface AnnouncementExample {
  readonly title: string;
  readonly code: string;
}

export interface AnnouncementImage {
  /** Site-public path. Announcement artwork is deliberately static and local. */
  readonly src: string;
  readonly alt: string;
  readonly caption?: string;
}

export interface Announcement {
  readonly slug: string;
  readonly date: string;
  readonly label: 'Upcoming release' | 'Engineering note' | 'Release note';
  readonly version?: string;
  readonly title: string;
  readonly summary: string;
  readonly audience: string;
  readonly image?: AnnouncementImage;
  readonly sections: readonly AnnouncementSection[];
}

export const announcements: readonly Announcement[] = [
  {
    slug: 'v0841-word-layout-refinements',
    date: '2026-09-01',
    label: 'Release note',
    version: 'v0.84.1',
    title: 'Word layout refinements in v0.84.1',
    summary: 'v0.84.1 improves legacy text-box spacing and keeps page-break behavior carefully limited to the Word documents it was designed for.',
    audience: 'Applications that display DOCX files. Existing setup and viewer options continue to work unchanged.',
    sections: [
      {
        title: 'More faithful, more focused',
        kind: 'summary',
        paragraphs: [
          'Legacy Word text boxes now keep their authored inner spacing in an additional setting combination. Table rows also stay together at the verified Word page boundaries, while other table layouts keep their previous behavior.',
        ],
        bullets: [
          'Preserve authored spacing in more legacy Word text boxes.',
          'Keep the page-break adjustment limited to the verified cases.',
          'Upgrade without changing application code.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No migration is required from v0.84.0. TIFF support and existing viewer integrations remain unchanged.',
        ],
      },
    ],
  },
  {
    slug: 'v084-tiff-images',
    date: '2026-09-01',
    label: 'Release note',
    version: 'v0.84.0',
    title: 'TIFF images in v0.84.0',
    summary: 'v0.84.0 adds opt-in TIFF image display across Word, Excel and PowerPoint files. Try Yours and the VS Code extension include it automatically.',
    audience: 'Applications that open DOCX, XLSX or PPTX files containing TIFF images. Existing viewers continue to work without source changes.',
    sections: [
      {
        title: 'TIFF images across Office files',
        kind: 'summary',
        paragraphs: [
          'Supported TIFF images can now appear in Word documents, Excel workbooks and PowerPoint presentations instead of being left blank. The same image support is available in regular and worker rendering modes.',
        ],
        bullets: [
          'Use one shared TIFF module with DOCX, XLSX and PPTX viewers.',
          'Keep TIFF code out of applications that do not need it.',
          'Open TIFF-containing files without extra setup in Try Yours or the VS Code extension.',
        ],
      },
      {
        title: 'Choose the integration',
        paragraphs: [
          'Library applications opt in by importing tiff from @silurus/ooxml/tiff and passing it to the Viewer or document engine. The Production decisions page lists every optional module and shows the complete setup.',
          'The official Try Yours experience and the VS Code extension enable all first-party optional renderers, including TIFF, so end users receive the full viewing feature set.',
          'The TIFF module is built for images inside Office files, not as a general-purpose TIFF library. As a small by-product, it can also provide a simple preview of a supported standalone TIFF file.',
        ],
        examples: [
          {
            title: 'Enable TIFF images',
            code: `import { DocxViewer } from '@silurus/ooxml/docx';
import { tiff } from '@silurus/ooxml/tiff';

const viewer = new DocxViewer(canvas, { tiff });
await viewer.load(source);`,
          },
        ],
      },
      {
        title: 'More faithful Word page layout',
        paragraphs: [
          'Word documents now preserve the intended inner margins of legacy text boxes and avoid separating parallel table-cell content when a row has no safe place to break on the current page.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No migration is required. Existing applications keep the same defaults, and applications that do not display TIFF images do not need to add the module.',
          'The initial release supports a bounded set of TIFF images commonly embedded by Office. Unsupported TIFF variants are skipped without stopping the rest of the document from rendering.',
        ],
      },
    ],
  },
  {
    slug: 'v083-progressive-viewing',
    date: '2026-08-28',
    label: 'Release note',
    version: 'v0.83.0',
    title: 'Open large Word and PowerPoint files sooner in v0.83.0',
    summary: 'v0.83.0 can show the opening pages or slides sooner for large Word and PowerPoint files, speeds up large Word tables, and uses embedded PowerPoint fonts when available.',
    audience: 'Applications that open larger DOCX or PPTX files, or need closer PowerPoint font fidelity. Existing viewer setup continues to work unchanged.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'Large Word documents and PowerPoint presentations can now become useful before all background preparation finishes. Progressive viewing is opt-in, so existing applications keep their current loading behavior.',
        ],
        bullets: [
          'Show the opening pages of a DOCX file while the remaining pages are prepared.',
          'Show the opening slide of a PPTX file while later slides are prepared, with a stable scrollbar from the first view.',
          'Open documents with large Word tables faster and render PowerPoint text with embedded fonts when the presentation provides them.',
        ],
      },
      {
        title: 'Start viewing sooner',
        paragraphs: [
          'Enable progressiveLayout on a Word or PowerPoint Viewer when showing useful content quickly matters more than waiting for the entire document. It works with both the regular and worker rendering modes; worker mode also keeps more of the remaining work away from the application UI.',
          'Word publishes pages as pagination advances. PowerPoint knows the final slide count at the start, so its scrollbar remains stable and slides that are not ready yet show a loading state.',
          'Until Word pagination completes, pageCount means the pages available so far. PowerPoint slideCount is final from the first view, while availableSlideCount reports how many opening slides are ready.',
        ],
        examples: [
          {
            title: 'Enable progressive viewing',
            code: `import { DocxScrollViewer } from '@silurus/ooxml/docx';
import { PptxScrollViewer } from '@silurus/ooxml/pptx';

const wordViewer = new DocxScrollViewer(wordContainer, {
  progressiveLayout: true,
});

const slideViewer = new PptxScrollViewer(slideContainer, {
  progressiveLayout: true,
});`,
          },
        ],
      },
      {
        title: 'Rendering and review improvements',
        paragraphs: [
          'Documents with large Word tables avoid repeated pagination work, reducing the wait for the completed document.',
          'Long web addresses in Word documents now wrap naturally within the page instead of leaving a large gap or being clipped at the edge.',
          'PowerPoint can use fonts embedded in a presentation in both rendering modes. Text that extends beyond its shape also remains selectable.',
          'Word can now display tracked changes with author-coloured underlines, strikethroughs and margin change bars using showTrackedChanges: true. The accepted-final view remains the default.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No migration is required. Progressive viewing and tracked-change markup are both opt-in, and existing Viewer options keep their previous defaults.',
          'If an operation needs the completed document, such as printing, exporting or showing a final Word page count, await waitUntilLayoutComplete() first. Ordinary viewing can begin as soon as load() resolves.',
        ],
      },
      {
        title: 'Technical note',
        paragraphs: [
          'DOCX progressive layout resumes one pagination session instead of rebuilding earlier pages. PPTX prepares slides in order after an initial presentation bootstrap, which makes the final slide count and scroll extent available from first paint. Both formats share the same progressive callbacks and completion-waiting lifecycle. Worker mode reduces competition with the application UI; it does not guarantee a shorter total preparation time.',
        ],
      },
    ],
  },
  {
    slug: 'v082-review-comments',
    date: '2026-08-26',
    label: 'Release note',
    version: 'v0.82.0',
    title: 'Comments and tracked changes in v0.82.0',
    summary: 'v0.82.0 adds read-only comment presentation across Word, Excel and PowerPoint, together with detached Word tracked-change data.',
    audience: 'Applications that display review comments or Word tracked changes. Existing viewers continue to work without source changes.',
    sections: [
      {
        title: 'Comments in context',
        kind: 'summary',
        paragraphs: [
          'DOCX and PPTX ScrollViewers can place comment cards beside the page or slide, close to the authored location. XLSX viewers retain cell-anchored comment markers and cards. Replies and resolved state are kept when present in the file.',
          'The feature is read-only. Editing, replying, resolving and application-specific review workflows remain the responsibility of the host application.',
        ],
      },
      {
        title: 'Use the built-in presentation',
        paragraphs: [
          'Pass comments: true to a DOCX or PPTX ScrollViewer to use the built-in margin. The same comments option controls visibility and resolved-thread policy across DOCX, XLSX and PPTX. XLSX keeps its existing default comment presentation.',
          'The default structure is intentionally simple. Applications can adjust its appearance with stable CSS classes and custom properties without recreating the Viewer.',
        ],
        examples: [
          {
            title: 'Show DOCX comments',
            code: `import { DocxScrollViewer } from '@silurus/ooxml/docx';

const viewer = new DocxScrollViewer(container, {
  comments: true,
});

await viewer.load(source);`,
          },
        ],
      },
      {
        title: 'Compose a different UI',
        paragraphs: [
          'Applications that need a different structure can use format-scoped comment data and anchor geometry. DOCX resolves comment threads per rendered page, PPTX exposes comments per slide and element bounds, and XLSX exposes comments per sheet plus current-viewport cell rectangles.',
          'This lower-level path leaves framework components, interaction, list virtualization and editing workflows under application ownership.',
        ],
      },
      {
        title: 'Word tracked changes',
        paragraphs: [
          'DOCX parsing now retains recorded insertions, deletions and moves from the document body as detached data. The rendered document is the accepted-final state: deletions and move sources do not appear on the page, and there is no built-in tracked-change markup view.',
          'Comments and tracked changes remain separate document concepts and separate public data. The Viewer does not combine them into an editable review model.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No existing option is removed or renamed. DOCX and PPTX comments appear only when enabled, while XLSX preserves its established default. Applications that do not display review information require no changes.',
        ],
      },
    ],
  },
  {
    slug: 'v081-chartex-opt-in',
    date: '2026-08-25',
    label: 'Release note',
    version: 'v0.81.0',
    title: 'Migrating to v0.81.0',
    summary: 'v0.81.0 expands Microsoft ChartEx rendering and moves it to an opt-in module. Applications that display ChartEx charts must import and enable the renderer.',
    audience: 'Applications that display waterfall, histogram, Pareto, funnel, box-and-whisker, treemap or sunburst charts. Applications that use only classic charts need no changes.',
    sections: [
      {
        title: 'ChartEx support',
        kind: 'summary',
        paragraphs: [
          'v0.81.0 expands rendering for Microsoft ChartEx chart families, including waterfall, histogram, Pareto, funnel, box-and-whisker, treemap and sunburst charts.',
          'Classic charts remain built in and require no application changes.',
        ],
      },
      {
        title: 'Migration',
        paragraphs: [
          'ChartEx is now provided by the separate @silurus/ooxml/chart-ex module. Import chartEx and pass it to any DOCX, XLSX or PPTX viewer that needs these chart families.',
        ],
        examples: [
          {
            title: 'Enable ChartEx',
            code: `import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { chartEx } from '@silurus/ooxml/chart-ex';

const viewer = new XlsxViewer(container, { chartEx });
await viewer.load(source);`,
          },
        ],
      },
    ],
  },
  {
    slug: 'v080-worker-rendering',
    date: '2026-08-16',
    label: 'Release note',
    version: 'v0.80.0',
    title: 'Built-in renderer parity for worker mode in v0.80.0',
    summary: 'v0.80.0 extends the existing DOCX, XLSX and PPTX worker mode so the built-in math, 3-D chart and Region Map renderers use the same injection API and rendering path as main-thread mode.',
    audience: 'Browser applications that use worker mode and need equations, authored 3-D charts or country-level Region Maps without moving document layout and paint back to the main thread. Main-thread rendering remains the default, so existing applications do not need to change.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/math', '@silurus/ooxml/three-d', '@silurus/ooxml/region-map'],
        rationale: 'Choose worker mode for a more responsive UI during heavier rendering, or keep the default main mode for the smallest and simplest setup.',
        paragraphs: [
          'Worker rendering across DOCX, XLSX and PPTX was introduced in v0.59.0. v0.80.0 closes its remaining built-in renderer gaps: equations, 3-D charts and Region Maps now use the same math, threeD and regionMap options in main-thread and worker modes.',
          'Main-thread mode remains the default. No migration is required for existing applications.',
        ],
        bullets: [
          'Use main mode for smaller documents, the smallest worker download, the lowest single-frame overhead or custom renderer objects.',
          'Use worker mode when larger or more complex documents make scrolling, navigation or other application UI less responsive.',
          'The built-in math, 3-D chart and Region Map renderers now use the same injection options in both modes.',
          'Selection, find, navigation and viewer interactions retain the same public APIs.',
        ],
      },
      {
        title: 'Choose the mode that fits your app',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Worker mode trades a larger download and bitmap-transfer overhead for less layout and paint work on the browser UI thread.',
        paragraphs: [
          'In main mode, parsing already runs in a Worker, while layout and Canvas rendering run on the main thread. It is the best default for ordinary previews, smaller files and applications that prioritize the smallest download and lowest per-frame transfer overhead.',
          'In worker mode, parsing, layout and Canvas rendering run in a Web Worker. Choose it when rendering a larger document competes with scrolling, navigation, animation or other UI work in your application. Viewer controls and interactions remain available in both modes.',
          'A DOCX document that needs browser-only OpenType vertical-glyph selection automatically uses main mode for correct text shaping. Read the loaded document\'s mode when your integration needs to observe that fallback.',
          'Worker mode requires Worker and OffscreenCanvas support. It improves responsiveness rather than guaranteeing faster total rendering time, and it is not a separate process or a memory-safety boundary.',
        ],
      },
      {
        title: 'Use the same options in either mode',
        modules: ['@silurus/ooxml/math', '@silurus/ooxml/three-d', '@silurus/ooxml/region-map'],
        rationale: 'Switching modes should not require a second setup for equations, 3-D charts or Region Maps.',
        paragraphs: [
          'Pass math, threeD and regionMap exactly as in main-thread mode. The library recognizes its built-in renderers and reconstructs them inside the worker without exposing worker protocol objects through the public renderer contracts.',
          'Custom renderer objects remain a main-mode feature because arbitrary JavaScript objects cannot be transferred into a Worker. Use the built-in math, threeD and regionMap exports when those capabilities are required in worker mode.',
        ],
        examples: [
          {
            title: 'Render an XLSX Viewer off the main thread',
            code: `import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { math } from '@silurus/ooxml/math';
import { threeD } from '@silurus/ooxml/three-d';
import { regionMap } from '@silurus/ooxml/region-map';

const viewer = new XlsxViewer(container, {
  mode: 'worker',
  math,
  threeD,
  regionMap,
});

await viewer.load(source);`,
          },
        ],
      },
      {
        title: 'Trade-offs to consider',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Worker mode is primarily a UI-responsiveness choice, not an automatic speed or memory improvement.',
        paragraphs: [
          'Worker mode downloads a larger self-contained worker asset and transfers a rendered bitmap for each frame. Main mode has less worker code to download and avoids that frame transfer, but heavy layout or paint can occupy the UI thread.',
          'Both modes keep Viewer navigation, zoom, virtualized scrolling, selection, find, hyperlinks, equations, 3-D charts and Region Maps available. PowerPoint media controls and other DOM overlays continue to be presented by the Viewer in either mode.',
        ],
      },
      {
        title: 'Technical note',
        paragraphs: [
          'JavaScript renderer functions cannot cross the structured-clone boundary. Instead of exposing a second worker-specific API, v0.80 keeps the public math, threeD and regionMap objects as ordinary renderer contracts and records the built-in module identity privately. The worker reconstructs only those recognized built-ins, so application code uses the same options while transport details stay out of the public types.',
          'Production packaging was the less obvious part. A consumer bundler can treat a published Worker as an opaque asset: copying the entry file while leaving its split chunks behind, or rebasing a MathJax URL against the consumer output directory. The published render worker is therefore self-contained, while browser-resolved external asset URLs are handed across explicitly. Tests cover both the raw package output and a fresh Vite consumer rebundle.',
          'Math output also needed one drawing contract in both realms. Equations are rasterized through the same Canvas path in Window and Worker contexts on a size-bounded surface. A 256 px/em source is reduced in two stages and cached at 64 px/em for cleaner 100% display without turning ordinary document text into vector geometry.',
          'Finally, the worker path is compared against main mode in the same browser for public DOCX, XLSX and PPTX examples, equations, 3-D charts and Region Maps. The exercised frames are pixel-identical; CI retains a small tolerance only for browser text rasterization differences across environments.',
        ],
      },
    ],
  },
  {
    slug: 'v079-chart-rendering-addons',
    date: '2026-08-14',
    label: 'Release note',
    version: 'v0.79.0',
    title: '3-D charts, Region Maps and chart fidelity in v0.79.0',
    summary: 'v0.79.0 adds opt-in 3-D and offline Region Map renderers while improving shared chart axes, labels, legends and modern chart families across DOCX, XLSX and PPTX.',
    audience: 'Applications that display Office charts. Existing applications keep the established chart fallback without source changes; import the optional renderer modules only when authored 3-D charts or country-level Region Maps are required.',
    image: {
      src: '/announcements/chart-rendering-v079.webp',
      alt: 'A grid of eight synthetic chart renderings: 3-D columns, 3-D pie, an offline country Region Map, a combo chart with minor ticks, waterfall, treemap, bubble and box-and-whisker charts.',
      caption: 'Rendered by @silurus/ooxml from synthetic data. The map uses public-domain Natural Earth geometry.',
    },
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/three-d', '@silurus/ooxml/region-map'],
        rationale: '3-D camera and map geometry stay out of the default bundles unless an application enables them.',
        paragraphs: [
          'v0.79.0 keeps ordinary chart rendering in the format entries and provides 3-D and Region Map rendering as optional renderer modules. The same injected renderer works for charts hosted by Word, Excel and PowerPoint.',
          'Existing viewers continue to use the established chart fallback without configuration changes. No migration is required.',
        ],
        bullets: [
          'Authored 3-D chart views render through one model-space camera and projected mesh pipeline.',
          'Country-level ChartEx Region Maps render offline from worksheet or document data.',
          'Shared axis planning, titles, data labels, legend paint and ChartEx layout now follow more authored OOXML properties.',
          'Both renderers use separate entries and are loaded only when supplied. They launched for the main thread only in v0.79.0; current releases also reconstruct the built-ins inside render workers.',
        ],
      },
      {
        title: 'One 3-D scene for axes and data',
        modules: ['@silurus/ooxml/three-d', '@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Axes, walls, grids, bars, lines and surfaces must share one projection to remain geometrically coherent.',
        paragraphs: [
          'The 3-D renderer projects chart walls, axes and data through one homogeneous camera instead of applying unrelated screen-space offsets. Bar and column solids use projected meshes, area charts use extruded surfaces, and pie slices use bounded mesh geometry. Authored box, cylinder, cone, cone-to-max, pyramid and pyramid-to-max shapes are retained where the OOXML model supplies them.',
          'The renderer honors the view saved in the document, including rotation, perspective, right-angle axes, depth and height.',
        ],
      },
      {
        title: 'Offline country Region Maps',
        modules: ['@silurus/ooxml/region-map', '@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'An OOXML Region Map may store country names and values without embedding the geographic polygons used by Excel.',
        paragraphs: [
          'The parser keeps ChartEx country identities, numeric color values, authored projections and two- or three-stop color scales. The renderer module renders supported countries offline.',
          'The map uses public-domain Natural Earth Admin 0 Countries 1:110m geometry.',
          'v0.79.0 supports country-level world maps. Cached provider identities and state, county or postal views are not supported.',
        ],
      },
      {
        title: 'Enable the optional renderers',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/three-d', '@silurus/ooxml/region-map'],
        rationale: 'Dependency injection keeps the base entries small and gives every host format the same renderer contract.',
        paragraphs: [
          'Import each renderer from its separate package entry and pass it to a Viewer or headless engine at construction/load time. The same built-in renderer injection works in main and worker modes; custom renderer objects retain the documented worker fallback.',
        ],
        examples: [
          {
            title: 'Enable advanced charts in an XLSX Viewer',
            code: `import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { threeD } from '@silurus/ooxml/three-d';
import { regionMap } from '@silurus/ooxml/region-map';

const viewer = new XlsxViewer(container, {
  mode: 'worker',
  threeD,
  regionMap,
});

await viewer.load(source);`,
          },
        ],
      },
      {
        title: 'Chart fidelity and bounded work',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Office chart fidelity must not reintroduce unbounded tick, text, hierarchy or mesh expansion.',
        paragraphs: [
          'Classic and ChartEx charts now share more of the same value-axis planner, title defaults, data-label layout, rich text, legend paint and explicit-property precedence. This improves waterfall, funnel, box-and-whisker, histogram, Pareto, treemap, sunburst, bubble, line, area and combo charts without family-specific copies of the same policy.',
          'Parser/model inputs and painting work are bounded for ticks, data-label text, hierarchies, Region Map rows and expanded 3-D primitives. The optional entries keep 3-D and map geometry out of the base format bundles.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No existing option is removed or renamed. Upgrade normally to receive the shared 2-D chart fixes. Add the threeD or regionMap option only when the corresponding authored chart needs its optional renderer.',
          'The optional entries are ESM-only like the rest of the package. Custom renderer objects retain the documented fallback in worker mode.',
        ],
      },
    ],
  },
  {
    slug: 'v0781-docx-pagination-fix',
    date: '2026-08-12',
    label: 'Release note',
    version: 'v0.78.1',
    title: 'DOCX pagination fix in v0.78.1',
    summary: 'v0.78.1 fixes a regression that could move text in narrow Word table cells to the wrong page or hide intentional blank spacing.',
    audience: 'Users who display Word documents with narrow table columns or full-width spacing. No application changes are required when upgrading from v0.78.0.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This patch restores Word-compatible page breaks for affected table content and preserves intentional blank space created with full-width spaces.',
          'There are no public API changes and no migration is required.',
        ],
      },
    ],
  },
  {
    slug: 'v078-chart-fidelity-and-multi-selection',
    date: '2026-08-12',
    label: 'Release note',
    version: 'v0.78.0',
    title: 'Better charts and XLSX multi-selection in v0.78.0',
    summary: 'Charts in DOCX, XLSX and PPTX render more like Office, and XLSX sheets now support selecting multiple separate areas.',
    audience: 'Users who view charts in Office files or work with interactive worksheets. No application changes are required when upgrading from v0.77.1.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This release improves how charts look across Word, Excel and PowerPoint files and adds multiple-area selection to worksheets.',
        ],
        bullets: [
          'Chart axes, backgrounds, fills, labels, legends and stacked values more closely match Office.',
          'Waterfall, box-and-whisker, treemap and bubble charts have more accurate layout and styling.',
          'Hold Ctrl on Windows/Linux or Command on macOS to select multiple worksheet areas.',
        ],
      },
      {
        title: 'Charts look closer to Office',
        paragraphs: [
          'Charts now follow more of the formatting saved by Office, including axis ranges and ticks, chart and plot backgrounds, wrapped category labels, data-label placement, legend keys, pattern fills and percentage stacking.',
          'Modern chart types also receive substantial visual improvements. Bubble sizes, treemap hierarchy, waterfall styling and box-and-whisker geometry now more closely match their Office appearance.',
        ],
      },
      {
        title: 'Select multiple worksheet areas',
        paragraphs: [
          'Hold Ctrl on Windows/Linux or Command on macOS while dragging to add another cell, row or column area to the selection.',
          'Each selected area uses the configured selection color, while the active cell remains easy to identify. Adjacent areas no longer create doubled borders.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No migration is required. Existing viewer setup and single-area worksheet selection continue to work as before.',
        ],
      },
    ],
  },
  {
    slug: 'v0771-rendering-fidelity',
    date: '2026-08-11',
    label: 'Release note',
    version: 'v0.77.1',
    title: 'DrawingML and PowerPoint fidelity in v0.77.1',
    summary: 'v0.77.1 improves shared DrawingML geometry, theme styles and gradients, and fixes PowerPoint effects, charts, tables and text without changing the v0.77 API.',
    audience: 'Applications that display DOCX, XLSX or PPTX files. No source changes are required when upgrading from v0.77.0.',
    sections: [
      {
        title: 'In short',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/node', 'office-open-xml-viewer (VS Code extension)'],
        rationale: 'DrawingML concepts shared by Word, Excel and PowerPoint should be parsed once and retain the information each renderer needs.',
        kind: 'summary',
        paragraphs: [
          'This patch is compatible with v0.77.0. It changes rendering and parser fidelity only; it does not rename or remove public APIs.',
          'The largest improvements are visible in presentations that rely on theme fills, custom geometry, table and chart defaults, shadows, 3-D effects or reflected text. Shared DrawingML line and gradient fixes also benefit Word and Excel documents that use the same OOXML features.',
        ],
        bullets: [
          'DOCX, XLSX and PPTX now share custom-geometry guide evaluation and complete theme line and gradient models.',
          'PowerPoint theme backgrounds, image fills, chart defaults, table styles, bullets and terminal whitespace follow their authored or inherited values more closely.',
          'Shadows and related effects use the complete painted silhouette, including protruding callout leaders, strokes and line ends.',
          'Text reflections stay sharp where they meet the source and become blurrier with distance while keeping bounded rendering work.',
        ],
      },
      {
        title: 'Shared DrawingML improvements',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/node'],
        rationale: 'Custom geometry, themes, line styles and gradients are DrawingML features used by all three Office formats.',
        paragraphs: [
          'The parsers now evaluate custom-geometry formulas and quadratic Bézier paths through one shared implementation, including the standard fallback when a path omits its own coordinate size.',
          'Theme fill and line recipes retain their colors, dashes, joins, line ends and gradient geometry. Chart color-map overrides and XLSX theme relationships are resolved before rendering instead of being replaced with a flat fallback color.',
          'Theme XML retention, gradient tiles and projected effect surfaces remain bounded so these fidelity improvements do not remove the existing protections for large or hostile documents.',
        ],
      },
      {
        title: 'PowerPoint rendering fixes',
        modules: ['@silurus/ooxml/pptx', '@silurus/ooxml/node', 'office-open-xml-viewer (VS Code extension)'],
        rationale: 'PowerPoint builds the final appearance by combining inherited theme, layout and local properties before applying effects to the complete painted object.',
        paragraphs: [
          'Presentation backgrounds and shape fills now preserve theme images and PowerPoint color transforms. Hidden layout graphics stay hidden, while charts and tables inherit the intended axes, labels, banding, text, borders and fills.',
          'Outer shadows, reflections, soft edges and 3-D effects are composited after the full fill, stroke, callout and arrowhead silhouette is known. This prevents effects from disappearing, being clipped, or being applied separately to only part of an object.',
          'Text layout retains authored bullet and auto-number formatting, ignores paragraph-ending whitespace that should not create another line, and renders floor reflections with a legible contact edge and increasing blur farther from the text.',
        ],
      },
      {
        title: 'Compatibility and verification',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'A patch release must improve rendering without requiring application code changes.',
        paragraphs: [
          'The v0.77 public API remains unchanged. Existing Viewer, document, selection and MCP integrations continue to use the same calls and types.',
          'The renderer changes were checked with parser and TypeScript test suites, public API checks, focused PowerPoint/PDF comparisons, independent adversarial reviews, and the complete private DOCX/XLSX/PPTX visual regression corpus using v0.77.0 as the previous-renderer baseline.',
        ],
      },
    ],
  },
  {
    slug: 'v077-migration-guide',
    date: '2026-08-10',
    label: 'Release note',
    version: 'v0.77',
    title: 'Migrating to v0.77',
    summary: 'v0.77 updates selection APIs, adds selectable document elements, simplifies OOXML MCP tools, and makes errors from asynchronous Viewer methods easier to handle.',
    audience: 'Applications that control XLSX selection, consume DOCX/XLSX/PPTX selection context, use the OOXML MCP server, or configure Viewer onError callbacks.',
    sections: [
      {
        title: 'In short',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/node', 'office-open-xml-viewer (VS Code extension)', 'ooxml-mcp-server'],
        rationale: 'v0.77 makes one coherent selection and error contract available across the browser library, VS Code extension and MCP server.',
        kind: 'summary',
        paragraphs: [
          'Most applications that only load and display documents do not need source changes. Review the sections below if your code uses XLSX selection, selection context, removed option or type names, OOXML MCP tools, or Viewer onError callbacks.',
          'v0.77 deliberately removes short-lived compatibility surfaces instead of maintaining two APIs for the same operation. Update the affected calls before upgrading; there are no temporary aliases for the removed names.',
        ],
        bullets: [
          'XLSX selection: replace select(), selection and onSelectionChange with setSelection(), selectionState and onSelectionStateChange.',
          'Selection context: check context.kind before reading text, cell-range or element details.',
          'Element selection: opt in with enableElementSelection to select charts, pictures and shapes and show their non-editable outline.',
          'Context menus: use onContextMenu and await getContext(); originalEvent remains available synchronously for preventDefault().',
          'MCP: use the consolidated active-context and replacement format tools; removed subset tools are not retained as aliases.',
          'API cleanup: remove non-functional options and exact aliases, and stop passing load-only or render-only options to APIs that cannot use them.',
          'Errors: awaitable Viewer methods reject; onError is reserved for Viewer-managed background work with no Promise to await.',
        ],
      },
      {
        title: 'Why these changes ship together',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', 'office-open-xml-viewer (VS Code extension)', 'ooxml-mcp-server'],
        rationale: 'The Viewer keeps the current selection, while integrations receive a size-limited copy that they can inspect without editing the document.',
        paragraphs: [
          'The old XLSX API described only one range and could not preserve all Excel selection details. DOCX and PPTX also lacked a consistent way to pass selected text or objects to another application. v0.77 gives all three formats explicit selection and a size-limited copy of the selected content.',
          'The MCP and error-handling changes follow the same approach: each task now has one supported API, and asynchronous methods report errors through their returned Promise.',
        ],
      },
      {
        title: 'Replace the XLSX selection compatibility API',
        modules: ['@silurus/ooxml/xlsx'],
        rationale: 'The removed range-only model could not represent multiple selected areas, ActiveCell and the Shift-extension anchor as separate Excel concepts.',
        paragraphs: [
          'The previous select(), selection, onSelectionChange, CellRange and SelectionMode exports are removed. They could not represent Excel’s separate selected areas, ActiveCell and Shift-extension anchor without overloading one range value.',
          'Use setSelection(), selectionState, onSelectionStateChange, XlsxSelectionState and XlsxSelectionArea. A1 strings remain accepted by setSelection() for the common single-range case.',
          'Replace CellRange values with XlsxSelectionState, and describe each selected region with XlsxSelectionArea. If you constructed SelectionMode values directly, the old cols and all modes correspond to the new columns and sheet area kinds.',
        ],
        examples: [
          {
            title: 'Before',
            code: `const viewer = new XlsxViewer(container, {
  onSelectionChange(range) {
    updateSelection(range);
  },
});

viewer.select('B2:D6');`,
          },
          {
            title: 'After',
            code: `const viewer = new XlsxViewer(container, {
  onSelectionStateChange(state) {
    updateSelection(state);
  },
});

viewer.setSelection('B2:D6');`,
          },
        ],
      },
      {
        title: 'Check context.kind before reading selection details',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Selection details differ for text, cells and elements, so code must first check which kind it received.',
        paragraphs: [
          'The DocxSelectionContext and XlsxSelectionContext types can now contain selected text, XLSX cells, or an element such as a chart or picture. Check context.kind first so TypeScript knows which details are available.',
          'The native DOCX helper readDocxSelectionContext() is renamed to readDocxTextSelectionContext() because it reads only a DOM text selection. Use DocxViewer.getSelectionContext() for the text-or-element union.',
          'PptxElementSelectionContext is renamed to PptxElementContext. The data shape is unchanged, and the new name also fits direct point queries where no Viewer selection exists.',
        ],
        examples: [
          {
            title: 'Read a cross-format context safely',
            code: `const context = viewer.getSelectionContext();

if (context?.kind === 'text') {
  consumeText(context.text);
} else if (context?.kind === 'range') {
  consumeCells(context.cells);
} else if (context?.kind === 'element') {
  consumeElement(context.elementType);
}`,
          },
        ],
      },
      {
        title: 'Use element selection and context menus explicitly',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Element selection changes what a click does and requires additional processing, so applications enable it explicitly. Context-menu target details may take a moment to obtain.',
        paragraphs: [
          'enableElementSelection defaults to false. When enabled, clicking a supported chart, picture or shape selects it and draws a non-editable outline.',
          'onContextMenu provides originalEvent immediately and getContext() for the clicked content. getContext() returns a Promise; call originalEvent.preventDefault() before awaiting it when replacing the browser menu.',
        ],
        examples: [
          {
            title: 'Build a host-owned context menu',
            code: `const viewer = new DocxViewer(canvas, {
  enableElementSelection: true,
  onContextMenu: async ({ originalEvent, getContext }) => {
    originalEvent.preventDefault();

    try {
      showContextMenu(await getContext());
    } catch (error) {
      showContextError(error);
    }
  },
});`,
          },
        ],
      },
      {
        title: 'Replace removed MCP subset tools',
        modules: ['ooxml-mcp-server', 'office-open-xml-viewer (VS Code extension)'],
        rationale: 'Keeping two tools for the same read operation makes agent routing less reliable and doubles the contract surface without adding capability.',
        paragraphs: [
          'Normal VS Code use does not require a manual change. Update only custom prompts, tool allowlists, or other integrations that refer to one of the removed tool names.',
          'Use ooxml_get_active_context for the active OOXML preview and its current text, range or element context. This replaces the earlier active-selection tool name and keeps one routing entry point for all three formats.',
        ],
        bullets: [
          'Replace xlsx_get_sheet_names with xlsx_parse.',
          'Replace docx_get_paragraph with docx_get_body_element.',
          'Replace pptx_get_shape and pptx_get_shape_text with pptx_get_element.',
          'Replace ooxml_get_active_selection with ooxml_get_active_context.',
        ],
      },
      {
        title: 'Remove unused rendering options',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx', '@silurus/ooxml/node'],
        rationale: 'These options were accepted by TypeScript but could not affect the result of the operation on which they appeared.',
        paragraphs: [
          'DOCX showTrackChanges is removed from browser and Node render options because the retained paint pipeline never consulted it; true and false produced identical pixels. Tracked revisions continue to render as they do today, but the library no longer advertises a non-functional Final / No Markup switch.',
          'Borrowed Viewer factories no longer accept load-only mode because they use the mode of the document that is already loaded. DocxDocument.collectPageRuns() now accepts only width and currentDate. PptxPresentation.presentSlide() no longer accepts skipMediaControls.',
          'XlsxWorkbook.renderViewport() no longer accepts fetchImage or loadedImages because the workbook owns its image loading and cache.',
        ],
      },
      {
        title: 'Use public rendering option types',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx'],
        rationale: 'Application code should use types named for public operations, not the internal messages exchanged with a worker.',
        paragraphs: [
          'Use XlsxRenderViewportOptions with XlsxWorkbook.renderViewport() and RenderViewportToBitmapOptions with renderViewportToBitmap(). Use CollectPageRunsOptions with DocxDocument.collectPageRuns().',
          'WireRenderPageOptions, WireRenderViewportOptions and WireSizeOverrides were internal worker-message types and are no longer exported.',
        ],
      },
      {
        title: 'Rename exported type aliases',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'The removed aliases duplicated existing public types or used names that no longer matched the error stages they represented.',
        paragraphs: [
          'Replace XlsxChartSeries, SeriesDataLabels, DataLabelOverride, DataPointOverride, ErrBars and ManualLayout with ChartSeries, ChartSeriesDataLabels, ChartDataLabelOverride, ChartDataPointOverride, ChartErrBars and ChartManualLayout.',
          'Replace OoxmlErrorSource with OoxmlErrorStage. Rename old stage values as follows: zip-part → decompression, parser → parsing, serializer → serialization and renderer → rendering.',
        ],
      },
      {
        title: 'Password-protected files now load through Viewers',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Self-loading Viewers now pass the documented password option to their document engine.',
        paragraphs: [
          'No source change is needed. Existing Viewer code that supplies LoadOptions.password now works as the public API already promised.',
        ],
      },
      {
        title: 'Catch every awaitable Viewer failure',
        modules: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
        rationale: 'Errors from a method the application can await are reported by rejecting that Promise. onError remains for later background work.',
        paragraphs: [
          'Viewer load() and onError no longer form alternative completion channels. load(), navigation and every other public method that returns a Promise reject that Promise on failure, even when onError is configured. The same failure is never delivered twice.',
          'Keep onError only when the application needs failures from Viewer-managed work that has no Promise to await, such as later virtualized rendering or embedded-media playback.',
        ],
        examples: [
          {
            title: 'Separate awaited and background failures',
            code: `const viewer = new DocxViewer(canvas, {
  onError(error) {
    reportBackgroundFailure(error);
  },
});

try {
  await viewer.load(file);
} catch (error) {
  showLoadFailure(error);
}`,
          },
        ],
      },
    ],
  },
  {
    slug: 'v076-migration-guide',
    date: '2026-08-06',
    label: 'Release note',
    version: 'v0.76',
    title: 'Migrating to v0.76',
    summary: 'v0.76 makes shared-engine Viewer construction explicit, adds the canvas-mounted XLSX sheet Viewer, and replaces the synchronous Node parser compatibility APIs with one owned asynchronous pipeline.',
    audience: 'Applications that share one parsed document across multiple Viewers, use the Node parser helpers, or want to render individual XLSX sheets into caller-owned canvases.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'Ordinary browser Viewer code that constructs a Viewer and awaits load(source) does not change. The migration applies to shared-engine construction and the Node parser helpers.',
        ],
        bullets: [
          'Shared browser engine: replace the document, presentation or workbook constructor option with the matching named factory.',
          'Node parser helpers: replace synchronous parse and extraction exports with an asynchronous materializer or owned session.',
          'XLSX unit rendering: use XlsxSheetViewer for a caller-owned canvas without workbook tabs or footer controls.',
          'Archive policy: maxArchiveEntries is now available in Browser and Node resourceLimits.',
        ],
      },
      {
        title: 'Use a named factory for a shared engine',
        paragraphs: [
          'The document, presentation and workbook Viewer options are removed. Load the engine once, then use fromDocument(), fromPresentation() or fromWorkbook(). The factory is synchronous because the engine is already loaded; rendering and navigation remain asynchronous.',
          'Load-only settings such as mode, wasmUrl, resourceLimits, password and useGoogleFonts belong on the engine load call. A Viewer created by a factory cannot load another source.',
          'This cleanup obligation is not new: the removed constructor-option injection also borrowed its engine, so viewer.destroy() intentionally left that engine open. Destroy the borrowed Viewers before destroying their caller-owned engine once.',
        ],
        examples: [
          {
            title: 'Before: constructor-option injection',
            code: `const document = await DocxDocument.load(file);

const viewer = new DocxViewer(canvas, {
  document,
});`,
          },
          {
            title: 'After: explicit borrowed-engine factory',
            code: `const document = await DocxDocument.load(file, {
  mode: 'worker',
});

const viewer = DocxViewer.fromDocument(canvas, document);
await viewer.goToPage(0);

viewer.destroy();
document.destroy();`,
          },
        ],
        bullets: [
          'DOCX: DocxViewer.fromDocument() and DocxScrollViewer.fromDocument().',
          'PPTX: PptxViewer.fromPresentation() and PptxScrollViewer.fromPresentation().',
          'XLSX: XlsxViewer.fromWorkbook() and XlsxSheetViewer.fromWorkbook().',
        ],
      },
      {
        title: 'Replace synchronous Node parser helpers',
        paragraphs: [
          'The old synchronous exports could not provide the same cancellation, limits, metrics, archive reuse and deterministic cleanup as the asynchronous parser pipeline. v0.76 removes that compatibility path instead of maintaining two subtly different implementations.',
          'Use a materializer when the application needs a complete caller-owned model. Use an owned session for bounded sequential work and close it in finally. There is intentionally no synchronous wrapper around the new APIs.',
          'In v0.75, parseXlsx() returned ParsedWorkbook: workbook metadata and sheet list, styles, and shared strings. It did not include worksheet cell rows; those required parseXlsxSheet() or parseXlsxAllSheets(). The v0.76 names make those three materialization scopes explicit.',
          'session.workbookIndex is the already-parsed, read-only ParsedWorkbook property on a session returned by openXlsxWorkbook(); it is useful when the same session will also stream worksheetRows(). It is not a second direct replacement. If only the old parseXlsx() result is needed, use materializeXlsxWorkbookIndex().',
        ],
        bullets: [
          'parseDocx() → await materializeDocxDocument().',
          'parsePptx() → await materializePptxPresentation().',
          'parseXlsx() → await materializeXlsxWorkbookIndex() for metadata/index only.',
          'parseXlsxSheet() → await materializeXlsxWorksheet() for one caller-owned worksheet.',
          'parseXlsxAllSheets() → await materializeXlsxWorkbook() for the index and every worksheet; this has the highest time and retained-memory cost.',
          'PPTX image and media extraction → await session.getImage() or session.getMedia() on openPptxPresentation().',
        ],
        examples: [
          {
            title: 'Owned session',
            code: `const presentation = await openPptxPresentation(bytes);

try {
  for await (const slide of presentation.slides()) {
    consume(slide);
  }
} finally {
  await presentation.close();
}`,
          },
        ],
      },
      {
        title: 'Render one XLSX sheet into a canvas',
        paragraphs: [
          'XlsxSheetViewer mounts one active worksheet viewport into a caller-owned canvas and includes sheet navigation, logical viewport, selection, search and zoom APIs without workbook tabs, footer controls or a native scrollbar.',
          'This is an XLSX-specific Viewer boundary. A worksheet is not treated as equivalent to a DOCX page or PPTX slide; each format keeps the Viewer split that matches its own document model.',
          'fromWorkbook() does not materialize an arbitrary first worksheet. The first goToSheet(index) materializes only the requested sheet, which keeps parse-once multi-window integrations efficient. The full XlsxViewer still starts its initial sheet display immediately.',
        ],
        examples: [
          {
            title: 'Borrow one workbook across sheet Viewers',
            code: `const workbook = await XlsxWorkbook.load(file);
const sheet = XlsxSheetViewer.fromWorkbook(canvas, workbook);

await sheet.goToSheet(2);
await sheet.setViewportOffset({ x: 120, y: 80 });

sheet.destroy();
workbook.destroy();`,
          },
        ],
      },
      {
        title: 'Bound the number of archive entries',
        paragraphs: [
          'Browser and Node load or session options now accept resourceLimits.maxArchiveEntries. Omission uses the calibrated default, null disables that configurable limit, and the internal hard ceiling remains enforced.',
        ],
      },
    ],
  },
  {
    slug: 'v075-resource-governance',
    date: '2026-08-02',
    label: 'Release note',
    version: 'v0.75',
    title: 'Resource limits, typed failures and metrics for large files',
    summary: 'v0.75 applies default inflated-package limits to every DOCX, XLSX and PPTX load, reports measured limit failures with typed errors, and exposes content-free usage metrics.',
    audience: 'Applications that load user-supplied DOCX, XLSX or PPTX files, especially those that customize maxZipEntryBytes or accept unusually large files.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'Most applications do not need to change how they construct a Viewer: omitting resourceLimits selects the standard policy. Applications should, however, treat a typed limit error as an intentional refusal to preview the file rather than as an unknown renderer failure.',
        ],
        bullets: [
          'No resourceLimits today: no configuration change is required. v0.75 supplies the defaults.',
          'User-supplied files: catch OoxmlResourceLimitError and OoxmlDecodedImageLimitError and show a clear “too large to preview” result.',
          'Using maxZipEntryBytes: it remains a deprecated compatibility alias, but migrate to resourceLimits.maxArchiveEntryBytes.',
          'Need different limits: collect OoxmlResourceMetrics from representative production files before choosing values.',
        ],
      },
      {
        title: 'What is now bounded',
        paragraphs: [
          'DOCX, XLSX and PPTX now use the same admission policy while opening and lazily reading the ZIP package. The limits apply to inflated package parts—not to the compressed upload size and not to JavaScript heap usage. A part is charged by the largest amount read from it, so reading the same part again does not consume the distinct-total budget twice.',
          'Raster image decoding has separate, non-configurable browser guards. These are hard implementation ceilings because decoded memory and browser/GPU overhead do not map consistently to an application-supplied byte value across devices.',
        ],
        bullets: [
          '128 MiB for any one inflated XML, image, media or other package part by default.',
          '256 MiB across distinct inflated parts read during one package session by default.',
          '32 megapixels per raster image, 128 MiB aggregate decoded raster ownership, and two concurrent image decodes.',
          'Internal hard ceilings still apply when either configurable package limit is set to null.',
        ],
      },
      {
        title: 'Handle an intentional rejection',
        paragraphs: [
          'Catch limit errors wherever your application awaits load or later lazy document work. OoxmlResourceLimitError includes the measured limit and observed value in details.violation; OoxmlDecodedImageLimitError identifies a raster-image guard. The same classes are re-exported by the DOCX, XLSX and PPTX entry points.',
        ],
        examples: [
          {
            title: 'Show a specific preview error',
            code: `import {
  DocxViewer,
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
} from '@silurus/ooxml/docx';

const viewer = new DocxViewer(canvas);

try {
  await viewer.load(file);
} catch (error) {
  if (error instanceof OoxmlResourceLimitError) {
    const { limit, observed } = error.details.violation;
    showPreviewError(
      \`This file exceeds the preview limit (\${observed} of \${limit} bytes).\`,
    );
    return;
  }

  if (error instanceof OoxmlDecodedImageLimitError) {
    showPreviewError('This file contains an image that is too large to preview.');
    return;
  }

  throw error;
}`,
          },
        ],
      },
      {
        title: 'Choose limits from observed files',
        paragraphs: [
          'Start with the defaults rather than guessing. onResourceMetrics receives a content-free report when the initial load settles, including failed loads. After a successful load and any lazy page, sheet, slide, image or media access, getResourceMetrics() returns a fresh snapshot. The library does not transmit or persist either report.',
          'Metrics exclude filenames, URLs, package paths, document text, passwords and raw error messages. Sizes, counts and timings are still document-derived metadata, so collect them only under your application’s consent and retention policy. debug: true is a separate development aid that prints the same class of data to the console.',
        ],
        examples: [
          {
            title: 'Collect metrics without console output',
            code: `const viewer = new DocxViewer(canvas, {
  onResourceMetrics(metrics) {
    usageMetrics.record(metrics);
  },
});

await viewer.load(file);

// Includes package work observed after the initial load.
usageMetrics.record(await viewer.getResourceMetrics());`,
          },
          {
            title: 'Apply values chosen from your own data',
            code: `const MiB = 1024 * 1024;

const viewer = new DocxViewer(canvas, {
  resourceLimits: {
    maxArchiveEntryBytes: 64 * MiB,
    maxTotalInflatedBytes: 192 * MiB,
  },
});`,
          },
        ],
      },
      {
        title: 'Migrate maxZipEntryBytes',
        paragraphs: [
          'A positive maxZipEntryBytes value keeps its existing per-entry meaning in v0.75, so this migration is not required immediately. New code should use resourceLimits. Do not supply conflicting values through both options; that is rejected before parsing begins.',
        ],
        examples: [
          {
            title: 'Before',
            code: `const viewer = new DocxViewer(canvas, {
  maxZipEntryBytes: 64 * 1024 * 1024,
});`,
          },
          {
            title: 'After',
            code: `const viewer = new DocxViewer(canvas, {
  resourceLimits: {
    maxArchiveEntryBytes: 64 * 1024 * 1024,
  },
});`,
          },
        ],
      },
      {
        title: 'What these limits cannot guarantee',
        paragraphs: [
          'Package counters do not measure peak process memory. XML trees, document models, canvas backing stores, decoded images, renderer state and browser-managed memory can require several times the measured inflated bytes. The defaults reject known measurable hazards earlier, but cannot guarantee that every browser and device will avoid an out-of-memory termination.',
          'A residual WebAssembly trap is reported conservatively as parser-crashed, not parser-oom. At the current WASM boundary, Rust panic, allocation failure, stack overflow and explicit unreachable can converge on the same WebAssembly.RuntimeError, so the original cause cannot be recovered reliably after the trap. Worker mode can keep parser and renderer work away from the Window and improve failure containment, but a Worker is not a separate operating-system process or a strict memory sandbox.',
        ],
      },
    ],
  },
];

export const latestAnnouncements = announcements.slice(0, 3);

export function formatAnnouncementDate(value: string): string {
  return new Intl.DateTimeFormat('en', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  }).format(new Date(`${value}T00:00:00Z`));
}
