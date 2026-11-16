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
  /** Site-internal guide pages for the section, rendered after its text. */
  readonly links?: readonly AnnouncementLink[];
}

export interface AnnouncementLink {
  readonly label: string;
  /** Site-relative path, such as `/fonts`. */
  readonly href: `/${string}`;
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
    slug: 'v090-opendocument-and-font-sources',
    date: '2026-11-16',
    label: 'Release note',
    version: 'v0.90.0',
    title: 'Font sources and OpenDocument reading in v0.90.0',
    summary: 'v0.90.0 adds the fonts option, which sets where web fonts are loaded from, and optional, experimental reading of ODP presentations, ODT documents and ODS spreadsheets. It renders Symbol and Wingdings text on every host and reduces the work in the font loading path when web fonts are enabled. The useGoogleFonts option is deprecated in favour of fonts: [googleFonts()].',
    audience: 'All Viewer and Node integrations. Applications that pass useGoogleFonts, load web fonts under a Content Security Policy or compile against the low-level model types should read Upgrading.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This release adds application-configured font sources and an optional add-on that reads OpenDocument files. By default, no font is requested from any network host, and ordinary OOXML viewing loads no OpenDocument code.',
        ],
        bullets: [
          'The `fonts` option loads web fonts from Google Fonts, a mirror or proxy, a hosted font service, your own files or your own resolver, for DOCX, XLSX and PPTX in main and worker mode.',
          'The experimental `@silurus/ooxml-opendocument` add-on opens .odp, .odt and .ods files in the PPTX, DOCX and XLSX viewers.',
          'Symbol, Wingdings, Wingdings 2, Wingdings 3 and Webdings text renders on every host instead of tofu or Latin letters.',
          'With web fonts enabled, font loading does less work before first display; large XLSX sheets repaint without rebuilding their render index on every frame.',
          '`useGoogleFonts` is deprecated: `useGoogleFonts: true` becomes `fonts: [googleFonts()]`. Under a Content Security Policy, the font file host must be allowed in `connect-src` as well as `font-src`.',
        ],
      },
      {
        title: 'OpenDocument add-on',
        modules: ['@silurus/ooxml-opendocument/odp', '@silurus/ooxml-opendocument/odt', '@silurus/ooxml-opendocument/ods'],
        rationale: 'Reading OpenDocument files into the existing models lets applications show them with the same viewers, without a conversion service.',
        paragraphs: [
          'The experimental `@silurus/ooxml-opendocument` add-on reads OpenDocument presentations, text documents and spreadsheets directly into the PPTX slide, DOCX document and XLSX workbook models; the ordinary renderers draw them, and no intermediate OOXML package is generated. Pass `odp` to the PPTX viewers and loaders, `odt` to the DOCX ones and `ods` to the XLSX ones, in the browser and in Node sessions.',
          'The readers aim to render every feature they read, as ODF 1.3 specifies, with LibreOffice as the reference where the specification leaves a choice. Where the target model has no exact counterpart, a feature is drawn with its nearest equivalent; a few, listed as known differences, are not drawn. Only corrupt, encrypted or unsafe packages, or packages over a resource limit, are rejected. The readers never run scripts, macros or forms and never fetch external resources.',
          'ODP covers master pages, the style cascade, custom-shape geometry, gradient and bitmap fills, text, pictures, tables, speaker notes and charts. Importing a reader fetches nothing; its source module and WASM load only when a matching file is opened. For custom asset URLs, create one with `createOdp({ wasmUrl, moduleUrl })` (likewise `createOdt` and `createOds`).',
        ],
        examples: [
          {
            title: 'ODP presentations in the PPTX viewer',
            code: `import { PptxViewer } from '@silurus/ooxml/pptx';
import { odp } from '@silurus/ooxml-opendocument/odp';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { odp });
await viewer.load(await file.arrayBuffer()); // .odp, .otp or .pptx`,
          },
        ],
        links: [{ label: 'OpenDocument files guide', href: '/opendocument' }],
      },
      {
        title: 'Font sources',
        modules: ['DOCX', 'XLSX', 'PPTX'],
        rationale: 'Applications differ in where they may fetch fonts from: a public CDN, a regional mirror, an internal proxy, a licensed font service or their own servers. The fonts option lets the application name its sources.',
        paragraphs: [
          'The new `fonts` load option takes an ordered list of sources built with `googleFonts()`, `fontCssApi()`, `fontCss()`, `fontFiles()` and `fontResolver()`. `googleFonts({ baseUrl })` sends the stylesheet requests to a Google-compatible mirror or proxy, and `googleFonts({ additionalFamilies })` adds Google families outside the preset catalog. `fontCssApi()` covers other per-family CSS APIs, `fontCss()` a static kit stylesheet, `fontFiles()` your own files or bytes, and `fontResolver()` authenticated or private services through a function that runs in the page.',
          'Sources are consulted per family, in order; the first source that has a face claims the family, and a failure keeps the host font with a `console.warn` instead of rejecting the load. Every source takes `substitutions`, such as the exported `officeFontSubstitutions` (Calibri to Carlito, Cambria to Caladea and others). A `FontSource` is opaque, so create every entry with a factory. Font sources have fixed limits; a factory input beyond one throws the new `OoxmlError` code `font-source-limit`.',
          'DOCX and XLSX now load only the `unicode-range` subsets that the document paints, as PPTX already did. DOCX also lays out a served face with Word\'s line rule from the face\'s own OpenType tables, so served fonts no longer set lines too tight. To read those tables, the library fetches the font files itself.',
        ],
        examples: [
          {
            title: 'A private resolver, then an internal Google-compatible mirror',
            code: `import { fontResolver, googleFonts, officeFontSubstitutions } from '@silurus/ooxml';

const privateFonts = fontResolver(async (demand, { signal }) => {
  const response = await fetch('/api/fonts/resolve', {
    method: 'POST',
    body: JSON.stringify({ authoredFamilies: demand.authoredFamilies, fallbackFamilies: demand.fallbackFamilies }),
    signal,
  });
  return response.json(); // FontFaceDeclaration[]
}, { substitutions: officeFontSubstitutions });

const viewer = new DocxViewer(canvas, {
  fonts: [privateFonts, googleFonts({ baseUrl: 'https://fonts.internal.example' })],
});`,
          },
        ],
        links: [{ label: 'Font sources guide', href: '/fonts' }],
      },
      {
        title: 'Symbol font rendering',
        modules: ['DOCX', 'XLSX', 'PPTX'],
        paragraphs: [
          'Text, `w:sym` and `a:sym` symbols and list markers in Symbol, Wingdings, Wingdings 2, Wingdings 3 and Webdings now render on every host. Greek letters and maths operators of Symbol, and the pictographs of the Wingdings and Webdings fonts, used to draw as tofu or as Latin letters wherever the font was missing or, like Apple\'s Symbol, not symbol-encoded.',
          'Each code is checked against the face that actually paints it: a face that draws the code keeps its own glyph, as on Windows, and any other code becomes its Unicode character. A converted pictograph takes the text colour and is never drawn as a colour emoji: where only a colour emoji face has it, the authored code is kept. In XLSX this covers cell values, number-format output, rich-text runs, text boxes and slicer labels, and overflow, wrapping and automatic row heights follow the symbols as drawn.',
        ],
      },
      {
        title: 'Font loading and XLSX repaint performance',
        modules: ['DOCX', 'XLSX', 'PPTX'],
        paragraphs: [
          'With font sources enabled, more content now has a precise font demand: accented and other composed text, right-to-left text, ruby, vertical Word sections and table cells, most charts, raster and SVG pictures, and numbered or vertical PowerPoint text no longer make a document fetch every `unicode-range` subset of its families. Some content still does, for example metafile pictures, equations, region-map charts and Word vertical text boxes.',
          'The font load path does less work on the main thread. Faces are built as their downloads land instead of after a whole batch; `googleFonts()` on its default origin downloads files with a wider parallel pool and, in pages without a service worker, decodes files from their downloaded bytes on the main thread only for the first MiB of a load and only while that has taken at most half of the load\'s time so far, registering the other files from their URLs so the browser decodes them off the main thread; `unicode-range` lists are parsed in one pass and families sliced alike share their range work; and the font-source code is imported while the worker parses instead of afterwards. The PowerPoint font-demand preflight no longer repeats work and keeps the opening slide it already transferred, and the XLSX workbook demand skips text the renderer already covers. Rendering is unchanged.',
          'XLSX scrolling and repainting no longer rebuild the sheet render index on every frame, which made up most of each frame on very large sheets.',
        ],
      },
      {
        title: 'Other fixes and additions',
        modules: ['DOCX', 'XLSX', 'PPTX', 'Node'],
        paragraphs: [
          'PowerPoint table cells now draw their diagonal borders, and end-of-paragraph run properties apply to stacked vertical text. XLSX inline-string cells now show their furigana when phonetic display is on. A DOCX whose main document part cannot be parsed now opens with the "could not be displayed" page instead of rejecting.',
          'New additive APIs: `partFailures` and the `onPartFailure` viewer option list parts that failed to parse; `errorPlaceholder` replaces the placeholder\'s title and detail text; `getImage()` on DOCX and XLSX Node sessions returns picture bytes as PPTX sessions already did; and `PptxScrollViewer` gains `hiddenSlideMode` with `\'skip\'` and `\'dim\'` modes.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'If you pass `useGoogleFonts`, migrate it. `useGoogleFonts: true` becomes `fonts: [googleFonts()]`, which sends the same stylesheet requests; `useGoogleFonts: false` is omitted. The alias still works in v0.90.0 with a warning printed once per JavaScript realm, and when `fonts` is also given, `fonts` wins: `useGoogleFonts` adds nothing, a warning is printed and nothing is thrown. From v0.91.0 it is ignored with the same warning and loads no web fonts, but never throws.',
          'If your page has a Content Security Policy and loads web fonts, allow the font file host (for example `fonts.gstatic.com`) in `connect-src` as well as `font-src`: font files are now fetched by the library, and a blocked fetch falls back to the host fonts.',
          'If your application inspects `OoxmlError.code`, one new code can appear: `font-source-limit`, thrown by a font source factory for a value beyond a fixed limit. It reports an application configuration problem, not a document problem.',
          'TypeScript users of the low-level model types may see compile errors where the types now match the parser output: some fields became optional or nullable, and `ChartThreeD.view3DPresent` and PPTX `TableCell.diagonalTL`/`diagonalTR` are renamed to `view3dPresent` and `diagonalTl`/`diagonalTr`. The data sent by the parser is unchanged.',
          'To try OpenDocument reading, install `@silurus/ooxml-opendocument` at exactly the same version as `@silurus/ooxml` and upgrade both together. The same exact-version rule applies to `@silurus/ooxml-legacy-office`; a mismatched pair fails before loading with `companion-runtime-mismatch`.',
        ],
        bullets: [
          '`useGoogleFonts: true` → `fonts: [googleFonts()]`',
          '`useGoogleFonts: false` → omit the option',
          'A custom Google-compatible origin → `googleFonts({ baseUrl })`; other sources → `fontCssApi()`, `fontCss()`, `fontFiles()` or `fontResolver()`',
        ],
        links: [
          { label: 'Font sources guide, including migration', href: '/fonts#migration' },
          { label: 'OpenDocument files guide', href: '/opendocument' },
        ],
      },
    ],
  },
  {
    slug: 'experimental-legacy-office',
    date: '2026-10-19',
    label: 'Release note',
    version: 'v0.89.0',
    title: 'Windows-Office layout, full picture effects and experimental legacy Office reading',
    summary: 'v0.89.0 lays out Word documents from the fonts actually installed on the viewing host and from Windows Office observations, renders the full DrawingML picture-effect chain in every format, speeds up first display, hardens parsing against hostile input and adds optional, experimental reading of a supported subset of .doc, .xls and .ppt files. OOXML Viewer remains an OOXML viewer; legacy reading is enabled separately.',
    audience: 'All Viewer and Node integrations. No migration is required; applications that branch on OoxmlError codes should handle the new codes listed below.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This release concentrates on dependable, spec-honest display. Word, Excel and PowerPoint layout now follows rules measured against Windows Office instead of single-sample constants, pictures keep their authored effects, first display is faster, and malformed or hostile files are bounded and rejected with typed errors.',
        ],
        bullets: [
          'Lay out DOCX text from the installed face that actually paints, with Windows Word line, spacing, grid and table placement rules.',
          'Render the full `CT_Blip` picture-effect chain, path gradients and preset patterns across DOCX, XLSX and PPTX.',
          'Bring PowerPoint text, Excel cells and ChartEx charts closer to their Office hosts.',
          'Paint the first XLSX viewport before the worksheet finishes loading, restore PPTX first-display speed and keep DOCX main-mode layout responsive.',
          'Bound parser amplification, fail closed on resource budgets and stop document text from injecting Markdown links, HTML or shell commands.',
          'Optionally preview a supported subset of .doc, .xls and .ppt files through the experimental `@silurus/ooxml-legacy-office` add-on.',
        ],
      },
      {
        title: 'Word layout from the host font and Windows Word rules',
        modules: ['DOCX'],
        rationale: 'Line heights and breaks are decided by the font that paints the text, so layout is computed from that font rather than from a guessed substitute or from the viewer\'s operating system.',
        paragraphs: [
          'Lines are sized from the installed face that paints each run, including absent authored faces and empty paragraph marks, and layout no longer depends on the viewer OS. Where Word for Windows and Word for Mac differ on identical font data, the release follows Windows.',
          'Calibrated single-sample constants were replaced with rules derived from Windows Word controls: automatic spacing between East Asian and Latin text, justified interword compression and kerning, line-grid snapping for inline objects and minimum-spaced text, blank-header reserve, page-top spacing and footnote stacking.',
          'Page-anchored floating tables now flow by their anchor as Word does, including moving to the next page and yielding to later tables, and floats from different paragraphs are placed where Word places them.',
          'Picture bullets, picture brightness and contrast, grouped VML, inline groups, rotated text boxes and Word numbering-level behaviour now render as authored.',
        ],
      },
      {
        title: 'Picture effects, gradients and patterns in every format',
        modules: ['DOCX', 'XLSX', 'PPTX'],
        paragraphs: [
          'Pictures render their complete DrawingML effect chain: colour changes, duotone, luminance, tints, blur, glow, shadows, soft edges, reflection, fill overlays and blend containers. Rectangular and outline path gradients, preset patterns with Office cell phase, WMF clipping and LZW TIFF strips are drawn as authored, and DrawingML line ends are sized like Office.',
          'A picture-effect chain that is malformed or cannot be represented now raises `OoxmlError(\'unsupported-image-effect\')` instead of silently drawing a different image.',
        ],
      },
      {
        title: 'PowerPoint text, Excel sheets and charts',
        modules: ['PPTX', 'XLSX', 'Charts'],
        paragraphs: [
          'PowerPoint text follows the list-style chain for faces and sizes, built-in and per-level paragraph spacing, font alignment, compatible line spacing, per-language East Asian and complex-script theme fonts, and stacked vertical WordArt. Table style fallbacks, theme overrides and equation fallbacks also match PowerPoint more closely.',
          'Excel derives column widths from the workbook\'s first font, applies table and PivotTable style font colours as Excel does, formats time-only and fractional-second formats, keeps cached `TODAY()` and `NOW()` values and renders one-cell-anchored pictures. The viewer gains a sheet list popover from the tab navigation, whole-row and whole-column selection outlines, and resizing of several selected bands at once.',
          'Box-and-whisker, histogram and Pareto ChartEx charts follow Office aggregation, binning and paint order, and date and time axis labels use the worksheet number formatter.',
          'Chart category, value, legend, data-table, title and text-box labels now keep authored line breaks, including `<a:br>` in chart text, and a multi-paragraph chart title renders one line per paragraph.',
        ],
      },
      {
        title: 'Faster first display and lower memory',
        paragraphs: [
          'XLSX paints the first viewport exactly before the worksheet pull completes and evicts inactive worksheet models before failing on the workbook budget. PPTX first display regains the speed lost since v0.88.0. DOCX main-mode layout and finalization are sliced to keep the main thread responsive, and retained and peak heap are reduced. CJK line breaking is now linear in the text length.',
        ],
      },
      {
        title: 'Hardening against hostile input',
        paragraphs: [
          'Parsers bound style-chain, field-stack, theme-style, column-style and input-sized allocations, and parser and image budgets fail closed with messages that name the resource and stage. Markdown projections escape document text, so documents can no longer inject links, HTML or shell commands. The MCP server confines file access to its configured path scope, the VS Code extension limits its webview resource roots, and the zip dependency was upgraded for published Rust advisories.',
          'Non-OOXML input raises `OoxmlError(\'not-ooxml\')`, and a DOCX numbering level that Word cannot open raises `OoxmlError(\'invalid-numbering\')`.',
        ],
      },
      {
        title: 'Experimental legacy Office reading, as an optional add-on',
        modules: ['@silurus/ooxml-legacy-office'],
        paragraphs: [
          'The experimental `@silurus/ooxml-legacy-office` add-on opens supported binary Word, Excel and PowerPoint files in the same viewers used for DOCX, XLSX and PPTX. It reads directly into the existing models without generating an intermediate OOXML package, keeps legacy readers outside ordinary OOXML imports, and rejects unsupported input instead of displaying a partial document. Try experimental legacy previews locally in Try Yours, without uploading a file.',
        ],
      },
      {
        title: 'One add-on, three explicit sources',
        modules: ['@silurus/ooxml-legacy-office/doc', '@silurus/ooxml-legacy-office/xls', '@silurus/ooxml-legacy-office/ppt'],
        rationale: 'A separate package lets applications choose legacy previews without adding binary readers to ordinary OOXML viewing.',
        paragraphs: [
          'Install `@silurus/ooxml` and `@silurus/ooxml-legacy-office` together; their versions must match exactly. Import `legacyDoc`, `legacyXls` or `legacyPpt` and pass it to the matching viewer or Node session as `{ legacyDoc }`, `{ legacyXls }` or `{ legacyPpt }`, like the other opt-in modules. Importing a reader fetches nothing; it loads only when a matching file is opened. For custom asset URLs, create one with `createLegacyDoc({ wasmUrl, moduleUrl })` (likewise `createLegacyXls` and `createLegacyPpt`).',
        ],
        examples: [
          {
            title: 'Enable experimental DOC reading',
            code: `import { DocxViewer } from '@silurus/ooxml/docx';
import { legacyDoc } from '@silurus/ooxml-legacy-office/doc';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { legacyDoc });
await viewer.load(await file.arrayBuffer());`,
          },
        ],
      },
      {
        title: 'An experimental subset with a conversion fallback',
        paragraphs: [
          'The readers support a bounded subset of each binary format. They do not execute macros, fields, formulas, actions or embedded objects. Unsupported structures reject the load; keep a conversion fallback for files the reader cannot represent. This is not a claim of Office fidelity.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No migration is required. Existing imports, options and loading calls keep working without code changes.',
          'If your application inspects `OoxmlError.code`, four new codes can now appear: `not-ooxml` for files that are not Office documents, `unsupported-image-effect` for picture effects that cannot be drawn, `invalid-numbering` for Word numbering that Word itself cannot open, and `unsupported-content` for files the legacy add-on cannot read. Show them like any other load error.',
          'Two optional settings are new. `xlsxWorksheetLimits` changes the worksheet size limits for one load, and `XlsxSheetViewer` options `minRows`, `minCols`, `marginRows` and `marginCols` set the minimum grid size and the blank rows and columns shown after the content.',
          'To try the legacy add-on, install `@silurus/ooxml-legacy-office` at exactly the same version as `@silurus/ooxml`, configure only the sources you need, and upgrade both packages together from then on.',
          'See the Legacy Office files guide for browser and Node examples, deployment assets, support boundaries and failure handling.',
        ],
      },
    ],
  },
  {
    slug: 'v088-chart-fidelity-and-safer-word-layout',
    date: '2026-09-21',
    label: 'Release note',
    version: 'v0.88.0',
    title: 'More faithful charts and safer Word layout in v0.88.0',
    summary: 'v0.88.0 brings classic Office charts closer to their authored appearance across Word, Excel and PowerPoint, while bounding runaway keep-with-next pagination in complex Word documents.',
    audience: 'Applications that display Office charts, complex Word documents, presentation SVGs or multilingual PowerPoint text. Most Viewer integrations can upgrade without code changes.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This release concentrates on visual fidelity and dependable viewing. Classic charts now follow the built-in Office styles more closely across all three formats, and Word layout stays bounded when a document contains unusually long groups that must remain together.',
        ],
        bullets: [
          'Render classic chart frames, colours, fills, outlines, labels and effects more like the Office host that authored them.',
          'Bound repeated keep-with-next pagination work before it can exhaust browser memory.',
          'Preserve PowerPoint SVG artwork when its compatibility image is unavailable.',
          'Prepare the selected regional CJK fallback font consistently in PowerPoint main-thread loading.',
        ],
      },
      {
        title: 'Classic charts across Word, Excel and PowerPoint',
        modules: ['DOCX', 'XLSX', 'PPTX'],
        paragraphs: [
          'Classic charts now resolve the built-in numeric style range 1 through 48 while retaining each host application\'s visual conventions. Chart frames, automatic text contrast, series and point styling, picture fills, markers, labels, legends, axes and three-dimensional surfaces all keep authored formatting ahead of built-in style defaults.',
          'The change is intentionally visible in affected documents. Existing charts may gain the frame, palette, text contrast, transparency or outline that Word, Excel or PowerPoint applies to the saved style. Direct formatting and linked chart styles remain authoritative, and ChartEx continues through its existing optional entry point.',
        ],
      },
      {
        title: 'Safer and more faithful Word layout',
        modules: ['DOCX', 'Node'],
        paragraphs: [
          'Word documents with very long keep-with-next groups now stop speculative measurement once the group cannot fit on a fresh page. Paragraph layout also keeps a bounded working set and fails with a layout diagnostic if repeated acquisition exceeds its operational limit, instead of continuing until the browser or Node process runs out of memory.',
          'Default paragraph and table styles now participate in inheritance for every valid on/off form.',
        ],
      },
      {
        title: 'Presentation and spreadsheet details',
        modules: ['PPTX', 'XLSX'],
        paragraphs: [
          'PowerPoint image-filled shapes can use their SVG source when the older compatibility raster is missing or unreadable. If neither source resolves, the authored image fill remains transparent instead of turning into an unrelated solid fill. When Google Fonts loading is enabled, main-thread PowerPoint loading also prepares the selected regional CJK fallback family, matching the other loading modes.',
          'The built-in Excel zoom slider now settles naturally on its 100% centre mark when dragged nearby, while continuous zoom remains available immediately outside that small attraction zone. Programmatic zoom and the existing buttons and wheel controls keep their previous behaviour.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No public option or method is removed or renamed, and most Viewer integrations require no code changes. The release intentionally changes rendering for documents that rely on corrected Office chart styles, Word layout rules, SVG fallbacks or regional CJK font selection. Applications with pixel baselines should review those differences against the authored document or an Office-produced reference instead of expecting every pixel to match v0.87.0.',
          'Applications that implement a custom chart renderer against the low-level public chart model should review the added optional style data and the broader image-capable fill type. Built-in chart rendering requires no setup change.',
          'The Word layout limits apply only after unusually repeated acquisition work. Ordinary documents keep the same loading contract; a document that exceeds the limit now reports non-convergence instead of consuming memory without a bound.',
        ],
      },
      {
        title: 'Technical note',
        paragraphs: [
          'The public chart model gains optional style and effect metadata, and a small number of paint fields accept the broader Fill union so picture fills can retain their authored form. Viewer integrations need no changes. Applications that exhaustively match the low-level chart model should account for the added optional data and image-fill case when upgrading.',
          'Word paragraph acquisition now uses bounded internal caching and a pagination-scoped operational budget. The budget counts cache misses within one pagination session: it permits 25,000 misses, returns a NON_CONVERGENCE layout diagnostic on miss 25,001, and retains only the two most recent placements for each paragraph. The safety boundary does not change the successful layout contract for ordinary documents.',
        ],
      },
    ],
  },
  {
    slug: 'v087-regional-cjk-fallbacks',
    date: '2026-09-13',
    label: 'Release note',
    version: 'v0.87.0',
    title: 'Regional CJK fallbacks and steadier loading in v0.87.0',
    summary: 'v0.87.0 lets applications choose regional CJK fallback forms and picks up several reliability fixes for progressive loading, navigation, and file opening.',
    audience: 'Applications that display CJK documents, use progressive DOCX or PPTX loading, or need consistent programmatic navigation. Existing integrations can upgrade without API changes.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'This release focuses on predictable multilingual display rather than a large new viewing feature. Word, Excel and PowerPoint content can use the appropriate regional fallback for ambiguous Han characters, alongside reliability fixes for loading, navigation and file opening.',
        ],
        bullets: [
          'Choose Simplified Chinese, Traditional Chinese, Hong Kong, Japanese or Korean fallback forms.',
          'Use the host language automatically, or choose one fallback region independently of the host locale.',
          'Keep fast progressive loads and programmatic page or slide navigation consistent.',
        ],
      },
      {
        title: 'Regional CJK fallback',
        modules: ['DOCX', 'XLSX', 'PPTX', 'Node'],
        paragraphs: [
          'The same Unicode Han character can have different regional glyph shapes. The new cjkFallback option chooses which Simplified Chinese, Traditional Chinese, Hong Kong, Japanese or Korean fallback family is tried first when the document does not already identify a region. Choose sc, tc, hk, jp or kr explicitly, or leave the default auto mode to resolve the host language once when loading starts.',
          'The document\'s requested font remains first. Recognized regional CJK font names, an East Asian run language where the format provides one, and unambiguous Kana or Hangul can determine the region before cjkFallback is consulted. Text without Han keeps its existing font route.',
          'cjkFallback changes the fallback preference; it does not install or download fonts. An explicit value makes that regional choice independent of the host locale, but pixel-identical output across hosts still requires the same fonts. Applications can provide those fonts themselves or separately opt into useGoogleFonts.',
        ],
        examples: [
          {
            title: 'Choose a fallback region independently of the host locale',
            code: `const viewer = new DocxViewer(canvas, {
  cjkFallback: 'sc',
});`,
          },
        ],
      },
      {
        title: 'Steadier completion, navigation and file opening',
        paragraphs: [
          'Progressive DOCX and PPTX loads now deliver one successful onLayoutComplete notification even when a small file finishes before load() returns. Programmatic page and slide navigation also stays on the requested item when a browser rounds a fractional scroll position.',
          'The shared package reader accepts otherwise consistent Office files whose duplicated ZIP headers differ only in legacy timestamp metadata. Word intrinsic-width measurement also handles empty anchored runs without failing the document layout.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No option or method is removed or renamed, and most applications can upgrade without changes. The auto CJK fallback can intentionally choose different regional Han glyphs when the host language provides that preference. Set cjkFallback explicitly when the fallback region must not depend on the browser locale, server or Node host. This stabilizes the regional choice; exact visual output still depends on the fonts available in each environment.',
          'Applications that observe progressive onLayoutComplete should allow the callback after every successful progressive load. It is still called exactly once; failures before any layout is published continue to reject load() directly.',
        ],
      },
    ],
  },
  {
    slug: 'v086-presentation-text-and-csv-previews',
    date: '2026-09-05',
    label: 'Release note',
    version: 'v0.86.0',
    title: 'More faithful presentation text, plus CSV previews, in v0.86.0',
    summary: 'v0.86.0 improves multiline and CJK text in PowerPoint, keeps usable Excel content visible in more files, and changes Markdown exports to collect review comments separately. CSV and TSV previews are included as a small extra.',
    audience: 'Applications that export Office files as Markdown, display multilingual presentations, or open varied Excel files. Most applications can upgrade without code changes.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'PowerPoint presentations now keep multiline and mixed-language text closer to their authored appearance. Excel also continues displaying usable content in more cases when one item on a sheet cannot be shown. Markdown exports now collect review discussions separately from document text.',
        ],
        bullets: [
          'Improve PowerPoint multiline spacing and wrapping across mixed CJK and Latin text.',
          'Match more native Noto CJK font names across Word, Excel and PowerPoint.',
          'Keep usable Excel content visible when an unsupported picture or legacy sheet is encountered.',
          'Collect review comments and replies after the document body in Markdown exports.',
        ],
      },
      {
        title: 'Presentation text, Markdown, and steadier viewing',
        paragraphs: [
          'DOCX, XLSX and PPTX Markdown exports now collect review comments and their replies in a final quoted appendix, separate from the document body. PowerPoint speaker notes remain with their slide. This remains a best-effort text projection for search, comparison and AI workflows, not a reconstruction of visual layout.',
          'PowerPoint text wraps more naturally across mixed CJK and Latin content, and multiline spacing more closely follows PowerPoint when no explicit spacing is stored. Shared CJK font matching also recognizes more native Noto family names across Word, Excel and PowerPoint.',
          'Excel workbooks created by some older libraries open more reliably. Legacy dialog sheets now show a neutral notice, and an unsupported TIFF picture no longer prevents the usable cells and other pictures on that sheet from appearing.',
        ],
      },
      {
        title: 'A small extra: preview delimited text',
        paragraphs: [
          'As a small extra, CSV and TSV files can now use the same read-only sheet surface as Excel workbooks. Try Yours recognizes those two formats automatically, and library applications can preview text separated by another chosen delimiter in regular or worker rendering mode.',
          'Pass an explicit format when loading delimited text with XlsxSheetViewer. CSV uses a comma and TSV uses a tab by default; delimited-text accepts another single-character separator. Dates, leading zeroes, long identifiers and values beginning with = remain authored text instead of being inferred as spreadsheet values.',
          'This is a focused preview convenience rather than spreadsheet import. XlsxWorkbook and the container-backed XlsxViewer continue to open OOXML workbooks only.',
        ],
        examples: [
          {
            title: 'Preview delimited text',
            code: `await sheet.load('/export.csv', { format: 'csv' });
await sheet.load('/report.txt', {
  format: 'delimited-text',
  delimiter: '|',
});`,
          },
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'No viewer API migration is required. Existing DOCX, XLSX and PPTX loading keeps the same defaults and public methods. Applications that parse generated Markdown should account for review comments moving to the final appendix.',
          'Delimited-text preview is opt-in for library applications: pass format: \'csv\', format: \'tsv\' or format: \'delimited-text\' to XlsxSheetViewer.load(). Existing XLSX calls without a format continue to open workbooks as before.',
        ],
      },
    ],
  },
  {
    slug: 'v085-large-images-and-rendering',
    date: '2026-09-02',
    label: 'Release note',
    version: 'v0.85.0',
    title: 'Larger images and steadier rendering in v0.85.0',
    summary: 'v0.85.0 opens more Office files with large images, broadens TIFF support and improves rendering stability in worker mode and during PowerPoint zoom.',
    audience: 'Applications that display image-heavy DOCX, XLSX or PPTX files, use worker rendering, or open TIFF images. Most applications can upgrade without code changes.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'Word, Excel and PowerPoint files with large, high-resolution images can now open more reliably. Images are prepared for the current display size when retaining their full source resolution would be excessive, and a later zoom can request a sharper result.',
        ],
        bullets: [
          'Display poster-sized and camera images that previously crossed the decoded-image limit.',
          'Support more TIFF images, including common black-and-white, grayscale, colour and Group 4 files.',
          'Keep embedded SVGs, PowerPoint text and styled Excel sheets dependable in worker mode.',
        ],
      },
      {
        title: 'Smoother, more faithful viewing',
        paragraphs: [
          'PowerPoint zoom now keeps the viewed slide in place while the sharper render settles, including at the maximum zoom level. Placeholder text, built-in table styles and several chart layouts also more closely follow their Office appearance.',
          'The same image handling is available across DOCX, XLSX and PPTX, in both regular and worker rendering modes.',
        ],
      },
      {
        title: 'Upgrading',
        paragraphs: [
          'Most applications can upgrade without code changes. The new image policy is automatic, and applications with a specific memory or image-quality requirement can optionally adjust imageResources.',
          'TIFF support remains optional. Without the TIFF module, the affected image is shown as unavailable while the rest of the document remains viewable. Add @silurus/ooxml/tiff only when those images need to be displayed.',
        ],
      },
      {
        title: 'Technical note',
        paragraphs: [
          'Adaptive image sizing remains bounded by non-disableable source, dimension and per-image safety limits. imageResources controls the decoded working budget and can select strict rejection when an application prefers it; it is not a promise about total browser or GPU memory.',
        ],
      },
    ],
  },
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

export interface InlineSegment {
  readonly text: string;
  readonly code: boolean;
}

/** Split announcement prose on backticks: `name` renders as inline code. An
 *  unpaired trailing backtick stays literal text. */
export function inlineSegments(value: string): InlineSegment[] {
  const parts = value.split('`');
  if (parts.length % 2 === 0) parts[parts.length - 2] += '`' + parts.pop();
  return parts
    .map((text, index) => ({ text, code: index % 2 === 1 }))
    .filter(({ text, code }) => code || text.length > 0);
}

export function formatAnnouncementDate(value: string): string {
  return new Intl.DateTimeFormat('en', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  }).format(new Date(`${value}T00:00:00Z`));
}
