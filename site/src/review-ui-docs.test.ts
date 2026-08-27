import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (path: string): string => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('Office comment UI integration guide', () => {
  const page = read('./pages/review-ui.astro');

  it('starts with the built-in path and keeps low-level composition available', () => {
    expect(page).toContain('<h1>Comments in DOCX, XLSX, and PPTX</h1>');
    expect(page).toContain('reads comments already stored in Office files');
    for (const concept of ['Comments and threads', 'Anchors', 'Built-in presentation', 'Surface geometry']) {
      expect(page).toContain(concept);
    }
    for (const viewer of ['DocxScrollViewer', 'PptxScrollViewer', 'XlsxViewer']) {
      expect(page).toContain(viewer);
    }
    expect(page).toContain('Built-in or custom');
    expect(page).toContain('Application-owned list, Viewer-owned target');
    expect(page).toContain('keeps the list outside the document surface');
    expect(page).toContain("viewer.goToComment(comment.id, { behavior: 'smooth' })");
    expect(page).toContain("viewer.goToComment(slideIndex, commentIndex, { behavior: 'smooth' })");
    expect(page).toContain("viewer.goToComment(sheetIndex, comment.cellRef, { align: 'center' })");
    expect(page).toContain('DocxScrollViewer.fromDocument(container, document');
    expect(page).toContain('PptxScrollViewer.fromPresentation(container, presentation');
    expect(page).toContain('XlsxViewer.fromWorkbook(container, workbook');
    expect(page).toContain('<CodeTabs id="comment-list-navigation-code" tabs={commentListTabs} />');
    expect(page).not.toContain('const independentListExample');
    expect(page.indexOf('id="css-variables-title"')).toBeLessThan(
      page.indexOf('id="stable-classes-title"'),
    );
    expect(page).toContain('await workbook.getComments(sheetIndex)');
    expect(page).toContain('<code>comments</code>');
    expect(page).toContain('<code>markers</code>');
    expect(page).toContain('<code>cards</code>');
    expect(page).toContain('Defaults to <code>true</code>');
    expect(page).toContain('Set it to <code>false</code> to hide only the icons');
    expect(page).toContain('markers: true, // Default. Set false to hide the message icons only.');
    expect(page).toContain('Defaults to <code>false</code> for DOCX/PPTX and <code>true</code> for XLSX');
    expect(page).toContain('--ooxml-comment-card-background');
    expect(page).toContain('--ooxml-comment-card-border-left: 3px solid #2563eb');
    expect(page).toContain('--ooxml-comment-card-border-right: 1px solid #cbd5e1');
    expect(page).toContain('--ooxml-comment-card-radius: 8px');
    expect(page).toContain('--ooxml-comment-marker-color: #2563eb');
    expect(page).toContain('.ooxml-comment-card__author');
    expect(page).toContain('.ooxml-comment-card[data-active="true"]');
    expect(page).toContain('backdrop-filter: blur(14px)');
    expect(page).toContain('fixed blue values deliberately replace per-author accents');
    expect(page).toContain('omit them to keep automatic author colors');
    expect(page).toContain('Optional DOCX/PPTX connectors');
    expect(page).toContain('Style the built-in UI');
    expect(page).toContain('Two complementary styling surfaces are available');
    expect(page).toContain('Stable classes');
    expect(page).toContain('Control the complete appearance');
    expect(page).toContain('CSS variables');
    expect(page).toContain('Switch common theme tokens');
    expect(page).toContain('does not mirror every CSS property');
    expect(page).toContain('updates mounted cards, highlights, and markers without recreating the Viewer');
    expect(page).toContain('font size, font family, or padding are measured again');
    expect(page).toContain(
      '<tr><th><code>data-active</code>, <code>data-focused</code></th><td>DOCX, PPTX</td>',
    );
    expect(page).not.toContain(
      '<tr><th><code>data-active</code>, <code>data-focused</code></th><td>DOCX, PPTX, XLSX</td>',
    );
    expect(page).toContain('var(--review-connector)');
    expect(page).toContain("side: 'auto'");
    expect(page).toContain("route: 'bezier'");
    expect(page).toContain("stroke: 'solid'");
    expect(page).toContain("activeColor: '#2563eb'");
    expect(page).not.toContain('All comment theme properties');
    expect(page).not.toContain('commentUi');
    expect(page).not.toContain('--ooxml-comment-avatar');
    expect(page).not.toContain('mountCard');
    expect(page).not.toContain('commentRenderer');
  });

  it('presents built-in options as a cross-format support matrix', () => {
    const match = page.match(/aria-label="Built-in comment options"[\s\S]*?<\/table>/);
    expect(match).not.toBeNull();
    const table = match?.[0] as string;

    expect(table).toContain('<th scope="col">Option</th>');
    for (const format of ['DOCX', 'XLSX', 'PPTX']) {
      expect(table).toContain(`<th class="format-column" scope="col">${format}</th>`);
    }
    expect(table).toContain('<th scope="col">Behavior</th>');
    expect(table.match(/<code>comments<\/code>/g)).toHaveLength(1);
    expect(table).not.toContain('comments: true');
    expect(table).not.toContain('comments: false');
    expect(table).toContain('Set to <code>true</code>');
    expect(table).toContain('or <code>false</code>');
    expect(table).toContain('aria-label="Supported">✓</td>');
    expect(table).toContain('aria-label="Not supported">—</td>');
  });

  it('keeps the built-in preview separated and gives its frame one height contract', () => {
    const builtIn = read('./components/BuiltInCommentViewer.astro');
    expect(page).toContain('.built-in-preview { margin-top: 32px; }');
    expect(builtIn).toContain('height: clamp(520px, 76vh, 720px)');
    expect(builtIn).toContain('.built-in-comment-demo__viewer { height: 100%; }');
    expect(builtIn).toContain('background: radial-gradient(120% 80% at 50% 0%, var(--preview-top), var(--preview-bottom) 70%)');
    expect(page).toContain('not by <code>ScrollViewer</code>');
    expect(builtIn).toContain("['docx', 'xlsx', 'pptx']");
    expect(builtIn).toContain('data-comment-demo-format={format}');
    for (const format of ['docx', 'xlsx', 'pptx']) {
      expect(builtIn).toContain(`sample-1.${format}`);
    }
    expect(builtIn).toContain("new DocxScrollViewer(host, { comments: true");
    expect(builtIn).toContain("new PptxScrollViewer(host, { comments: true");
    expect(builtIn).toContain("new XlsxViewer(host, { comments: true");
    expect(builtIn).not.toContain('min-height: 640px');
  });

  it('keeps the built-in viewer alive while the page is stored in the back-forward cache', () => {
    const builtIn = read('./components/BuiltInCommentViewer.astro');
    expect(builtIn).toContain("window.addEventListener('pagehide', (event) => {");
    expect(builtIn).toContain('if (event.persisted) return;');
    expect(builtIn).not.toContain("window.addEventListener('pagehide', destroy, { once: true });");
  });

  it('maps concepts to public APIs and distinguishes transcript from anchored UI', () => {
    const example = read('./examples/review-margin/index.ts');
    for (const api of ['doc.comments', 'getCommentThreads(pageIndex', 'workbook.getComments(sheetIndex)', 'getElementBoundsByIds']) {
      expect(page).toContain(api);
    }
    expect(page).toContain('The Viewer owns its built-in presentation.');
    expect(page).toContain('The application owns custom product behavior.');
    expect(page).toContain('Call <code>goToComment()</code> to reveal and highlight');
    expect(page).not.toContain('No page geometry');
    expect(page).not.toContain('Page geometry required');
    expect(page).toContain('<CodeTabs id="comment-primitives" tabs={primitiveTabs} />');
    for (const tab of ["label: 'Source records'", "label: 'DOCX page'", "label: 'XLSX sheet'", "label: 'PPTX slide'"]) {
      expect(page).toContain(tab);
    }
    expect(page).toContain('<CommentListNavigationDemo />');
    const listDemo = read('./components/CommentListNavigationDemo.astro');
    expect(listDemo).toContain("DocxScrollViewer.fromDocument(host, doc");
    expect(listDemo.match(/comments: \{ cards: false, markers: false \}/g)).toHaveLength(2);
    expect(listDemo).toContain("viewer.goToComment(comment.id, { behavior: 'smooth' })");
    expect(listDemo).toContain('PptxScrollViewer.fromPresentation(host, presentation');
    expect(listDemo).toContain("viewer.goToComment(slideIndex, commentIndex, { behavior: 'smooth' })");
    expect(listDemo).toContain("new XlsxViewer(host, { comments: false })");
    expect(listDemo).toContain("viewer.goToComment(sheetIndex, comment.cellRef, { align: 'center' })");
    expect(listDemo).toContain('xlsxThreads(next.getComments(), next, next.sheetIndex)');
    expect(listDemo).toContain(".filter((comment) => comment.parentId === undefined && !comment.resolved)");
    for (const format of ['docx', 'xlsx', 'pptx']) {
      expect(listDemo).toContain(`sample-1.${format}`);
    }
    expect(listDemo).toContain('comment-list-demo__reply');
    expect(listDemo).toContain('--comment-list-accent');
    expect(listDemo).toContain('data-comment-list-items');
    expect(listDemo).toContain('aria-live="polite"');
    expect(example).toContain('doc.commentAnchorRanges()');
    expect(example).toContain('onTextRun: (run) => runs.push(run)');
    expect(example).toContain('resolveDocxCommentThreads(');
    expect(example).toContain('{ includeResolved: false }');
    const apiReference = read('./lib/api-reference.ts');
    expect(apiReference).toContain('get comments(): readonly Readonly<DocComment>[]');
    expect(apiReference).toContain('get revisions(): readonly Readonly<DocRevision>[]');
    expect(apiReference).toContain('commentAnchorRanges(): readonly CommentAnchorRange[]');
    expect(apiReference).toContain('getCommentThreads(pageIndex: number');
    expect(apiReference).toContain('revisionAnchorRanges(): readonly RevisionAnchorRange[]');
    expect(page).toContain('DOCX comments are stored for the document, not authored as page-owned records');
    expect(apiReference).not.toContain("{ sig: 'resolveDocxCommentThreads");
    expect(apiReference).toContain('collectPageRuns(index');
    expect(apiReference).toContain("detailsHref: '/review-ui', detailsLabel: 'Comment UI guide'");
  });

  it('documents the same built-in and primitive boundary for every Office format', () => {
    const apiReference = read('./lib/api-reference.ts');
    const apiComponent = read('./components/ApiReference.astro');
    expect(page).toContain('How comments are represented');
    expect(page).toContain('Use the same <code>comments</code> option across formats');
    expect(page).toContain('presentation.getComments(slideIndex)');
    expect(page).toContain('workbook.getComments(sheetIndex)');
    expect(page).toContain('getCellViewportRect()');
    expect(page).toContain('getSelectionContext()');
    expect(apiReference).toContain('getComments(slideIndex: number)');
    expect(apiReference).toContain('goToComment(slideIndex: number, commentIndex: number');
    expect(apiReference).toContain('goToComment(commentId: string');
    expect(apiReference.match(/goToComment\(sheetIndex: number, cellRef: string/g)).toHaveLength(2);
    expect(apiReference).toContain('getComments(sheetIndex: number): Promise<readonly Readonly<XlsxComment>[]>');
    expect(apiReference.match(/getComments\(\): readonly Readonly<XlsxComment>\[\]/g)).toHaveLength(2);
    expect(apiReference).toContain('selected comment thread');
    expect(apiReference).toContain('including attached comments');
    for (const href of ['/api/docx#docx-scroll-viewer', '/api/xlsx#xlsx-viewer', '/api/pptx#pptx-scroll-viewer']) {
      expect(page).toContain(`href="${href}"`);
    }
    expect(page).toContain('href="/api/');
    expect(apiComponent).toContain('id={classAnchor(c.name)}');
    expect(page).toContain('font-size: 14px; }');
    expect(page).toContain('padding: 10px 14px;');
    expect(page).toContain('font: 600 12px var(--mono);');
  });

  it('keeps the overview short while linking the DOCX demo and reusable source', () => {
    const core = read('./examples/review-margin/core.ts');
    const markup = read('./examples/review-margin/index.html');
    const controller = read('./examples/review-margin/index.ts');
    const styles = read('./examples/review-margin/styles.css');
    expect(page).not.toContain('<ReviewGuideExample />');
    expect(page).not.toContain('<h2 id="live-example-title">Live example</h2>');
    expect(page).not.toContain('href="/docx#review-ui"');
    expect(page).toContain('href="#built-in-ui"');
    expect(markup).toContain('data-review-example');
    expect(core).toContain('export async function renderReviewPage(');
    expect(core).toContain('resolveDocxCommentThreads(');
    expect(page).toContain('href="/review-ui/source"');
    expect(page).toContain('Complete DOCX custom UI');
    const customUiSection = page.slice(
      page.indexOf('<section class="review-section" aria-labelledby="minimum-flow">'),
      page.indexOf('<section class="review-section next-steps"'),
    );
    expect(customUiSection).toContain('href="/review-ui/source"');
    expect(page.match(/href="\/review-ui\/source"/g)).toHaveLength(1);
    const sourcePage = read('./pages/review-ui/source.astro');
    expect(sourcePage).toContain('<ReviewDemo showCode={false} />');
    expect(sourcePage).toContain('blob/main/site/src/components/ReviewDemo.astro');
    expect(sourcePage).toContain('blob/main/site/src/lib/review-demo.ts');
    expect(sourcePage).not.toContain("ReviewDemo.astro?raw");
    expect(sourcePage).not.toContain('<CodeTabs id="review-example-complete-source"');
    expect(sourcePage).toContain('<CodeTabs id="review-example-portable-source"');
    expect(sourcePage).toContain('uses only <code>@silurus/ooxml/docx</code>');
    expect(controller).toContain('export async function mountReviewExample(root: HTMLElement, signal?: AbortSignal)');
    expect(controller).toContain('const pageIndex = Math.max(0, Math.min(requestedPage, doc.pageCount - 1))');
    expect(controller).toContain('await doc.renderPage(stage, pageIndex');
    expect(controller).toContain("Number(root.dataset.page ?? 0)");
    expect(controller).toContain('updatePage(pageIndex: number): Promise<void>');
    expect(controller).toContain('const destroy = (): void =>');
    expect(controller).toContain('request !== generation');
    expect(controller).not.toContain('function lineBands(');
    expect(controller).toContain('thread.anchors.flatMap(({ rects }) => rects)');
    expect(styles).toContain('.review-example__margin');
    expect(styles).toContain('@container (max-width: 720px)');
    expect(styles).not.toContain('@media (max-width: 720px)');
  });

  it('keeps the public guide focused on comments and delegates change-history design', () => {
    for (const term of ['Insertions', 'move destinations', 'Deletions', 'move sources']) {
      expect(page).not.toContain(term);
    }
    expect(page).not.toContain('accepted-final document');
    expect(page).not.toContain('DOCX tracked changes');
    expect(page).toContain('Canvas content needs an accessible transcript');
    expect(page).toContain('href="/production#rendering-mode"');
    expect(page).not.toContain('href="/production#ownership"');
    expect(page).toContain('Open the finished custom UI, its exact source, and a portable starter.');
    expect(page).toContain('Look up comment records, anchor ranges, text runs, and method signatures.');
    expect(page).toContain('Choose main-thread or Worker rendering for your application.');
    expect(page).not.toContain('Manage a loaded document when it is shared by more than one view.');
    for (const detail of ['geometryFallback', 'storyInstance', 'sourceRunIndex', 'Device pixel ratio', 'renderPageToBitmap', 'useGoogleFonts', 'same-origin or return suitable CORS headers']) {
      expect(page).not.toContain(detail);
    }
  });

  it('keeps cross-format change history distinct from comments without inventing placeholder APIs', () => {
    const design = read('../../docs/review-ui-extension-design.md');
    const apiLayout = read('./layouts/ApiPage.astro');
    const changeHistory = read('./components/ChangeHistoryApiGuide.astro');
    expect(design).toContain('Change-history boundary and future API symmetry');
    expect(design).toContain('SpreadsheetML revision headers and revision logs');
    expect(design).toContain('PresentationML has no general revision-log model equivalent');
    expect(design).toContain('same three-layer architecture as comments');
    expect(design).toContain('record and geometry types remain format-specific');
    expect(design).toContain('Unsupported formats expose no placeholder method');
    expect(design).toContain('not masquerade as revision records read from one presentation');
    expect(apiLayout).toContain('<a href="#review-data">Review data</a>');
    expect(apiLayout).toContain('<ChangeHistoryApiGuide format={format} />');
    expect(apiLayout.indexOf('<ApiReference format={format} />')).toBeLessThan(
      apiLayout.indexOf('<ChangeHistoryApiGuide format={format} />'),
    );
    expect(changeHistory).toContain('<h2>Review data</h2>');
    expect(changeHistory).toContain('<h3 id="revision-records">Office revision records</h3>');
    expect(changeHistory).toContain('<a href="/review-ui">Comments guide →</a>');
    expect(changeHistory).toContain('<strong>Available for DOCX.</strong>');
    expect(changeHistory).toContain('<strong>Not available for XLSX.</strong>');
    expect(changeHistory).toContain('<strong>Not available for PPTX.</strong>');
    expect(changeHistory).toContain('document.revisions');
    expect(changeHistory).toContain('resolveRevisionAnchorRuns(range, runs)');
    expect(changeHistory).toContain('You can');
    expect(changeHistory).toContain('You cannot');
  });

  it('renders resolved threads, fallback carets, linked controls, and live states', () => {
    const example = read('./examples/review-margin/index.ts');
    const markup = read('./examples/review-margin/index.html');
    expect(example).toContain('resolveDocxCommentThreads(');
    expect(example).toContain("item.marker === 'range' ? band.width : 3");
    expect(example).toContain("rect.setAttribute('aria-hidden', 'true')");
    expect(example).toContain("rect.setAttribute('focusable', 'false')");
    expect(example).not.toContain("rect.setAttribute('role', 'button')");
    expect(example).toContain("button.setAttribute('aria-controls', anchorIds.join(' '))");
    expect(example).toContain("root.addEventListener('mouseover'");
    expect(example).toContain("root.addEventListener('focusin'");
    expect(example).toContain('const controllers = mounted.filter');
    expect(example).toContain('updatePage(page).catch');
    expect(example).not.toContain('sameSource');
    expect(markup).toContain('role="status"');
    expect(markup).toContain('data-review-empty');
    expect(markup).toContain('aria-label="Review transcript"');
  });

  it('keeps code tabs keyboard-operable and announces copy results', () => {
    const tabs = read('./components/CodeTabs.astro');
    expect(tabs).toContain('aria-controls={`${id}-panel-${t.id}`}');
    expect(tabs).toContain('aria-labelledby={`${id}-tab-${t.id}`}');
    expect(tabs).toContain("event.key === 'ArrowRight'");
    expect(tabs).toContain("event.key === 'Home'");
    expect(tabs).toContain("event.key === 'End'");
    expect(tabs).toContain('aria-live="polite"');
    expect(tabs).toContain('Copy failed. Select the code and copy it manually.');
  });

  it('is linked as durable format guidance and included in the sitemap', () => {
    expect(read('./pages/docx.astro')).toContain('<FormatCommentsDemo format="docx" />');
    expect(read('./pages/pptx.astro')).toContain('<FormatCommentsDemo format="pptx" />');
    expect(read('./pages/review-ui/source.astro')).toContain('<ReviewDemo showCode={false} />');
    expect(read('./pages/sitemap.xml.ts')).toContain("'/review-ui/'");
    expect(read('./pages/sitemap.xml.ts')).toContain("'/review-ui/source/'");
  });

  it('keeps the format pages simple and moves the consumer-owned UI to its source page', () => {
    const docxPage = read('./pages/docx.astro');
    const pptxPage = read('./pages/pptx.astro');
    const builtIn = read('./components/FormatCommentsDemo.astro');
    const sourcePage = read('./pages/review-ui/source.astro');
    const component = read('./components/ReviewDemo.astro');
    const controller = read('./lib/review-demo.ts');
    const sampleCopy = read('../scripts/copy-samples.mjs');

    expect(docxPage).toContain('<FormatCommentsDemo format="docx" />');
    expect(pptxPage).toContain('<FormatCommentsDemo format="pptx" />');
    expect(docxPage).not.toContain('<ReviewDemo />');
    expect(pptxPage).not.toContain('<ReviewDemo />');
    expect(docxPage.indexOf('<FormatCommentsDemo')).toBeGreaterThan(docxPage.indexOf('kind="masterdetail"'));
    expect(pptxPage.indexOf('<FormatCommentsDemo')).toBeGreaterThan(pptxPage.indexOf('kind="masterdetail"'));
    expect(builtIn).toContain("viewer: 'DocxScrollViewer'");
    expect(builtIn).toContain("viewer: 'PptxScrollViewer'");
    expect(builtIn).toContain('comments: true');
    expect(builtIn).toContain('formats={[format]}');
    expect(builtIn).toContain('href="/review-ui"');
    expect(sourcePage).toContain('<ReviewDemo showCode={false} />');
    expect(component).toContain('data-review-connectors');
    expect(component).toContain('id="review-ui"');
    expect(component).toContain('class="review-margin"');
    expect(component).toContain('Core wiring');
    expect(component).toContain('resolveCommentAnchorRuns');
    expect(component).toContain('resolveRevisionAnchorRuns');
    expect(component).toContain('href="/review-ui"');
    expect(component).toContain('themes={codeThemes} defaultColor={false}');
    expect(component).not.toContain('data-review-mode');
    expect(component).not.toContain('Review pane');
    expect(component).not.toContain('>Overlay<');
    expect(component).not.toContain('backdrop-filter');
    expect(component).toContain('border: 0; border-radius: 10px');
    expect(component).toContain('background: color-mix(in srgb, var(--panel) 96%, var(--text) 4%)');
    expect(component).toContain("var(--text)");
    expect(component).toContain("fill: rgba(43, 124, 255, .12); stroke: none");
    expect(component).toContain("fill: rgba(43, 124, 255, .3)");
    expect(component).toContain("background: var(--panel)");
    expect(component).not.toContain('review-spinner');
    expect(component).not.toContain("background: rgba(18, 31, 50, .9)");
    expect(controller).toContain('documentModel.comments');
    expect(controller).toContain('documentModel.revisions.length');
    expect(controller).toContain('documentModel.commentAnchorRanges()');
    expect(controller).toContain('documentModel.revisionAnchorRanges()');
    expect(controller).toContain('resolveCommentAnchorRuns(anchor, candidateRuns)');
    expect(controller).toContain('resolveRevisionAnchorRuns(anchor, candidateRuns)');
    expect(controller).toContain("revision?.kind === 'deletion'");
    expect(controller).toContain('review-change-card');
    expect(controller).toContain('mergeReviewHighlightRuns(geometry)');
    expect(controller).toContain('previous.paragraphKey === paragraphKey');
    expect(controller).toContain("canvas.style.width = '100%'");
    expect(controller).toContain("canvas.style.height = 'auto'");
    expect(controller).toContain('new ResizeObserver(layoutMargin)');
    expect(controller).toContain('const upwardShift = Math.min(');
    expect(controller).toContain('const top = unshiftedTop - upwardShift');
    expect(controller).toContain('path.dataset.reviewId = entry.id');
    expect(controller).not.toContain('index % 3');
    expect(component).toContain('<strong>sample-1.docx</strong>');
    expect(component).toContain('samples/sample-1.docx');
    expect(component).not.toContain('review-sample.docx');
    expect(sampleCopy).toContain("packages/docx/public/demo/sample-1.docx");
    expect(sampleCopy).not.toContain("review-sample.docx");
  });
});
