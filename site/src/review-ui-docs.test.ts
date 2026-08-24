import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (path: string): string => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('DOCX review UI integration guide', () => {
  const page = read('./pages/review-ui.astro');

  it('starts with the built-in path and keeps low-level composition available', () => {
    expect(page).toContain('<h1>DOCX comments and tracked changes</h1>');
    expect(page).toContain('reads review data already stored in a DOCX file');
    for (const concept of ['Comments and threads', 'Tracked changes', 'Anchors', 'Rendered geometry']) {
      expect(page).toContain(concept);
    }
    expect(page).toContain('includes an optional read-only comment');
    expect(page).toContain('List, search, or export');
    expect(page).toContain('Margin or separate pane');
    expect(page).toContain('Highlights or markers');
    expect(page).toContain('showComments: true');
    expect(page).toContain('renderCommentCard(host, { comment, active, activate })');
    expect(page).not.toContain('commentRenderer');
  });

  it('maps concepts to public APIs and distinguishes transcript from anchored UI', () => {
    const example = read('./examples/review-margin/index.ts');
    for (const api of ['doc.comments', 'doc.revisions', 'commentAnchorRanges()', 'revisionAnchorRanges()', 'onTextRun', 'resolveCommentAnchorRuns()', 'resolveRevisionAnchorRuns()']) {
      expect(page).toContain(api);
    }
    expect(page).toContain('No page geometry');
    expect(page).toContain('Page geometry required');
    expect(page).toContain('Transcript path');
    expect(page).toContain('Anchored page path');
    expect(example).toContain('doc.commentAnchorRanges()');
    expect(example).toContain('doc.revisionAnchorRanges()');
    expect(example).toContain('onTextRun: (run) => runs.push(run)');
    expect(example).toContain('resolveCommentAnchorRuns(anchor, runs)');
    expect(example).toContain('resolveRevisionAnchorRuns(anchor, runs)');
    const apiReference = read('./lib/api-reference.ts');
    expect(apiReference).toContain('get comments(): DocComment[]');
    expect(apiReference).toContain('get revisions(): DocRevision[]');
    expect(apiReference).toContain('commentAnchorRanges(): readonly CommentAnchorRange[]');
    expect(apiReference).toContain('revisionAnchorRanges(): readonly RevisionAnchorRange[]');
    expect(apiReference).toContain('resolveCommentAnchorRuns()');
    expect(apiReference).toContain('collectPageRuns(index');
  });

  it('keeps the overview short while linking the complete page-aware example source', () => {
    const component = read('./components/ReviewGuideExample.astro');
    const core = read('./examples/review-margin/core.ts');
    const markup = read('./examples/review-margin/index.html');
    const controller = read('./examples/review-margin/index.ts');
    const styles = read('./examples/review-margin/styles.css');
    expect(page).toContain('<ReviewGuideExample />');
    expect(component).toContain("index.html?raw");
    expect(component).not.toContain("core.ts?raw");
    expect(component).not.toContain("index.ts?raw");
    expect(component).not.toContain("styles.css?raw");
    expect(component).toContain("import '../examples/review-margin/index'");
    expect(markup).toContain('data-review-example');
    expect(core).toContain('export async function renderReviewPage(');
    expect(core).toContain("isPositionHint: revision.kind === 'deletion' || revision.kind === 'moveFrom'");
    expect(component).toContain('href="/review-ui/source"');
    expect(component).toContain('Open complete example source');
    const sourcePage = read('./pages/review-ui/source.astro');
    expect(sourcePage).toContain("core.ts?raw");
    expect(sourcePage).toContain("index.ts?raw");
    expect(sourcePage).toContain("index.html?raw");
    expect(sourcePage).toContain("styles.css?raw");
    expect(sourcePage).toContain('<CodeTabs id="review-example-complete-source"');
    expect(controller).toContain('export async function mountReviewExample(root: HTMLElement, signal?: AbortSignal)');
    expect(controller).toContain('const pageIndex = Math.max(0, Math.min(requestedPage, doc.pageCount - 1))');
    expect(controller).toContain('await doc.renderPage(stage, pageIndex');
    expect(controller).toContain("Number(root.dataset.page ?? 0)");
    expect(controller).toContain('updatePage(pageIndex: number): Promise<void>');
    expect(controller).toContain('const destroy = (): void =>');
    expect(controller).toContain('request !== generation');
    expect(controller).toContain('function lineBands(');
    expect(controller).toContain('previous.transform === run.transform');
    expect(controller).toContain('gap <= Math.max(2, run.h * .4)');
    expect(styles).toContain('.review-example__margin');
    expect(styles).toContain('@container (max-width: 720px)');
    expect(styles).not.toContain('@media (max-width: 720px)');
    expect(component).not.toContain('<CodeTabs');
    expect(component).toContain('max-height: min(76vh, 720px)');
  });

  it('keeps only the essential accepted-final rule and delegates advanced details', () => {
    for (const term of ['Insertions', 'move destinations', 'Deletions', 'move sources']) {
      expect(page).toContain(term);
    }
    expect(page).toContain('accepted-final document');
    expect(page).toContain('nearby final-state position');
    expect(page).toContain('Canvas content needs an accessible transcript');
    expect(page).toContain('Destroy an owned <code>DocxDocument</code>');
    expect(page).toContain('href="/production#rendering-mode"');
    expect(page).not.toContain('href="/production#ownership"');
    expect(page).toContain('See the controller, HTML, and CSS used by the live example.');
    expect(page).toContain('Look up review records, anchor ranges, text runs, and method signatures.');
    expect(page).toContain('Choose main-thread or Worker rendering for your application.');
    expect(page).not.toContain('Manage a loaded document when it is shared by more than one view.');
    for (const detail of ['geometryFallback', 'storyInstance', 'sourceRunIndex', 'Device pixel ratio', 'renderPageToBitmap', 'useGoogleFonts', 'same-origin or return suitable CORS headers']) {
      expect(page).not.toContain(detail);
    }
  });

  it('renders all revision kinds, threads, fallback carets, linked controls, and live states', () => {
    const example = read('./examples/review-margin/index.ts');
    const markup = read('./examples/review-margin/index.html');
    for (const kind of ['insertion', 'deletion', 'moveFrom', 'moveTo']) {
      expect(example).toContain(kind);
    }
    expect(example).toContain('threads.rootOf(comment)');
    expect(example).toContain("orphaned reply");
    expect(example).toContain("item.marker === 'range' ? band.w : 3");
    expect(example).toContain("rect.setAttribute('aria-hidden', 'true')");
    expect(example).toContain("rect.setAttribute('focusable', 'false')");
    expect(example).not.toContain("rect.setAttribute('role', 'button')");
    expect(example).toContain("button.setAttribute('aria-controls', anchorIds.join(' '))");
    expect(example).toContain("root.addEventListener('mouseover'");
    expect(example).toContain("root.addEventListener('focusin'");
    expect(example).toContain('const controllers = mounted.filter');
    expect(example).toContain('updatePage(page).catch');
    expect(example).toContain('run.sourceRunIndex !== undefined && sameSource');
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

  it('is linked as durable DOCX guidance and included in the sitemap', () => {
    expect(read('./pages/docx.astro')).toContain('<ReviewDemo />');
    expect(read('./components/ReviewDemo.astro')).toContain('href="/review-ui"');
    expect(read('./pages/sitemap.xml.ts')).toContain("'/review-ui/'");
    expect(read('./pages/sitemap.xml.ts')).toContain("'/review-ui/source/'");
  });

  it('ships a real, consumer-owned review-margin demo on the DOCX page', () => {
    const docxPage = read('./pages/docx.astro');
    const component = read('./components/ReviewDemo.astro');
    const controller = read('./lib/review-demo.ts');
    const sampleCopy = read('../scripts/copy-samples.mjs');

    expect(docxPage).toContain('<ReviewDemo />');
    expect(docxPage.indexOf('<ReviewDemo />')).toBeGreaterThan(docxPage.indexOf('kind="masterdetail"'));
    expect(component).toContain('data-review-connectors');
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
    expect(controller).toContain('path.dataset.reviewId = entry.id');
    expect(controller).not.toContain('index % 3');
    expect(component).toContain('<strong>sample-1.docx</strong>');
    expect(component).toContain('samples/sample-1.docx');
    expect(component).not.toContain('review-sample.docx');
    expect(sampleCopy).toContain("packages/docx/public/demo/sample-1.docx");
    expect(sampleCopy).not.toContain("review-sample.docx");
  });
});
