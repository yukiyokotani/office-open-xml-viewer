// Live demos for the per-format detail pages. Each demo mirrors one Storybook
// story (Demo / ScrollView / ThumbnailGrid / MasterDetail) using the real API.
// Demos are mounted lazily (on scroll) so a page with several of them doesn't
// parse the same file many times at once.
import { PptxPresentation, PptxViewer } from '@silurus/ooxml-pptx';
import { DocxDocument, DocxViewer } from '@silurus/ooxml-docx';
import { XlsxSheetViewer, XlsxViewer, XlsxWorkbook } from '@silurus/ooxml-xlsx';

export type Format = 'pptx' | 'docx' | 'xlsx';
export type DemoKind = 'demo' | 'scroll' | 'thumbnails' | 'masterdetail' | 'sheet' | 'sheetWindows';

const DPR = () => Math.min(typeof window !== 'undefined' ? window.devicePixelRatio : 1, 2);

// ── headless engine adapter (pptx slides / docx pages) ──────────────
type LoadedDoc = {
  kind: 'pptx';
  engine: PptxPresentation;
  count: number;
  render: (c: HTMLCanvasElement, i: number, width: number) => Promise<void>;
} | {
  kind: 'docx';
  engine: DocxDocument;
  count: number;
  render: (c: HTMLCanvasElement, i: number, width: number) => Promise<void>;
};

async function loadDoc(format: Format, url: string): Promise<LoadedDoc> {
  if (format === 'pptx') {
    const presentation = await PptxPresentation.load(url, { useGoogleFonts: true });
    return {
      kind: 'pptx',
      engine: presentation,
      count: presentation.slideCount,
      render: (c, i, width) => presentation.renderSlide(c, i, { width, dpr: DPR() }),
    };
  }
  const document = await DocxDocument.load(url, { useGoogleFonts: true });
  return {
    kind: 'docx',
    engine: document,
    count: document.pageCount,
    render: (c, i, width) => document.renderPage(c, i, { width, dpr: DPR() }),
  };
}

// ── viewer adapter (pptx / docx built-in viewers) ───────────────────
type ViewerCtl = {
  load: (url: string) => Promise<void>;
  go: (i: number) => Promise<void>;
  next: () => Promise<void>;
  prev: () => Promise<void>;
  index: () => number;
  count: () => number;
};

type BorrowedViewerCtl = Omit<ViewerCtl, 'load'> & { destroy: () => void };

function makeViewer(format: Format, canvas: HTMLCanvasElement, width: number): ViewerCtl {
  // NB: no enableTextSelection here. The demo canvases are downscaled with CSS
  // (.demo-page width:100%/height:auto) to fit the card, but the viewer's text
  // overlay is sized to the un-scaled page — leaving it on would inflate the
  // scroll area and add a big empty gap below the page. Text selection is shown
  // in the API reference instead.
  if (format === 'pptx') {
    const v = new PptxViewer(canvas, { width, useGoogleFonts: true });
    return {
      load: (u) => v.load(u), go: (i) => v.goToSlide(i), next: () => v.nextSlide(),
      prev: () => v.prevSlide(), index: () => v.slideIndex, count: () => v.slideCount,
    };
  }
  const v = new DocxViewer(canvas, { width, dpr: DPR(), useGoogleFonts: true });
  return {
    load: (u) => v.load(u), go: (i) => v.goToPage(i), next: () => v.nextPage(),
    prev: () => v.prevPage(), index: () => v.currentPage, count: () => v.pageCount,
  };
}

function makeBorrowedViewer(doc: LoadedDoc, canvas: HTMLCanvasElement, width: number): BorrowedViewerCtl {
  if (doc.kind === 'pptx') {
    const viewer = PptxViewer.fromPresentation(canvas, doc.engine, { width });
    return {
      go: (i) => viewer.goToSlide(i), next: () => viewer.nextSlide(), prev: () => viewer.prevSlide(),
      index: () => viewer.slideIndex, count: () => viewer.slideCount, destroy: () => viewer.destroy(),
    };
  }
  const viewer = DocxViewer.fromDocument(canvas, doc.engine, { width, dpr: DPR() });
  return {
    go: (i) => viewer.goToPage(i), next: () => viewer.nextPage(), prev: () => viewer.prevPage(),
    index: () => viewer.currentPage, count: () => viewer.pageCount, destroy: () => viewer.destroy(),
  };
}

const UNIT = (f: Format) => (f === 'pptx' ? 'Slide' : 'Page');

// ── public entry ────────────────────────────────────────────────────
export function mountDemoInto(el: HTMLElement, kind: DemoKind, format: Format, url: string): void {
  el.innerHTML = '';
  if (format === 'xlsx') {
    switch (kind) {
      case 'sheetWindows': return mountSheetWindows(el, url);
      default: return mountSheet(el, url);
    }
  }
  switch (kind) {
    case 'scroll': return mountScroll(el, format, url);
    case 'thumbnails': return mountThumbnails(el, format, url);
    case 'masterdetail': return mountMasterDetail(el, format, url);
    default: return mountDemo(el, format, url);
  }
}

interface SheetWindowSession {
  readonly popup: Window;
  readonly viewer: Omit<XlsxSheetViewer, 'load'>;
  readonly themeObserver: MutationObserver;
}

// Excel — parse once in the parent and project borrowed sheet viewers into
// same-origin popup canvases. Each child owns view state; the parent owns the
// workbook, archive cache, and worker lifecycle.
function mountSheetWindows(el: HTMLElement, url: string): void {
  const launcher = document.createElement('div');
  launcher.className = 'demo-sheet-window-launcher';
  el.appendChild(launcher);
  const st = status(launcher, 'Parsing workbook once…');

  XlsxWorkbook.load(url, { useGoogleFonts: true })
    .then((workbook) => {
      const sessions = new Map<number, SheetWindowSession>();
      const summary = document.createElement('div');
      summary.className = 'demo-sheet-window-summary';
      const summaryCopy = document.createElement('div');
      const summaryTitle = document.createElement('strong');
      summaryTitle.textContent = 'Workbook parsed once';
      const summaryDetail = document.createElement('span');
      summaryDetail.textContent = `${workbook.sheetCount} sheets · shared archive, cache and worker`;
      summaryCopy.append(summaryTitle, summaryDetail);
      const parseBadge = document.createElement('span');
      parseBadge.className = 'demo-sheet-window-badge';
      parseBadge.textContent = '1× parse';
      const summaryActions = document.createElement('div');
      summaryActions.className = 'demo-sheet-window-actions';
      const closeAll = document.createElement('button');
      closeAll.type = 'button';
      closeAll.textContent = 'Close all windows';
      closeAll.disabled = true;
      summaryActions.append(parseBadge, closeAll);
      summary.append(summaryCopy, summaryActions);

      const list = document.createElement('div');
      list.className = 'demo-sheet-window-list';
      const sheetButtons: HTMLButtonElement[] = [];
      const popupError = document.createElement('p');
      popupError.className = 'demo-sheet-window-error';
      popupError.hidden = true;
      workbook.sheetNames.forEach((name, index) => {
        const row = document.createElement('div');
        row.className = 'demo-sheet-window-row';
        const identity = document.createElement('div');
        identity.className = 'demo-sheet-window-identity';
        const number = document.createElement('span');
        number.textContent = String(index + 1).padStart(2, '0');
        const label = document.createElement('strong');
        label.textContent = name;
        identity.append(number, label);

        const open = document.createElement('button');
        open.type = 'button';
        open.textContent = 'Open in window \u2197\uFE0E';
        sheetButtons.push(open);
        open.addEventListener('click', () => {
          const existing = sessions.get(index);
          if (existing && !existing.popup.closed) {
            existing.popup.focus();
            return;
          }

          const popup = window.open(
            '',
            '_blank',
            'popup=yes,width=1100,height=720,resizable=yes',
          );
          if (!popup) {
            popupError.textContent = 'The browser blocked the popup. Allow popups and try again.';
            popupError.hidden = false;
            return;
          }
          popupError.hidden = true;

          const popupDocument = popup.document;
          popupDocument.title = `${name} · XLSX Sheet Viewer`;
          const popupStyle = popupDocument.createElement('style');
          popupStyle.textContent =
            ':root{color-scheme:light;--sheet-window-page:#eef2f6;--sheet-window-bar:#172333;' +
            '--sheet-window-bar-dark:#05090e;--sheet-window-bar-border:#34445a;' +
            '--sheet-window-bar-meta:#9fe2c4}' +
            ':root[data-theme="dark"]{color-scheme:dark;--sheet-window-page:#030508;' +
            '--sheet-window-bar:var(--sheet-window-bar-dark);--sheet-window-bar-border:#18232f;' +
            '--sheet-window-bar-meta:#72d9a8}' +
            'html,body{width:100%;height:100%;margin:0;overflow:hidden;background:var(--sheet-window-page);' +
            'font-family:Inter,ui-sans-serif,system-ui,sans-serif}' +
            '.sheet-window{height:100%;display:grid;grid-template-rows:48px minmax(0,1fr)}' +
            '.sheet-window__bar{display:flex;align-items:center;gap:12px;padding:0 16px;' +
            'background:var(--sheet-window-bar);color:#fff;border-bottom:1px solid var(--sheet-window-bar-border)}' +
            '.sheet-window__bar strong{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}' +
            '.sheet-window__bar span{margin-left:auto;font:11px ui-monospace,monospace;color:var(--sheet-window-bar-meta)}' +
            '.sheet-window__viewport{position:relative;min-height:0;background:#fff}' +
            '.sheet-window__canvas{display:block;width:100%;height:100%;background:#fff}' +
            '.sheet-window__loading{position:absolute;inset:0;display:grid;place-items:center;' +
            'z-index:10;background:#fff;color:#59687a;font-size:13px}';
          popupDocument.head.appendChild(popupStyle);

          const syncPopupTheme = () => {
            popupDocument.documentElement.dataset.theme =
              document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
          };
          syncPopupTheme();
          const themeObserver = new MutationObserver(syncPopupTheme);
          themeObserver.observe(document.documentElement, {
            attributes: true,
            attributeFilter: ['data-theme'],
          });

          const shell = popupDocument.createElement('main');
          shell.className = 'sheet-window';
          const toolbar = popupDocument.createElement('header');
          toolbar.className = 'sheet-window__bar';
          const popupTitle = popupDocument.createElement('strong');
          popupTitle.textContent = name;
          const shared = popupDocument.createElement('span');
          shared.textContent = `Shared workbook · Sheet ${index + 1} / ${workbook.sheetCount}`;
          toolbar.append(popupTitle, shared);
          const viewport = popupDocument.createElement('div');
          viewport.className = 'sheet-window__viewport';
          const canvas = popupDocument.createElement('canvas');
          canvas.className = 'sheet-window__canvas';
          const loading = popupDocument.createElement('div');
          loading.className = 'sheet-window__loading';
          loading.textContent = 'Rendering from the shared workbook…';
          viewport.append(canvas, loading);
          shell.append(toolbar, viewport);
          popupDocument.body.replaceChildren(shell);

          const renderState: { error: Error | null } = { error: null };
          const viewer = XlsxSheetViewer.fromWorkbook(canvas, workbook, {
            onError: (error) => {
              renderState.error = error;
              loading.textContent = err(error);
            },
          });
          const session = { popup, viewer, themeObserver };
          sessions.set(index, session);
          closeAll.disabled = false;
          open.textContent = 'Rendering…';
          open.disabled = true;

          popup.addEventListener('pagehide', (event) => {
            if (event.persisted) return;
            themeObserver.disconnect();
            viewer.destroy();
            if (sessions.get(index) === session) sessions.delete(index);
            open.textContent = 'Open in window \u2197\uFE0E';
            open.disabled = false;
            closeAll.disabled = sessions.size === 0;
          });

          viewer.goToSheet(index)
            .then(() => {
              if (sessions.get(index) !== session) return;
              if (renderState.error) open.textContent = 'View error \u2197\uFE0E';
              else {
                loading.remove();
                open.textContent = 'Focus window \u2197\uFE0E';
              }
              open.disabled = false;
            })
            .catch((error) => {
              if (sessions.get(index) !== session) return;
              loading.textContent = err(error);
              open.textContent = 'View error \u2197\uFE0E';
              open.disabled = false;
            });
        });

        row.append(identity, open);
        list.appendChild(row);
      });

      closeAll.addEventListener('click', () => {
        const openSessions = [...sessions.values()];
        sessions.clear();
        openSessions.forEach(({ popup, viewer, themeObserver }) => {
          themeObserver.disconnect();
          viewer.destroy();
          popup.close();
        });
        sheetButtons.forEach((button) => {
          button.textContent = 'Open in window \u2197\uFE0E';
          button.disabled = false;
        });
        closeAll.disabled = true;
      });

      st.remove();
      launcher.append(summary, popupError, list);
      window.addEventListener('pagehide', (event) => {
        if (event.persisted) return;
        sessions.forEach(({ popup, viewer, themeObserver }) => {
          themeObserver.disconnect();
          viewer.destroy();
          popup.close();
        });
        sessions.clear();
        workbook.destroy();
      });
    })
    .catch((error) => { st.textContent = err(error); });
}

function status(el: HTMLElement, text: string): HTMLDivElement {
  const d = document.createElement('div');
  d.className = 'demo-status';
  d.setAttribute('role', 'status');
  d.setAttribute('aria-live', 'polite');
  const circle = document.createElement('span');
  circle.className = 'demo-progress-circle';
  circle.setAttribute('aria-hidden', 'true');
  const label = document.createElement('span');
  label.textContent = text;
  d.append(circle, label);
  el.appendChild(d);
  return d;
}

// Demo — single viewer with built-in navigation
function mountDemo(el: HTMLElement, format: Format, url: string): void {
  const bar = document.createElement('div');
  bar.className = 'demo-bar';
  const prev = button('‹');
  const next = button('›');
  const info = document.createElement('span');
  info.className = 'demo-info';
  info.textContent = 'Loading…';
  bar.append(prev, info, next);

  const stage = document.createElement('div');
  stage.className = 'demo-stage';
  const canvas = document.createElement('canvas');
  canvas.className = 'demo-page';
  canvas.hidden = true;
  stage.appendChild(canvas);
  const st = status(stage, 'Parsing document…');
  el.append(bar, stage);

  const v = makeViewer(format, canvas, 960);
  const viewerWrapper = canvas.parentElement as HTMLDivElement;
  viewerWrapper.hidden = true;
  const sync = () => {
    const n = v.count();
    info.textContent = n ? `${UNIT(format)} ${v.index() + 1} / ${n}` : 'Loading…';
    prev.disabled = v.index() <= 0;
    next.disabled = v.index() >= n - 1;
  };
  prev.addEventListener('click', () => void v.prev().then(sync));
  next.addEventListener('click', () => void v.next().then(sync));
  v.load(url).then(() => {
    canvas.hidden = false;
    viewerWrapper.hidden = false;
    st.remove();
    sync();
  }).catch((e) => {
    st.textContent = err(e);
    info.textContent = 'Failed';
  });
}

// ScrollView — every page stacked on a backdrop
function mountScroll(el: HTMLElement, format: Format, url: string): void {
  const sc = document.createElement('div');
  sc.className = 'demo-scroll';
  el.appendChild(sc);
  const st = status(sc, 'Parsing…');
  loadDoc(format, url).then(async (doc) => {
    st.remove();
    for (let i = 0; i < doc.count; i++) {
      const c = document.createElement('canvas');
      c.className = 'demo-page';
      sc.appendChild(c);
      await doc.render(c, i, 1100);
    }
  }).catch((e) => { st.textContent = err(e); });
}

// ThumbnailGrid — every page at a glance
function mountThumbnails(el: HTMLElement, format: Format, url: string): void {
  const grid = document.createElement('div');
  grid.className = 'demo-grid';
  el.appendChild(grid);
  const st = status(grid, 'Rendering thumbnails…');
  loadDoc(format, url).then(async (doc) => {
    for (let i = 0; i < doc.count; i++) {
      const cell = document.createElement('div');
      cell.className = 'demo-cell';
      const c = document.createElement('canvas');
      c.className = 'demo-page';
      const cap = document.createElement('span');
      cap.className = 'demo-cap';
      cap.textContent = `${UNIT(format)} ${i + 1}`;
      cell.append(c, cap);
      grid.appendChild(cell);
      await doc.render(c, i, 320);
    }
    st.remove();
  }).catch((e) => { st.textContent = err(e); });
}

// MasterDetail — thumbnail rail + large preview
function mountMasterDetail(el: HTMLElement, format: Format, url: string): void {
  const layout = document.createElement('div');
  layout.className = 'demo-md';
  const rail = document.createElement('div');
  rail.className = 'demo-rail';
  const detail = document.createElement('div');
  detail.className = 'demo-detail';
  const detailCanvas = document.createElement('canvas');
  detailCanvas.className = 'demo-page';
  detailCanvas.hidden = true;
  detail.appendChild(detailCanvas);
  const st = status(detail, 'Parsing document…');
  layout.append(rail, detail);
  el.appendChild(layout);

  loadDoc(format, url)
    .then(async (doc) => {
      const viewer = makeBorrowedViewer(doc, detailCanvas, 960);
      const detailViewerWrapper = detailCanvas.parentElement as HTMLDivElement;
      detailViewerWrapper.hidden = true;
      window.addEventListener('pagehide', (event) => {
        if (event.persisted) return;
        viewer.destroy();
        doc.engine.destroy();
      });
      await viewer.go(0);
      detailCanvas.hidden = false;
      detailViewerWrapper.hidden = false;
      st.remove();
      const cells: HTMLDivElement[] = [];
      const select = async (i: number) => {
        cells.forEach((c, k) => c.classList.toggle('active', k === i));
        await viewer.go(i);
        // Reset scroll so the new page is shown from its top, not wherever the
        // previous page was scrolled to.
        detail.scrollTop = 0;
      };
      for (let i = 0; i < doc.count; i++) {
        const cell = document.createElement('div');
        cell.className = 'demo-rail-cell';
        const c = document.createElement('canvas');
        c.className = 'demo-page';
        cell.appendChild(c);
        cell.addEventListener('click', () => void select(i));
        rail.appendChild(cell);
        cells.push(cell);
        await doc.render(c, i, 200);
      }
      cells[0]?.classList.add('active');
    })
    .catch((e) => { st.textContent = err(e); });
}

// Excel — the full viewer (sheets + selection + zoom)
function mountSheet(el: HTMLElement, url: string): void {
  const host = document.createElement('div');
  host.className = 'demo-xlsx';
  el.appendChild(host);
  const viewer = new XlsxViewer(host, { useGoogleFonts: true, showZoomSlider: true });
  viewer.load(url).catch((error) => host.setAttribute('data-error', err(error)));
}

function button(label: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = 'demo-btn';
  b.textContent = label;
  b.disabled = true;
  return b;
}
function err(e: unknown): string {
  return `Failed: ${e instanceof Error ? e.message : String(e)}`;
}
