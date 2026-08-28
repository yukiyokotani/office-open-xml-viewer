// Implementation snippets shown beside each per-format live demo. Generated
// from a small config so pptx/docx stay in sync; all checked against the API.

interface Cfg {
  Viewer: string;
  ScrollViewer: string;
  Engine: string;
  engineVariable: 'document' | 'presentation';
  fromEngine: 'fromDocument' | 'fromPresentation';
  sub: 'pptx' | 'docx';
  count: string;
  render: string;
  next: string;
  prev: string;
  go: string;
}

const pptx: Cfg = {
  Viewer: 'PptxViewer', ScrollViewer: 'PptxScrollViewer', Engine: 'PptxPresentation', engineVariable: 'presentation', fromEngine: 'fromPresentation', sub: 'pptx',
  count: 'slideCount', render: 'renderSlide', next: 'nextSlide', prev: 'prevSlide', go: 'goToSlide',
};
const docx: Cfg = {
  Viewer: 'DocxViewer', ScrollViewer: 'DocxScrollViewer', Engine: 'DocxDocument', engineVariable: 'document', fromEngine: 'fromDocument', sub: 'docx',
  count: 'pageCount', render: 'renderPage', next: 'nextPage', prev: 'prevPage', go: 'goToPage',
};

export interface DemoSnippets {
  demo: string;
  scroll: string;
  thumbnails: string;
  masterdetail: string;
}

function build(c: Cfg): DemoSnippets {
  return {
    demo: `import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// The built-in viewer tracks the current ${c.sub === 'pptx' ? 'slide' : 'page'} for you.
const viewer = new ${c.Viewer}(canvas, { width: 960, useGoogleFonts: true });
await viewer.load('/sample.${c.sub}');

nextBtn.addEventListener('click', () => viewer.${c.next}());
prevBtn.addEventListener('click', () => viewer.${c.prev}());`,

    scroll: `import { ${c.ScrollViewer} } from '@silurus/ooxml/${c.sub}';

// The built-in scroll viewer virtualizes a long ${c.sub === 'pptx' ? 'slide deck' : 'document'} for you.
const scroller = document.querySelector('#scroller') as HTMLElement;
const viewer = new ${c.ScrollViewer}(scroller, {
  enableTextSelection: true,
  useGoogleFonts: true,
});

await viewer.load('/sample.${c.sub}');

window.addEventListener('pagehide', (event) => {
  if (event.persisted) return;
  viewer.destroy();
});`,

    thumbnails: `import { ${c.Engine} } from '@silurus/ooxml/${c.sub}';

// Render each ${c.sub === 'pptx' ? 'slide' : 'page'} small, wire up navigation.
const ${c.engineVariable} = await ${c.Engine}.load('/sample.${c.sub}');

for (let i = 0; i < ${c.engineVariable}.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => open(i));
  grid.appendChild(thumb);
  await ${c.engineVariable}.${c.render}(thumb, i, { width: 320 });
}`,

    masterdetail: `import { ${c.Engine}, ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// Parse once, then lend the loaded engine to every view that needs it.
const ${c.engineVariable} = await ${c.Engine}.load('/sample.${c.sub}');

// A large preview on the right borrows the engine and cannot acquire another source.
const viewer = ${c.Viewer}.${c.fromEngine}(detailCanvas, ${c.engineVariable}, {
  width: 960,
  enableTextSelection: true,
});
await viewer.${c.go}(0);

// The thumbnail rail on the left renders from that same engine.
for (let i = 0; i < ${c.engineVariable}.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => viewer.${c.go}(i));  // jump the preview
  rail.appendChild(thumb);
  await ${c.engineVariable}.${c.render}(thumb, i, { width: 200 });
}

window.addEventListener('pagehide', (event) => {
  if (event.persisted) return;
  viewer.destroy();
  ${c.engineVariable}.destroy(); // borrowed engines remain caller-owned
});`,
  };
}

export const pptxSnippets = build(pptx);
export const docxSnippets = build(docx);

export const xlsxSheetWindowsSnippet = `import { XlsxSheetViewer, XlsxWorkbook } from '@silurus/ooxml/xlsx';

const workbook = await XlsxWorkbook.load('/sample.xlsx');
const viewers = new Map<Window, XlsxSheetViewer>();

async function openSheetInWindow(sheetIndex: number): Promise<void> {
  // Call this function directly from a click handler so popup blockers allow it.
  const popup = window.open('', '_blank', 'popup,width=1100,height=720,resizable=yes');
  if (!popup) throw new Error('The browser blocked the popup');

  const canvas = popup.document.createElement('canvas');
  canvas.style.cssText = 'display:block;width:100%;height:100%';
  popup.document.body.style.margin = '0';
  popup.document.body.appendChild(canvas);

  const viewer = XlsxSheetViewer.fromWorkbook(canvas, workbook);
  viewers.set(popup, viewer);
  popup.addEventListener('pagehide', (event) => {
    if (event.persisted) return;
    viewer.destroy();
    viewers.delete(popup);
  });

  await viewer.goToSheet(sheetIndex);
}

// Example: openSheetInWindow(1) from a sheet button's click handler.
window.addEventListener('pagehide', (event) => {
  if (event.persisted) return;
  viewers.forEach((viewer, popup) => {
    viewer.destroy();
    popup.close();
  });
  workbook.destroy();
});`;

export const xlsxSheetSnippet = `import { XlsxViewer } from '@silurus/ooxml/xlsx';

// XlsxViewer owns its canvas, sheet-tab bar and zoom slider — hand it a
// container element (not a canvas). Click-drag selects a range; Ctrl/Cmd+C
// copies it as TSV.
const container = document.getElementById('sheet') as HTMLElement;
const viewer = new XlsxViewer(container, { showZoomSlider: true });

await viewer.load('/sample.xlsx');`;
