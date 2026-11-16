import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const page = readFileSync(new URL('./pages/xlsx.astro', import.meta.url), 'utf8');
const demos = readFileSync(new URL('./lib/demos.ts', import.meta.url), 'utf8');
const snippets = readFileSync(new URL('./lib/demo-snippets.ts', import.meta.url), 'utf8');
const styles = readFileSync(new URL('./styles/global.css', import.meta.url), 'utf8');

describe('XLSX multi-window sheet demo', () => {
  it('keeps the full workbook viewer first and presents multi-window composition as an extension', () => {
    expect(page.indexOf('kind="sheet"')).toBeLessThan(page.indexOf('kind="sheetWindows"'));
    expect(page).toContain('kind="sheetWindows"');
    expect(page).toContain('One workbook in several windows');
    expect(page).toContain('Parse the file once in the parent page');
  });

  it('parses once and opens borrowed sheet viewers in same-origin popup canvases', () => {
    expect(demos).toContain("import { XlsxSheetViewer, XlsxViewer, XlsxWorkbook, googleFonts } from '@silurus/ooxml-xlsx';");
    expect(demos).toContain("case 'sheetWindows': return mountSheetWindows(el, url);");
    expect(demos).toContain('XlsxWorkbook.load(url, { fonts: [googleFonts()] })');
    expect(demos).toContain("window.open(");
    expect(demos).toContain('const popupDocument = popup.document');
    expect(demos).toContain('const viewer = XlsxSheetViewer.fromWorkbook(canvas, workbook, {');
    expect(demos).not.toContain('showScrollbars: false');
    expect(demos).not.toContain('scrollbars=no');
    expect(demos).toContain('workbook,');
    expect(demos).toContain('viewer.goToSheet(index)');
    expect(demos).toContain('popup.addEventListener(\'pagehide\'');
  });

  it('shows copyable TypeScript with the same parse-once composition pattern', () => {
    expect(snippets).toContain('export const xlsxSheetWindowsSnippet');
    expect(snippets).toContain("import { XlsxSheetViewer, XlsxWorkbook } from '@silurus/ooxml/xlsx';");
    expect(snippets).toContain("await XlsxWorkbook.load('/sample.xlsx')");
    expect(snippets).toContain("window.open('', '_blank'");
    expect(snippets).toContain('XlsxSheetViewer.fromWorkbook(canvas, workbook)');
    expect(snippets).toContain('await viewer.goToSheet(sheetIndex)');
    expect(snippets).toContain('viewer.destroy()');
  });

  it('uses a responsive launcher that explains the single parse', () => {
    expect(styles).toContain('.demo-sheet-window-launcher {');
    expect(styles).toContain('.demo-sheet-window-summary {');
    expect(styles).toContain('.demo-sheet-window-row button {');
    expect(demos).toContain("parseBadge.textContent = '1× parse'");
    expect(demos).toContain("closeAll.textContent = 'Close all windows'");
    expect(demos).toContain('const openSessions = [...sessions.values()]');
    expect(demos.match(/if \(sessions\.get\(index\) !== session\) return;/g)).toHaveLength(2);
  });

  it('uses the site theme for launcher controls and keeps popup themes in sync', () => {
    expect(styles).toMatch(/\.demo-sheet-window-summary\s*\{[\s\S]*?background: var\(--bg-elev\);[\s\S]*?color: var\(--text\);/);
    expect(styles).toMatch(/\.demo-sheet-window-row\s*\{[\s\S]*?background: var\(--bg-elev\);[\s\S]*?color: var\(--text\);/);
    expect(styles).toMatch(/\.demo-sheet-window-row button\s*\{[\s\S]*?background: var\(--bg-elev-2\);[\s\S]*?color: var\(--text\);/);
    expect(styles).toMatch(/\.demo-sheet-window-identity span\s*\{[^}]*color: var\(--text-faint\);/);
    expect(styles).not.toContain('background: color-mix(in srgb, #fff 94%, var(--accent));');
    expect(demos).toContain('--sheet-window-bar-dark:#05090e');
    expect(demos).toContain("document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light'");
    expect(demos).toContain("attributeFilter: ['data-theme']");
    expect(demos).toContain('themeObserver.disconnect()');
  });

  it('stacks the workbook summary controls inside the card on mobile', () => {
    expect(styles).toMatch(/@media \(max-width: 620px\)[\s\S]*?\.demo-sheet-window-summary\s*\{[^}]*flex-direction: column;/);
    expect(styles).toMatch(/@media \(max-width: 620px\)[\s\S]*?\.demo-sheet-window-actions\s*\{[^}]*width: 100%;[^}]*margin-left: 0;/);
  });
});
