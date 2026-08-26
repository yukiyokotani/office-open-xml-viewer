import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { sitemapPaths } from './pages/sitemap.xml';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');
const page = source('./pages/selection-context.astro');
const formatTabs = source('./components/FormatTabs.astro');

describe('official selection-context guide', () => {
  it('distinguishes UI selection state from bounded content handoff', () => {
    expect(page).toContain('Selection state is the UI authority');
    expect(page).toContain('Selection context is a detached snapshot');
    expect(page).toContain('getSelectionContext()');
    expect(page).toContain('onSelectionContextChange');
    expect(page).toContain('enableTextSelection: true');
    expect(page).toContain('enableElementSelection: true');
    expect(page).toContain('non-editable outline');
    expect(page).toContain('format');
    expect(page).toContain('kind');
    expect(page).toContain("import { Code } from 'astro:components'");
    expect(page).toContain('code={browserIntegrationSnippet}');
    expect(page).toContain('code={contextMenuSnippet}');
    expect(page).toContain('lang="ts"');
    expect(page).toContain('themes={codeThemes}');
    expect(page).toContain('defaultColor={false}');
  });

  it('documents every current context kind and the VS Code MCP outcomes', () => {
    for (const context of [
      'docx / text', 'docx / element', 'xlsx / range', 'xlsx / element',
      'pptx / text', 'pptx / element',
    ]) {
      expect(page).toContain(context);
    }
    expect(page).toContain('ooxml_get_active_context');
    expect(page).toContain('context: null');
    expect(page).toContain('selection: null');
    expect(page).toContain('available: false');
    expect(page).toContain('there is no active');
    expect(page).toContain('inspect <code>reason</code>');
    expect(page).toContain('Remote VS Code documents');
    expect(page).toContain('only a document name');
    expect(page).toContain('no local <code>document.path</code>');
    expect(page).toContain('untrusted input');
    expect(page).toContain('does not edit the Office file');
    expect(page).toContain('GitHub Copilot Chat in Agent mode');
    expect(page).toContain('Claude Code and Codex VS Code extensions use separate MCP');
    expect(page).toContain('process receives no active Viewer selection');
    expect(page).toContain('MCP: List Servers');
    expect(page).toContain('does not add its own chat panel');
  });

  it('demonstrates the bounded handoff before an AI integration', () => {
    expect(page).toContain('data-selection-context-demo');
    expect(page).toContain('data-selection-context-output');
    for (const format of ['xlsx', 'docx', 'pptx']) {
      expect(formatTabs).toContain(`'${format}'`);
    }
    expect(page).toContain('`sample-1.${format}`');
    expect(page).toContain('<FormatTabs selected="xlsx"');
    expect(source('./components/LiveShowcase.astro')).toContain('<FormatTabs selected="docx"');
    expect(formatTabs).toContain('data-format-tab={format}');
    expect(page).toContain("import { DocxScrollViewer } from '@silurus/ooxml-docx'");
    expect(page).toContain("import { PptxScrollViewer } from '@silurus/ooxml-pptx'");
    expect(page).toContain('enableTextSelection: true');
    expect(page).toContain('enableElementSelection: true');
    expect(page.match(/enableElementSelection: true/g)?.length).toBeGreaterThanOrEqual(5);
    expect(page).not.toContain("window.open('', 'ooxml-selection-context-inspector'");
    expect(page).toContain('onSelectionContextChange(context)');
    expect(page).not.toContain('onSelectionStateChange: refresh');
    expect(page).toContain('maxTextCharacters: 4_096');
    expect(page).toContain('JSON.stringify(latestContext, null, 2)');
    expect(page).toContain('grid-template-columns: minmax(0, 7fr) minmax(0, 5fr)');
    expect(page).toContain('background: radial-gradient(120% 80% at 50% 0%, var(--preview-top), var(--preview-bottom) 70%)');
    expect(page).not.toContain("background: '#eef2f7'");
    expect(page).not.toContain("viewer.setSelection('B2:D6')");
    expect(page.match(/updateContext\(null\);/g)?.length).toBeGreaterThanOrEqual(3);
    expect(page).toContain('onContextMenu: async ({ originalEvent, getContext })');
    expect(page).toContain('originalEvent.preventDefault()');
    expect(page).toContain('There is no separate <code>onClick</code> option.');
    expect(page).toMatch(/add a native\s*<code>click<\/code> listener to a stable application-owned host/);
    expect(page).toContain('text-overlay clicks, whose target is a sibling of the canvas');
    expect(page).toContain('const context = await getContext()');
    expect(page).toContain('same memoized Promise');
    expect(page).toContain('Viewer installs no listener');
    expect(page).not.toContain('void context.then');
  });

  it('links the guide from cross-format decisions and every format API', () => {
    expect(page).toContain('canonicalPath="/selection-context/"');
    expect(sitemapPaths).toContain('/selection-context/');
    expect(source('./components/Nav.astro')).not.toContain('href="/selection-context"');
    expect(source('./components/SiteFooter.astro')).not.toContain('href="/selection-context"');
    expect(source('./pages/production.astro')).toContain('href="/selection-context"');
    expect(source('./components/ApiReference.astro')).toContain('href="/selection-context"');
    expect(source('./pages/review-ui.astro')).toContain('href="/selection-context"');
    expect(page).not.toContain('XLSX 0.77 migration');
    expect(source('./lib/announcements.ts')).toContain("slug: 'v077-migration-guide'");
    expect(source('./lib/announcements.ts')).toContain('setSelection(), selectionState');
  });
});
