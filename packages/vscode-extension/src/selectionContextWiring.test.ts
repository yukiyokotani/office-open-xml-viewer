import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const bootstrap = readFileSync(new URL('./webview/bootstrap.ts', import.meta.url), 'utf8');
const editorHost = readFileSync(new URL('./providers/baseEditor.ts', import.meta.url), 'utf8');
const mcpActivation = readFileSync(new URL('./mcp/index.ts', import.meta.url), 'utf8');
const extensionManifest = readFileSync(new URL('../package.json', import.meta.url), 'utf8');
const extensionReadme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
const mcpReadme = readFileSync(new URL('../../mcp-server/README.md', import.meta.url), 'utf8');

function initializer(name: 'Xlsx' | 'Docx' | 'Pptx'): string {
  const source = bootstrap.match(
    new RegExp(`async function init${name}\\([\\s\\S]*?(?=\\n// ──|$)`),
  )?.[0];
  expect(source).toBeDefined();
  return source!;
}

describe('VS Code selection-context wiring', () => {
  it('publishes canonical context from every Viewer instead of raw browser events', () => {
    expect(initializer('Xlsx')).toContain('viewer.getSelectionContext(');
    expect(initializer('Docx')).toContain('onSelectionContextChange');
    expect(initializer('Pptx')).toContain('onSelectionContextChange');
    expect(bootstrap).toContain("type: 'selection-context'");
  });

  it('enables read-only element focus for AI context in every format', () => {
    expect(initializer('Xlsx')).toContain('enableElementSelection: true');
    expect(initializer('Docx')).toContain('enableElementSelection: true');
    expect(initializer('Pptx')).toContain('enableElementSelection: true');
  });

  it('does not contribute a dead copy command or orphaned webview protocol', () => {
    expect(extensionManifest).not.toContain('ooxmlViewer.copySelection');
    expect(bootstrap).not.toContain('copy-request');
    expect(editorHost).not.toContain("msg.type === 'copy'");
  });

  it('echoes a host-issued session so a stale webview cannot restore old context', () => {
    expect(bootstrap).toContain('selectionSession');
    expect(bootstrap).toContain("type: 'webview-ready', selectionSession");
    expect(bootstrap).toContain('msg.selectionSession !== selectionSession');
    expect(editorHost).toContain('msg.selectionSession !== view.selectionSession');
    // One init per webview document of that session (session + instance).
    expect(editorHost).toContain(
      'view.initialized?.session === session && view.initialized.instance === instance',
    );
  });

  it('makes the explicit MCP enable command work for previews outside the workspace', () => {
    expect(mcpActivation).toContain("await cfg.update('enabled', 'always'");
    expect(mcpActivation).toContain('GitHub Copilot Chat in Agent mode');
    expect(mcpActivation).not.toContain('Copilot, Claude, etc.');
    expect(extensionManifest).toContain("GitHub Copilot Chat in Agent mode");
    expect(extensionManifest).toContain('Claude Code and Codex require separate MCP configuration');
    expect(extensionManifest).not.toContain('Copilot, Claude, etc.');
  });

  it('re-evaluates installation when MCP settings change', () => {
    expect(mcpActivation).toContain("e.affectsConfiguration('ooxmlViewer.mcpServer')");
    expect(mcpActivation).toContain('await promptIfNeeded(context, provider)');
  });

  it('does not retain the pre-0.77 loadFailed completion channel', () => {
    expect(bootstrap).not.toContain('loadFailed');
  });

  it('documents element selection for every preview format', () => {
    for (const readme of [extensionReadme, mcpReadme]) {
      expect(readme).toContain('chart, picture, or shape in any format');
      expect(readme).not.toMatch(/PPTX element|PowerPoint shape|In PPTX/);
    }
  });

  it('describes remote identity and opt-in AI disclosure accurately', () => {
    for (const readme of [extensionReadme, mcpReadme]) {
      expect(readme).not.toContain('document.uri');
      expect(readme).not.toContain('sanitized URI');
    }
    expect(extensionReadme).toContain('exposes requested file or active-preview context');
    expect(extensionManifest).toContain('Optional MCP integration can share requested context');
  });
});
