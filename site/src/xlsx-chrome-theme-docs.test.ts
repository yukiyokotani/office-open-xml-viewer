import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const guide = readFileSync(new URL('./components/XlsxChromeThemeGuide.astro', import.meta.url), 'utf8');
const apiPage = readFileSync(new URL('./layouts/ApiPage.astro', import.meta.url), 'utf8');

describe('XLSX Viewer chrome theme guide', () => {
  it('documents every supported chrome CSS variable', () => {
    for (const property of [
      '--ooxml-xlsx-chrome-background',
      '--ooxml-xlsx-chrome-surface',
      '--ooxml-xlsx-chrome-surface-muted',
      '--ooxml-xlsx-chrome-text',
      '--ooxml-xlsx-chrome-text-muted',
      '--ooxml-xlsx-chrome-border',
      '--ooxml-xlsx-chrome-selection-background',
      '--ooxml-xlsx-chrome-accent',
      '--ooxml-xlsx-chrome-scrollbar-color',
      '--ooxml-xlsx-focus-ring',
    ]) {
      expect(guide).toContain(property);
    }
  });

  it('explains runtime switching and the authored-content boundary', () => {
    expect(guide).toContain("data-theme='dark'");
    expect(guide).toContain('without recreating or reloading the Viewer');
    expect(guide).toContain('authored cells, charts, pictures, or shapes');
  });

  it('is linked from the XLSX API table of contents', () => {
    expect(apiPage).toContain('<XlsxChromeThemeGuide />');
    expect(apiPage).toContain('href="#viewer-chrome-theme"');
    expect(guide).toContain('scroll-margin-top: 94px');
  });
});
