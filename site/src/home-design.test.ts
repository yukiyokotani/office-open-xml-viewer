import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const showcase = readFileSync(new URL('./components/LiveShowcase.astro', import.meta.url), 'utf8');
const formatTabs = readFileSync(new URL('./components/FormatTabs.astro', import.meta.url), 'utf8');
const live = readFileSync(new URL('./lib/live.ts', import.meta.url), 'utf8');
const capabilities = readFileSync(new URL('./components/Capabilities.astro', import.meta.url), 'utf8');
const home = readFileSync(new URL('./pages/index.astro', import.meta.url), 'utf8');

describe('official-site home design', () => {
  it('keeps the format switcher inside the preview toolbar', () => {
    expect(showcase).toMatch(/<div class="panel-head">\s*<FormatTabs selected="docx"/);
    expect(formatTabs).toContain(".tab[aria-pressed='true'] {");
    expect(formatTabs).toContain('background: var(--border);');
    expect(formatTabs).not.toMatch(/\.tab\[aria-pressed='true'\][\s\S]*?var\(--signal\)/);
    expect(formatTabs).not.toContain('tab-dot');
    expect(formatTabs).not.toContain('--fmt');
  });

  it('uses circular feature bullets and one-pixel separators', () => {
    expect(capabilities).not.toContain("content: '—'");
    expect(capabilities).toContain("border-radius: 50%; background: var(--text);");
    expect(capabilities).toContain('border-top: 1px solid var(--border-bright);');
    expect(capabilities).toContain('border-bottom: 1px solid var(--border);');
    expect(capabilities).not.toMatch(/border-(?:top|bottom): [2-9]px/);
  });

  it('forces external-link arrows to use text presentation on mobile', () => {
    expect(capabilities).toContain('Full support matrix &#x2197;&#xFE0E;');
    expect(home).toContain('GitHub &#x2197;&#xFE0E;');
  });

  it('reserves format colours for format-heading underlines', () => {
    expect(capabilities).toContain('background: linear-gradient(transparent 68%, var(--c) 68%, var(--c) 88%, transparent 88%);');
    expect(capabilities).toContain('border-bottom: 1px solid currentColor;');
    expect(capabilities).not.toContain('border-bottom: 1px solid var(--c);');
  });

  it('stacks every main section introduction without forced heading breaks', () => {
    expect(home).toContain('<h2 class="section-title">Live renderer examples.</h2>');
    expect(home).toContain('<h2 class="section-title">Browser-based OOXML rendering.</h2>');
    expect(home).toContain('<h2 class="section-title">Format support at a glance.</h2>');
    expect(home).toContain('<h2 class="section-title">Release and migration notices.</h2>');
    expect(home).toContain('.grid-intro { display: block; margin-bottom: 52px; }');
    expect(home).toContain('.news-head { display: block; margin-bottom: 56px; }');
    expect(home).toContain('.news-head .section-title { max-width: none; margin: 18px 0 12px; }');
    expect(home).not.toMatch(/(?:Live renderer|Browser-based|Format support|Release and)<br \/>/);
  });

  it('places release notices after the product content', () => {
    expect(home.indexOf('id="announcements"')).toBeGreaterThan(home.indexOf('id="capabilities"'));
    expect(home.indexOf('id="announcements"')).toBeLessThan(home.indexOf('<footer class="footer">'));
  });

  it('keeps the hero headline at documentation-site scale', () => {
    expect(home).toContain('font-size: clamp(44px, 6.3vw, 82px);');
    expect(home).toContain('font-size: clamp(42px, 12vw, 60px);');
    expect(home).not.toContain('font-size: clamp(56px, 8.6vw, 116px);');
  });

  it('separates the hero copy and icon without a vertical rule', () => {
    expect(home).not.toMatch(/\.hero-art \{[^}]*border-left/);
  });

  it('lets the live showcase blend into the dark page background', () => {
    expect(home).toContain('.showcase-section { background: var(--showcase-bg); }');
  });

  it('keeps one lazily-created viewer per format instead of reparsing on every tab switch', () => {
    expect(showcase).toContain('data-format-pane={format}');
    expect(showcase).toContain('const cache = new Map<string, CachedViewer>();');
    expect(showcase).toContain('cache.get(id)');
    expect(showcase).toContain('cache.set(id, entry)');
    expect(showcase).toContain('function warmFormats(activeId: string): void');
    expect(showcase).not.toContain('current?.destroy()');
    expect(showcase).toContain('for (const entry of cache.values()) entry.controller.destroy();');
  });

  it('preserves the live viewers across back-forward cache restores', () => {
    expect(showcase).toContain("window.addEventListener('pagehide', (event) => {");
    expect(showcase).toContain('if (event.persisted) return;');
    expect(showcase).not.toMatch(/pagehide[\s\S]*?\{ once: true \}/);
    expect(showcase).toContain("window.addEventListener('pageshow', (event) => {");
    expect(showcase).toContain('if (!event.persisted || !loaded) return;');
    expect(showcase).toContain('cache.get(loaded)?.controller.activate?.();');
  });

  it('uses selectable scroll viewers for the DOCX and PPTX samples', () => {
    expect(live).toContain('DocxScrollViewer.fromDocument');
    expect(live).toContain('PptxScrollViewer.fromPresentation');
    expect(live.match(/enableTextSelection:\s*true/g)).toHaveLength(2);
  });

  it('starts paginated samples at the top without removing their desk margin', () => {
    expect(live).not.toMatch(/paddingTop:\s*0/);
    expect(live.match(/gap:\s*26/g)).toHaveLength(2);
  });

  it('centres a smaller hero icon and aligns the second-row format links on mobile', () => {
    expect(home).toContain('.hero-art { grid-column: 1 / 13; min-height: 280px; margin-top: 16px; opacity: 1; }');
    expect(home).toContain('.hero-icon { width: min(62%, 260px); }');
    expect(home).toMatch(/@media \(max-width: 680px\)[\s\S]*?\.hero-art \{[^}]*opacity: 1;/);
    expect(home).toContain('.format-rail a + a:nth-child(odd) { padding-left: 0; }');
    expect(home).toContain('.format-rail a:nth-child(even) { padding-left: 18px; }');
  });

  it('keeps cached mobile previews constrained to the showcase width', () => {
    expect(showcase).toContain('min-width: 0;');
    expect(showcase).toMatch(/@media \(max-width: 620px\)[\s\S]*?\.panel-body \{ height: min\(68vh, 520px\); min-height: 340px; \}/);
  });

  it('caps paginated samples on wide screens while leaving XLSX full width', () => {
    expect(live).toContain('const MAX_SAMPLE_WIDTH = 880;');
    expect(live).toContain('MAX_SAMPLE_WIDTH / (presentation.slideWidth / EMU_PER_PX)');
    expect(live).toContain('document.pageSize(index).widthPt * PT_TO_PX');
    expect(showcase).toContain('max-width: none;');
  });
});
