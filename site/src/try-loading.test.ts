import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const source = readFileSync(new URL('./pages/try.astro', import.meta.url), 'utf8');
const renderer = readFileSync(new URL('./lib/try.ts', import.meta.url), 'utf8');

describe('Try Yours parsing progress', () => {
  it('uses concise, functional copy for the file workflow', () => {
    expect(source).toContain('Drop in a file.<br />See how it renders.');
    expect(source).toContain('Choose another file');
    expect(source).not.toContain('Inspect freely');
    expect(source).not.toContain('Verify in source');
    expect(source).not.toContain('Replace file ↗');
  });

  it('uses a semantic file-picker button without decorative step numbers', () => {
    expect(source).toContain('<button class="dropzone" id="dropzone" type="button">');
    expect(source).toContain("dz.addEventListener('click', () => input.click())");
    expect(source).not.toContain('dz-index');
    expect(source).not.toContain('02 / Privacy');
  });

  it('shows an accessible progress circle in the preview while renderFile is pending', () => {
    expect(source).toContain('id="stage-progress" role="status" aria-live="polite" hidden');
    expect(source).toContain('class="try-progress-circle" aria-hidden="true"');
    expect(source).toMatch(/\.try-stage-progress\s*\{[\s\S]*?color: var\(--preview-text\);/);
    expect(source).toMatch(/stageProgress\.hidden = false;[\s\S]*await renderFile\(stage, file,/);
  });

  it('hides the progress UI on both the current render success and failure paths', () => {
    expect(source.match(/stageProgress\.hidden = true;/g)).toHaveLength(2);
  });

  it('opens DOCX files progressively and reports when the authoritative page count arrives', () => {
    expect(renderer).toContain("mode: 'worker'");
    expect(renderer).toContain('progressiveLayout: true');
    expect(renderer).toContain('viewer.waitUntilLayoutComplete().then(mountAllPages)');
    expect(renderer).toContain('onLayoutPartial: ({ availableUnits }) =>');
    expect(renderer).toContain('viewerOptions.overscan = availableUnits');
    expect(source).toContain('available · opened in');
    expect(source).toContain('available · loading for');
    expect(source).toContain('res.finalUnits.then');
  });

  it('opens PPTX files progressively in worker mode with a stable final count', () => {
    const branch = renderer.match(/if \(ext === 'pptx'\) \{([\s\S]*?)\n  \}\n\n  const host/)?.[1] ?? '';
    expect(branch).toContain("mode: 'worker'");
    expect(branch).toContain('progressiveLayout: true');
    expect(branch).toContain('viewer.waitUntilLayoutComplete().then(mountAllSlides)');
  });

  it('keeps the preview frame for XLSX as well as DOCX and PPTX', () => {
    expect(source).not.toContain(".try-preview[data-format='xlsx'] { border: 0; }");
  });

  it('prewarms parser engines without downloading fonts for bundled samples', () => {
    const prewarm = renderer.slice(renderer.indexOf('export function prewarmEngines'));
    expect(prewarm).not.toContain('useGoogleFonts: true');
    expect(prewarm.match(/useGoogleFonts:\s*false/g)).toHaveLength(3);
  });
});
