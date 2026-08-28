import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('DOCX progressive-layout guide', () => {
  const guide = source('./components/DocxProgressiveLayoutGuide.astro');

  it('uses the shared code theme contract', () => {
    expect(guide).toContain('themes={codeThemes} defaultColor={false}');
    expect(guide).toContain('background: var(--code-bg) !important');
  });

  it('keeps the section title above the copy and code content row', () => {
    expect(guide).toContain('class="progressive-head"');
    expect(guide).toContain('class="progressive-body"');
    expect(guide.indexOf('class="progressive-head"')).toBeLessThan(
      guide.indexOf('class="progressive-body"'),
    );
  });
});
