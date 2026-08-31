import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const bootstrap = readFileSync(new URL('./webview/bootstrap.ts', import.meta.url), 'utf8');

function initializer(name: 'Docx' | 'Pptx'): string {
  const source = bootstrap.match(
    new RegExp(`async function init${name}\\([\\s\\S]*?(?=\\n// ──|$)`),
  )?.[0];
  expect(source).toBeDefined();
  return source as string;
}

describe('VS Code progressive continuous-scroll previews', () => {
  it.each(['Docx', 'Pptx'] as const)('enables progressive layout for %s', (format) => {
    expect(initializer(format)).toContain('progressiveLayout: true');
  });
});
