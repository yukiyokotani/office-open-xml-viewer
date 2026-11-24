import { describe, expect, it } from 'vitest';
import { apiReference } from './lib/api-reference';
import { apiMarkdown, llmsFormats, llmsFullTxt, llmsTxt } from './lib/llms';
import { GET as getLlms } from './pages/llms.txt';
import { GET as getLlmsFull } from './pages/llms-full.txt';
import { GET as getApiMarkdown } from './pages/api/[format].md';

describe('llms.txt', () => {
  it('follows the llmstxt.org shape: H1, summary blockquote, link sections', () => {
    const body = llmsTxt();
    expect(body.startsWith('# @silurus/ooxml\n\n> ')).toBe(true);
    expect(body).toContain('## API reference');
    expect(body).toContain('## Optional');
    for (const format of llmsFormats) {
      expect(body).toContain(`(https://ooxml.silurus.dev/api/${format}.md)`);
    }
    expect(body).toContain('(https://ooxml.silurus.dev/llms-full.txt)');
  });

  it('renders every class, option and method from the API reference data', () => {
    const full = llmsFullTxt();
    for (const format of llmsFormats) {
      const md = apiMarkdown(format);
      for (const apiClass of apiReference[format]) {
        expect(md).toContain(`### ${apiClass.name}`);
        expect(full).toContain(`#### ${apiClass.name}`);
        for (const option of apiClass.options ?? []) expect(md).toContain(`| \`${option.name}\` |`);
        for (const method of apiClass.methods) expect(md).toContain(method.sig);
      }
    }
  });

  it('keeps option tables one row per option', () => {
    for (const format of llmsFormats) {
      const rows = apiMarkdown(format).split('\n').filter((line) => line.startsWith('| `'));
      const options = apiReference[format].flatMap((c) => c.options ?? []);
      expect(rows).toHaveLength(options.length);
    }
  });

  it('serves the endpoints as UTF-8 text', async () => {
    expect(getLlms().headers.get('Content-Type')).toBe('text/plain; charset=utf-8');
    expect(await getLlmsFull().text()).toBe(llmsFullTxt());
    const docx = getApiMarkdown({ params: { format: 'docx' } });
    expect(docx.headers.get('Content-Type')).toBe('text/markdown; charset=utf-8');
    expect(await docx.text()).toBe(apiMarkdown('docx'));
  });
});
