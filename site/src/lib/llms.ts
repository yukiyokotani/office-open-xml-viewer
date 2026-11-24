// llms.txt (https://llmstxt.org/) output for AI agents. The site is not a
// Markdown/content-collection site, so HTML-scraping integrations would only
// recover flattened tables. Instead the Markdown is rendered from the same
// structured data the API pages use (api-reference.ts), so the agent-facing
// text cannot drift from the human-facing reference.
import {
  apiReference,
  formatRenderModeGuidance,
  optionalRenderers,
  type ApiClass,
} from './api-reference';
import { docxSnippets, pptxSnippets, xlsxSheetSnippet, xlsxSheetWindowsSnippet } from './demo-snippets';
import { canonicalPageUrl, siteMetadata } from './seo';

export const llmsFormats = ['docx', 'xlsx', 'pptx'] as const;
export type LlmsFormat = (typeof llmsFormats)[number];

const formatName: Record<LlmsFormat, string> = { docx: 'DOCX', xlsx: 'XLSX', pptx: 'PPTX' };

export const apiMarkdownPath = (format: LlmsFormat): string => `/api/${format}.md`;

/** Guides linked from llms.txt; titles and summaries mirror each page's own metadata. */
export const llmsGuides: readonly { path: string; title: string; summary: string }[] = [
  { path: '/production/', title: 'Production decisions', summary: 'Rendering modes (main / worker), model ownership and optional renderers.' },
  { path: '/errors/', title: 'Error handling', summary: 'Typed errors such as OoxmlError, OoxmlResourceLimitError and their fields.' },
  { path: '/fonts/', title: 'Font sources', summary: 'The fonts option: Google Fonts, mirrors, hosted font services, own files or a resolver.' },
  { path: '/frameworks/', title: 'Framework guides', summary: 'Using the viewers from React, Vue, Svelte and Solid.' },
  { path: '/selection-context/', title: 'Selection context for AI and MCP', summary: 'Bounded, read-only selection context for AI assistants and external integrations.' },
  { path: '/review-ui/', title: 'Comments', summary: 'Built-in comment display, or an application-owned review UI from comment data.' },
  { path: '/legacy-office/', title: 'Legacy Office files', summary: 'Experimental DOC, XLS and PPT reading through @silurus/ooxml-legacy-office.' },
  { path: '/opendocument/', title: 'OpenDocument files', summary: 'Experimental ODT, ODS and ODP reading through @silurus/ooxml-opendocument.' },
  { path: '/bundle-size/', title: 'Bundle size', summary: 'Current production bundle-size measurements.' },
];

// Site-relative links may carry a fragment (`/errors#delivery`); canonicalize
// only the path so the trailing slash lands before the `#`.
const absolute = (href: string): string => {
  if (!href.startsWith('/')) return href;
  const hash = href.indexOf('#');
  return hash < 0 ? canonicalPageUrl(href) : canonicalPageUrl(href.slice(0, hash)) + href.slice(hash);
};

// The same snippets the per-format pages show beside their live demos.
const examples: Record<LlmsFormat, readonly { title: string; code: string }[]> = {
  docx: [
    { title: 'Single-canvas viewer', code: docxSnippets.demo },
    { title: 'Scrolling viewer', code: docxSnippets.scroll },
  ],
  pptx: [
    { title: 'Single-canvas viewer', code: pptxSnippets.demo },
    { title: 'Scrolling viewer', code: pptxSnippets.scroll },
  ],
  xlsx: [
    { title: 'Workbook viewer', code: xlsxSheetSnippet },
    { title: 'Sheet viewers over one workbook', code: xlsxSheetWindowsSnippet },
  ],
};

// Table cells cannot hold raw pipes or newlines.
const cell = (text: string): string => text.replace(/\|/g, '\\|').replace(/\s*\n\s*/g, ' ');
const codeCell = (text: string): string => `\`${cell(text).replace(/`/g, "'")}\``;

function classMarkdown(apiClass: ApiClass): string {
  const lines = [`### ${apiClass.name}`, ''];
  if (apiClass.note) lines.push(apiClass.note, '');
  lines.push('```ts', apiClass.ctor, '```', '');
  if (apiClass.options?.length) {
    lines.push('#### Options', '', '| Name | Type | Default | Description |', '| --- | --- | --- | --- |');
    for (const option of apiClass.options) {
      const details = option.detailsHref
        ? ` See [${option.detailsLabel ?? 'details'}](${absolute(option.detailsHref)}).`
        : '';
      lines.push(`| ${codeCell(option.name)} | ${codeCell(option.type)} | ${option.def ? codeCell(option.def) : '—'} | ${cell(option.desc)}${details} |`);
    }
    lines.push('');
  }
  if (apiClass.methods.length) {
    lines.push('#### Methods', '');
    for (const method of apiClass.methods) lines.push(`- \`${method.sig}\`: ${method.desc}`);
    lines.push('');
  }
  return lines.join('\n');
}

/** Markdown version of /api/<format>/ (the llms.txt convention for page URLs + `.md`). */
export function apiMarkdown(format: LlmsFormat): string {
  const name = formatName[format];
  return [
    `# ${name} API reference — ${siteMetadata.name}`,
    '',
    `> Public options and methods of the ${siteMetadata.name} ${name} viewers and document model. Import from \`@silurus/ooxml/${format}\`. HTML version: ${canonicalPageUrl(`/api/${format}/`)}`,
    '',
    `Rendering modes: ${formatRenderModeGuidance[format]}`,
    '',
    '## Examples',
    '',
    ...examples[format].flatMap(({ title, code }) => [`### ${title}`, '', '```ts', code.trim(), '```', '']),
    '## Classes',
    '',
    ...apiReference[format].map(classMarkdown),
  ].join('\n');
}

function optionalRenderersMarkdown(): string {
  return [
    '## Optional renderers',
    '',
    'Each is a separate entry point, imported once and passed to a viewer or document option of the same name.',
    '',
    ...optionalRenderers.map((r) => `- **${r.name}**: \`import { ${r.exportName} } from '${r.entry}'\` (${r.contract}). ${r.desc}`),
    '',
  ].join('\n');
}

export function llmsTxt(): string {
  return [
    `# ${siteMetadata.name}`,
    '',
    `> ${siteMetadata.description}`,
    '',
    'Install `@silurus/ooxml` and import a format entry: `@silurus/ooxml/docx`, `@silurus/ooxml/xlsx` or `@silurus/ooxml/pptx`. Each entry exports Canvas viewer classes and a headless document, workbook or presentation model (listed below). Equations, ChartEx, 3-D charts, Region Maps and TIFF are opt-in renderers from their own entries. To read a file\'s text without rendering, load the headless model and call `toMarkdown()`.',
    '',
    '## API reference',
    '',
    ...llmsFormats.map((format) => `- [${formatName[format]} API reference](${canonicalPageUrl(apiMarkdownPath(format))}): ${apiReference[format].map((c) => c.name).join(', ')} options and methods`),
    '',
    '## Guides',
    '',
    ...llmsGuides.map((g) => `- [${g.title}](${canonicalPageUrl(g.path)}): ${g.summary}`),
    '',
    '## Optional',
    '',
    `- [Full API reference](${canonicalPageUrl('/llms-full.txt')}): every format's API reference and the optional renderers in one file`,
    '- [Source repository](https://github.com/yukiyokotani/office-open-xml-viewer): README, changelog and migration guides',
    '',
  ].join('\n');
}

/** Nest a page under llms-full.txt's H1, leaving fenced code untouched. */
function demoteHeadings(markdown: string): string {
  let fenced = false;
  return markdown.split('\n').map((line) => {
    if (line.startsWith('```')) fenced = !fenced;
    return !fenced && line.startsWith('#') ? `#${line}` : line;
  }).join('\n');
}

export function llmsFullTxt(): string {
  return [
    `# ${siteMetadata.name}`,
    '',
    `> ${siteMetadata.description}`,
    '',
    optionalRenderersMarkdown(),
    ...llmsFormats.map((format) => demoteHeadings(apiMarkdown(format))),
  ].join('\n');
}

export const markdownResponse = (body: string, type = 'text/markdown'): Response =>
  new Response(body, { headers: { 'Content-Type': `${type}; charset=utf-8` } });
