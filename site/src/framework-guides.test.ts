import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const siteRoot = new URL('.', import.meta.url);
const repositoryRoot = new URL('../../', siteRoot);
const exampleFiles = {
  react: { integration: 'react/src/useOfficeViewer.tsx', app: 'react/src/App.tsx' },
  vue: { integration: 'vue/src/useOfficeViewer.ts', app: 'vue/src/App.vue' },
  svelte: { integration: 'svelte/src/createOfficeViewer.ts', app: 'svelte/src/App.svelte' },
  solid: { integration: 'solid/src/createOfficeViewer.ts', app: 'solid/src/App.tsx' },
} as const;
const exampleFrameworks = ['react', 'vue', 'svelte', 'solid'] as const;
const exampleSource = (path: string) => readFileSync(
  new URL(`examples/frameworks/${path}`, repositoryRoot),
  'utf8',
);

describe('framework integration guides', () => {
  it.each(exampleFrameworks)('publishes an independent SEO page for %s', (framework) => {
    const pageUrl = new URL(`pages/frameworks/${framework}.astro`, siteRoot);
    expect(existsSync(pageUrl)).toBe(true);

    const page = readFileSync(pageUrl, 'utf8');
    expect(page).toContain('FrameworkGuide');
    expect(page).toContain(`framework="${framework}"`);
  });

  it('uses a framework chooser as the navigation destination', () => {
    const nav = readFileSync(new URL('components/Nav.astro', siteRoot), 'utf8');
    const index = readFileSync(new URL('pages/frameworks/index.astro', siteRoot), 'utf8');

    expect(nav).toContain('href="/frameworks"');
    expect(nav).not.toContain('href="/frameworks/react"');
    expect(nav.indexOf('href="/try"')).toBeLessThan(nav.indexOf('href="/frameworks"'));
    expect(index).toContain('frameworkGuides.map');
    expect(index).toContain('React, Vue, Svelte, and Solid');
  });

  it('uses the standard site footer on the chooser and framework detail pages', () => {
    const index = readFileSync(new URL('pages/frameworks/index.astro', siteRoot), 'utf8');
    const guide = readFileSync(new URL('components/FrameworkGuide.astro', siteRoot), 'utf8');

    expect(index).toContain('<SiteFooter />');
    expect(guide).toContain('<SiteFooter />');
    expect(guide).not.toContain('guide-footer');
  });

  it.each(['docx', 'xlsx', 'pptx'])('keeps %s framework guidance on the dedicated pages', (format) => {
    const formatPage = readFileSync(new URL(`pages/${format}.astro`, siteRoot), 'utf8');

    expect(formatPage).not.toContain('FrameworkSection');
    expect(formatPage).not.toContain('In your framework');
    expect(existsSync(new URL('components/FrameworkSection.astro', siteRoot))).toBe(false);
  });

  it('keeps Storybook as a development tool rather than a public-site destination', () => {
    const publicNavigation = [
      'components/Nav.astro',
      'layouts/FormatPage.astro',
      'pages/index.astro',
      'pages/try.astro',
    ].map((path) => readFileSync(new URL(path, siteRoot), 'utf8')).join('\n');

    expect(publicNavigation).not.toContain('/storybook/');
    expect(publicNavigation).not.toContain('>Storybook<');
  });

  it('keeps Angular out of the supported framework registry', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    expect(frameworkGuides.map(({ id }) => id)).toEqual(['react', 'vue', 'svelte', 'solid']);
    expect(frameworkGuides.some(({ id }) => (id as string) === 'angular')).toBe(false);
  });

  it('keeps framework dependencies outside the root pnpm workspace', () => {
    const rootWorkspace = readFileSync(new URL('pnpm-workspace.yaml', repositoryRoot), 'utf8');
    const exampleWorkspace = readFileSync(new URL('examples/frameworks/pnpm-workspace.yaml', repositoryRoot), 'utf8');

    expect(rootWorkspace).not.toContain('examples/frameworks');
    expect(exampleWorkspace).toContain('- react');
    expect(exampleWorkspace).toContain('- vue');
    expect(exampleWorkspace).toContain('- svelte');
    expect(exampleWorkspace).toContain('- solid');
  });

  it('uses search-oriented titles without making them identical', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    const titles = frameworkGuides.map(({ title }) => title);

    expect(new Set(titles).size).toBe(4);
    for (const title of titles) {
      expect(title).toMatch(/^How to render Office files in the browser with /);
      expect(title).toContain('DOCX, XLSX, and PPTX');
    }
  });

  it('keeps every integration module portable outside the examples workspace', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    for (const guide of frameworkGuides) {
      const integration = exampleSource(exampleFiles[guide.id].integration);
      expect(integration).not.toContain('@ooxml-framework-examples');
      expect(integration).not.toContain('../shared');
      expect(integration).toContain("from '@silurus/ooxml/docx'");
      expect(integration).toContain("from '@silurus/ooxml/xlsx'");
      expect(integration).toContain("from '@silurus/ooxml/pptx'");
      expect(integration).not.toContain("await import('@silurus/ooxml");
      expect(integration).toContain('destroy');
      expect(integration).not.toMatch(/\blet\s+/);
      expect(integration).not.toMatch(/\bvoid\s+[A-Za-z_(]/);
    }
  });

  it('implements the React integration as a render hook with an internal ref and cleanup', () => {
    const hook = readFileSync(new URL('examples/frameworks/react/src/useOfficeViewer.tsx', repositoryRoot), 'utf8');
    expect(hook).toContain('const mountRef = useRef<HTMLDivElement>(null);');
    expect(hook).toContain('const renderOfficeViewer = useCallback(');
    expect(hook).toContain('<div {...props} ref={mountRef} />');
    expect(hook).toContain('useEffect(() =>');
    expect(hook).toContain('const controller = new AbortController();');
    expect(hook).toContain('viewer.destroy();');
    expect(hook).not.toMatch(/\blet\s+/);
    expect(hook).not.toMatch(/\bvoid\s+mountOfficeViewer/);
    expect(hook).not.toContain('targetRef:');
  });

  it('keeps viewer construction out of components and accepts a replaceable local file source', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    for (const guide of frameworkGuides) {
      const app = exampleSource(exampleFiles[guide.id].app);
      expect(app).not.toContain("import('@silurus/ooxml");
      expect(app).not.toContain("from '@silurus/ooxml");
      expect(app).not.toContain('DocxScrollViewer');
      expect(app).not.toContain('PptxScrollViewer');
      expect(app).not.toContain('XlsxViewer');
      expect(app).toContain('file.arrayBuffer()');
      expect(app).toContain('.docx,.xlsx,.pptx');
      expect(app).toContain("'Choose an Office file'");
      expect(app).not.toContain('raw.githubusercontent.com');
      expect(app).not.toContain('<canvas');
    }
  });

  it('embeds each runnable project with StackBlitz', async () => {
    const guide = readFileSync(new URL('components/FrameworkGuide.astro', siteRoot), 'utf8');
    const stackBlitz = readFileSync(new URL('components/FrameworkStackBlitz.astro', siteRoot), 'utf8');
    const { frameworkGuides } = await import('./lib/framework-guides');

    expect(guide).toContain('FrameworkStackBlitz');
    expect(guide).not.toContain('LiveShowcase');
    expect(guide).not.toContain('CodeTabs');
    expect(guide).not.toContain('Integration module');
    expect(guide).not.toContain('<h2>Component</h2>');
    expect(guide).not.toContain('Step 1');
    expect(guide).not.toContain('Step 2');
    expect(guide).not.toContain('Open live');
    expect(stackBlitz).toContain('<iframe');
    for (const framework of frameworkGuides) {
      expect(framework.stackBlitzEmbedUrl).toContain(
        `stackblitz.com/github/yukiyokotani/office-open-xml-viewer/tree/main/examples/frameworks/${framework.id}`,
      );
      expect(framework.stackBlitzEmbedUrl).toContain('embed=1');
      expect(framework.stackBlitzEmbedUrl).toContain('view=editor');
      expect(framework.stackBlitzEmbedUrl).toContain('showSidebar=1');
      expect(framework.stackBlitzEmbedUrl).not.toContain('startScript');
      expect(framework.stackBlitzUrl).toContain('startScript=dev');
      expect(framework.stackBlitzUrl).not.toContain('startScript=dev%3A');
    }
  });

  it.each(exampleFrameworks)('keeps the %s StackBlitz project self-contained', (framework) => {
    const mainExtension = framework === 'react' || framework === 'solid' ? 'tsx' : 'ts';
    const main = exampleSource(`${framework}/src/main.${mainExtension}`);
    const packageJson = JSON.parse(exampleSource(`${framework}/package.json`)) as {
      stackblitz?: { startCommand?: boolean };
    };

    expect(main).toContain("import './example.css';");
    expect(existsSync(new URL(`examples/frameworks/${framework}/src/example.css`, repositoryRoot))).toBe(true);
    expect(packageJson.stackblitz?.startCommand).toBe(false);
  });

  it.each(exampleFrameworks)('uses the current library minor line in the %s example', (framework) => {
    const rootPackage = JSON.parse(readFileSync(new URL('package.json', repositoryRoot), 'utf8')) as {
      version: string;
    };
    const packageJson = JSON.parse(exampleSource(`${framework}/package.json`)) as {
      dependencies?: Record<string, string>;
    };

    const [major, minor] = rootPackage.version.split('.');
    expect(packageJson.dependencies?.['@silurus/ooxml']).toBe(`^${major}.${minor}.0`);
  });

  it.each(exampleFrameworks)(
    'keeps the %s viewer surface transparent and owns the desk background in CSS',
    (framework) => {
      const integration = exampleSource(exampleFiles[framework].integration);
      const css = exampleSource(`${framework}/src/example.css`);

      expect(integration).not.toMatch(/\bbackground\s*:/);
      expect(css).toMatch(/\.stage\s*\{[\s\S]*?background:\s*#53606d;/);
    },
  );

});
