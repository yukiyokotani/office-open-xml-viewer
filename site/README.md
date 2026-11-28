# @silurus/ooxml — showcase site

The public site for [`@silurus/ooxml`](https://www.npmjs.com/package/@silurus/ooxml).
It embeds the real viewers (Rust/WASM + Canvas) rendering live sample
documents. Storybook (`pnpm storybook`) remains the component workbench.

## Develop

```bash
pnpm build:wasm                      # build the parsers first (once)
pnpm --filter @silurus/ooxml-site dev
```

`predev` and `prebuild` first build the `@silurus/ooxml-legacy-office` and
`@silurus/ooxml-opendocument` companions, then copy the redistributable demo
samples from `packages/*/public/demo` into `public/samples/` (office files are
not committed under `site/`).

`/llms.txt`, `/llms-full.txt` and `/api/{docx,xlsx,pptx}.md` give AI agents
Markdown API docs ([llmstxt.org](https://llmstxt.org/)). They are rendered from
`src/lib/api-reference.ts`, the same data as the `/api/*` pages.
`src/lib/api-reference-coverage.test.ts` fails when that data misses or keeps
an option or member of the checked-in public API baselines.

Dev-only `/preview/*` routes render individual sections in isolation; they are
not linked from the site.

## Build

```bash
pnpm --filter @silurus/ooxml-site build      # → site/dist/
SITE_BASE=/office-open-xml-viewer/ pnpm --filter @silurus/ooxml-site build  # project-pages base
```

Deployment (site at `/`, Storybook at `/storybook/`) is
`.github/workflows/deploy-pages.yml`.
