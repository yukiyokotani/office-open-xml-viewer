# Review UI extension design

## Status

This is the boundary for the read-only comment UI. Editing, replying, resolving,
and application-specific review workflows remain application responsibilities.

## Two integration levels

### Built-in UI

`DocxScrollViewer` and `PptxScrollViewer` can show a page-side comment margin.
`XlsxViewer` and `XlsxSheetViewer` can show cell-anchored comment popups. These
Viewers own placement, zoom updates, page or sheet lifecycle, and cleanup.

The built-in structure is intentionally fixed. Stable classes are the primary
presentation contract for changes that do not replace the structure:

- `.ooxml-comment-card`;
- `.ooxml-comment-card__author`;
- `.ooxml-comment-card__date`;
- `.ooxml-comment-card__body`;
- `.ooxml-comment-card__reply`;
- `.ooxml-comment-marker`.

CSS custom properties remain a smaller convenience layer for common theme
tokens such as colors, card borders, and corner radius. They do not mirror every
CSS property and do not replace the class contract.

Cards expose `data-active` and `data-focused` styling states. Defaults use
low-specificity `:where(...)` selectors. Inline styles are reserved for dynamic
geometry, Viewer zoom, and the computed author accent. Other internal
`data-ooxml-comment-*` attributes support Viewer behavior and tests and are not
a public styling or DOM-structure contract. There is no component-mount callback
or framework-specific adapter.

CSS cascade changes apply to already-mounted UI; applications can switch a
theme class, `data-theme`, or custom properties without recreating a Viewer.
Cards are observed for geometry changes, so font, line-height, family, padding,
and other size-affecting overrides trigger a new non-overlapping layout.

This keeps the common path small and makes it usable from plain TypeScript,
React, Vue, and other frameworks without giving the Viewer ownership of an
application component tree.

### Application-owned UI

Applications that need different structure or behavior build their UI from the
format APIs:

- comment records and replies;
- logical anchors;
- rendered text-run geometry for DOCX;
- slide coordinates for PPTX;
- cell references and `getCellViewportRect()` for XLSX.

The application then owns its DOM or framework components, interaction model,
virtualization, and cleanup. The Viewer does not expose an intermediate card
renderer abstraction: such an abstraction would still constrain component
ownership and would duplicate framework lifecycle rules.

## Shared policy and format boundaries

All formats use the single `comments` feature option. `comments: true` enables
the format default; `comments: { includeResolved: … }` also controls thread
visibility. DOCX and PPTX hide resolved or closed threads by default. XLSX
preserves its historical behavior and includes resolved threads by default.

The formats deliberately do not pretend their geometry is identical:

- DOCX comments attach to logical text ranges that are resolved against one
  rendered page's text runs;
- PPTX comments attach to slide coordinates;
- XLSX comments attach to cells and use a pointer or keyboard popup. Focusing
  the viewport establishes a cell selection, Arrow keys move it, Enter opens
  the selected cell's comment, and a polite live status announces its content.

Core owns the small built-in card style vocabulary. Each format owns projection
from its OOXML model into its UI geometry.

## Progressive loading and virtualization

The built-in UI follows the Viewer's mounted page, slide, or sheet lifecycle.
When progressive DOCX layout publishes or revises visible pages, comment
projection must be refreshed through the same relayout path; it must not infer
pagination independently.

Page and slide virtualization remain the first bounded layer. A future need for
card-level virtualization should be implemented inside the built-in margin,
without changing the public `comments` option. Application-owned UIs choose
their own list virtualization strategy from the primitive data APIs.

No public identity or geometry type is introduced solely to predict a future
virtualization implementation.

## Change-history boundary and future API symmetry

Comments and change history are related review features, but they are not one
OOXML model. The public API must not combine them into a shared `ReviewItem`
union merely to make the three formats look alike.

The formats currently have different source models:

| Format | Persisted source | Current library support |
| --- | --- | --- |
| DOCX | WordprocessingML revision containers such as `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo` | Detached body-story records and logical revision ranges |
| XLSX | SpreadsheetML revision headers and revision logs, including cell, row/column, move, formatting, sheet, name, and comment changes | Not parsed or exposed |
| PPTX | PresentationML has no general revision-log model equivalent to WordprocessingML tracked changes. PowerPoint comparison results and cloud collaboration indicators have different lifecycles and are not assumed to be self-contained revision records in every `.pptx` package. | Not parsed or exposed |

Future change-history work follows the same three-layer architecture as comments
without forcing the record shapes to match:

1. The owned engine exposes immutable, format-specific source records at their
   natural scope: a DOCX document, an XLSX workbook with sheet identity on its
   records, or a PPTX presentation/slide only when a persisted PowerPoint source
   model has been identified.
2. A separate format-specific projection resolves a source record to the
   rendered surface: page occurrences for DOCX, sheet references or ranges for
   XLSX, and slide coordinates or element identity for PPTX when available.
3. A composite Viewer may present those projections, while primitive engine and
   focused-Viewer APIs remain sufficient for an application-owned UI.

The query pattern should remain recognizable across engines (`getRevisions` or
an equally explicit final name for source records, plus a separate
surface-occurrence query), but record and geometry types remain format-specific.
The exact public names are fixed only when the first XLSX implementation and a
real PPTX persisted model are backed by specification references and
Office-produced boundary samples. This avoids freezing a DOCX-shaped abstraction
before the other domains are understood.

Unsupported formats expose no placeholder method. Returning `[]` would conflate
"the package contains no changes" with "this source model is not implemented".
Likewise, comparing two presentation files is a separate comparison API; it must
not masquerade as revision records read from one presentation.

## Non-goals

- comment or revision editing;
- Word, Excel, or PowerPoint chrome reproduction;
- a general application panel framework;
- React, Vue, or other framework adapters;
- arbitrary DOM replacement inside the built-in UI;
- combining comments and tracked changes into one artificial OOXML model.

## Acceptance checks

- the default UI works without callbacks;
- CSS custom properties cover simple theme tokens and stable classes allow
  presentation overrides without replacing the built-in structure;
- DOCX/PPTX cards and geometry follow zoom and virtualized surface lifecycle;
- XLSX popup data and visible markers use the same resolved-thread policy;
- outside interaction clears the active DOCX/PPTX card;
- primitive comment and anchor APIs remain sufficient for a completely
  application-owned UI;
- public API declarations contain no framework mount lifecycle.
