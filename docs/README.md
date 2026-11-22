# Documentation

The root [`README.md`](../README.md), the package READMEs and the generated
public API surfaces are the source of truth for supported features. This
directory holds the longer guides and the architecture contracts behind them.

Current documents are maintained in English. Parallel translations are not
kept because they drift from the implementation contract.

## Guides for library users

- [`selection-context.md`](selection-context.md) — the selection-context
  contract shared by the viewers.
- [`docx-footnote-continuations.md`](docx-footnote-continuations.md) — how
  page-bottom footnotes continue across pages, and how to opt out.
- [`legacy-office-conversion.md`](legacy-office-conversion.md) — the
  experimental `@silurus/ooxml-legacy-office` companion (DOC, XLS, PPT).
- [`opendocument.md`](opendocument.md) — the experimental
  `@silurus/ooxml-opendocument` companion.
- [`chart-compatibility-evidence.md`](chart-compatibility-evidence.md) — chart
  compatibility evidence and limits.
- [`chart-support-matrix.md`](chart-support-matrix.md) — parser, model and
  renderer support for each chart type.
- [`pptx-follow-path-compatibility.md`](pptx-follow-path-compatibility.md) —
  multiline Follow Path WordArt placement.

## Migration guides

Each guide covers one consumer-visible migration. They stay at these paths
because released packages and changelog entries link to them.

- [`migration-0.90.md`](migration-0.90.md) — font sources (`fonts`) replace
  `useGoogleFonts`.
- [`migration-selection-context-0.77.md`](migration-selection-context-0.77.md)
- [`migration-xlsx-selection-0.77.md`](migration-xlsx-selection-0.77.md)
- [`migration-viewer-errors-0.77.md`](migration-viewer-errors-0.77.md)
- [`migration-mcp-tools-0.77.md`](migration-mcp-tools-0.77.md)
- [`migration-0.76.md`](migration-0.76.md) — 0.75 to 0.76.

## Architecture and contracts

- [`api-architecture-0.76.md`](api-architecture-0.76.md) — public API,
  ownership and responsibility boundaries since 0.76.
- [`review-ui-extension-design.md`](review-ui-extension-design.md) — the
  read-only comment UI boundary.
- [`ooxml-resource-governance-design.md`](ooxml-resource-governance-design.md)
  and [`ooxml-resource-default-calibration.md`](ooxml-resource-default-calibration.md)
  — resource budgets and how their defaults were calibrated.
- [`docx-layout-engine-redesign.md`](docx-layout-engine-redesign.md) — DOCX
  layout architecture and invariants.
- [`docx-model-projection-layers.md`](docx-model-projection-layers.md) — the
  DOCX model-to-layout boundary and its costs.
- [`docx-page-layer-order.md`](docx-page-layer-order.md) — DOCX page layer
  ordering.

## Contributor workflow

- [`docx-architecture-audit.md`](docx-architecture-audit.md) — the review
  procedure for DOCX layout changes.
- [`../scripts/README.md`](../scripts/README.md) — what each repository script
  does and how to rerun the generators.

## Archive

[`archive/`](archive/README.md) keeps completed plans and audits for their
design rationale. They are not a backlog; use open GitHub issues for current
work.

`images/` holds the screenshots and icon used by the READMEs and the site.
