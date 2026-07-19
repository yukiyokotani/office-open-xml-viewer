# OOXML edge-case corpus

This directory contains redistributable, synthetic test data for the DOCX,
PPTX, and XLSX renderers. It is intentionally separate from
`packages/*/public/demo`: demo files showcase the product, while this corpus
exists to expose rendering and parsing boundaries.

The three primary files form one restrained visual series, **OOXML Edge Case
Field Guide**:

- `docx/edge-cases.docx`
- `pptx/edge-cases.pptx`
- `xlsx/edge-cases.xlsx`

Each case has a stable identifier in `manifest.json`. Case identifiers, rather
than page, slide, or sheet order, are the durable API for tests and bug reports.

## Design rules

- White or very light backgrounds, navy and teal accents, and system fonts.
- Decoration is deliberately minimal and must not introduce unrelated rendering
  dependencies.
- A feature is styled only when the styling itself is under test.
- All text, data, and imagery are synthetic and safe to redistribute.
- Corrupt packages, encryption, Strict namespace variants, relationship faults,
  and ZIP limits belong in small package fixtures rather than these visual files.

## Reference PDFs

Word and PowerPoint references are exported by the maintainer from Microsoft
Office and committed next to the source file:

- `docx/reference/edge-cases.pdf`
- `pptx/reference/edge-cases.pdf`

The files are intentionally absent until an Office export is available. Follow
the instructions in each `reference/README.md`; do not create or replace an
Office reference from viewer output.

XLSX uses structural assertions and renderer snapshots because a PDF depends on
Excel print-area and pagination settings rather than the workbook viewport.

## Maintenance

The committed Office files are the inputs consumed by tests. Authoring scripts
and inspection output are local-only and ignored; they are not part of this
corpus. When a binary changes, review `manifest.json` and render every page,
slide, and worksheet before committing it.

Private documents remain local-only under the existing
`packages/*/public/private/` paths. They may help diagnose a bug, but a permanent
fix should be represented by a synthetic case here whenever possible.
