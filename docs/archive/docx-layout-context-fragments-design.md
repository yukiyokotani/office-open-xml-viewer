# DOCX Layout Context and Measured Fragments Design

## Status

Approved architecture direction. This document defines the target design and the
staged delivery boundaries. It does not authorize sample-specific heuristics or
a one-shot renderer rewrite.

## Problem

The DOCX renderer already parses and renders many WordprocessingML layout
features, but layout policy and measurement are distributed across several
paths:

- document settings and section properties are copied into mutable render state;
- body paragraphs, table-cell paragraphs, and text boxes resolve line metrics
  through different call paths;
- pagination estimates content, then paint can measure it again;
- table pagination represents continuation by cloning model objects and adding
  runtime-only fields such as line and row slices;
- floating drawing registration, line wrapping, and page placement share mutable
  state without an explicit layout contract.

These paths can disagree even when each path is locally reasonable. A private,
local-only fixture exposed one such disagreement, but the defect class is not
fixture-specific: the renderer lacks one normalized context and one measured
geometry model for body and table flow.

## Normative Basis

The design follows these ECMA-376 Part 1 rules:

- Section 17.6.5, `docGrid`: a section grid defines line pitch and character
  pitch behavior.
- Section 17.18.14, `ST_DocGrid`: `lines` enables the line grid;
  `linesAndChars` enables line and character grids; `snapToChars` also applies
  both line and character pitch; `default` disables the document grid.
- Section 17.3.1.32, paragraph `snapToGrid`: a paragraph uses the section line
  grid by default and can opt out.
- Section 17.3.2.34, run `snapToGrid`: a run uses the section character grid by
  default and can opt out. This property does not control line pitch.
- Section 17.3.1.33, paragraph `spacing`: paragraph spacing, line spacing, and
  `auto`, `exact`, and `atLeast` line rules have distinct units and precedence.
- Section 17.6.5 states that `exact` line spacing overrides the section line
  pitch.
- Section 17.15.3.1, `adjustLineHeightInTable`: section line pitch is applied
  inside table cells only when this compatibility setting is enabled.
- Section 20.4.2.3, `anchor`: inline drawings participate in line layout;
  floating drawings are positioned independently from inline flow.
- Section 20.4.2.15, `wrapNone`: a floating drawing with no wrapping must not
  constrain text flow.

Microsoft implementation notes and controlled Office output comparisons may be
used when the standard leaves behavior unspecified. Any such compatibility
behavior must be isolated, documented, and tested separately from normative
rules.

## Goals

1. Normalize document, section, story, container, paragraph, and run layout
   policy into immutable TypeScript contexts.
2. Preserve Rust parser ownership of style-cascade resolution and TypeScript
   ownership of runtime layout meaning.
3. Make paragraph measurement placement-aware, deterministic, and reusable by
   body flow, table-cell flow, pagination, and paint.
4. Replace model-object layout stamps with explicit measured fragments owned by
   a `DocumentLayout` result.
5. Migrate body paragraphs and table pagination incrementally, with each stage
   independently testable and reviewable.
6. Keep existing anchored-drawing and text-box rendering operational while
   exposing contracts that later migrations can consume.

## Non-Goals

- Rewriting the complete DOCX renderer in one change.
- Migrating text-box content to the full WordprocessingML body model in this
  delivery series.
- Replacing anchored drawing registration, z-order, or positioning in this
  delivery series.
- Adding empirical constants or branches for a private fixture.
- Reinterpreting unrelated Office compatibility behavior while extracting the
  new layout kernel.

## Architectural Decisions

### Story and container are orthogonal

A table cell is not a document story. A story describes the WordprocessingML
content source. A container stack describes nested layout environments within
that story.

```ts
export type StoryKind =
  | 'body'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'textbox';

export type ContainerFrame =
  | { readonly kind: 'tableCell' };

export interface StoryContext {
  readonly story: StoryKind;
  readonly containers: readonly ContainerFrame[];
  readonly lineNumberingEligible: boolean;
}
```

The stack, rather than a separate depth integer, represents nested tables and
future combinations such as a table inside a text box. Container rules are
limited to behavior backed by the specification:

- table-cell line-grid compatibility gating;
- table-cell isolation from page-level floating wrap constraints;
- main-story eligibility for section line numbering.

Headers, footers, and notes do not receive hard-coded document-grid exceptions.
Their paragraph styles, including inherited `snapToGrid`, remain authoritative.

### Parser and renderer responsibilities remain separate

Rust resolves information that depends on the WordprocessingML style cascade:

- document defaults, named styles, table styles, numbering, character styles,
  and direct formatting;
- whether line spacing was explicit at a style/direct level;
- resolved paragraph and run properties.

Rust must additionally retain, without applying layout formulas:

- `w:compat/w:adjustLineHeightInTable` as an optional boolean;
- run-level `w:snapToGrid` as an optional boolean.

TypeScript resolves behavior that depends on runtime containment or placement:

- section grid activation and units;
- table-cell compatibility gating;
- paragraph line-grid and run character-grid participation;
- physical bidi indents;
- story/container-specific flow policy;
- line measurement, placement, fragmentation, and painting.

New parser fields use optional serialization so documents that do not contain
the properties retain their previous wire representation.

## Layout Contexts

### Document settings

```ts
export interface DocumentLayoutSettings {
  readonly kinsoku: KinsokuRules;
  readonly defaultTabPt: number;
  readonly characterSpacingControl?: string;
  readonly mathDefJc?: string;
  readonly documentHasEastAsianText: boolean;
  readonly compat: {
    readonly adjustLineHeightInTable: boolean;
    readonly useFeLayout: boolean;
    readonly balanceSingleByteDoubleByteWidth: boolean;
  };
}

export function resolveDocumentLayoutSettings(
  document: DocxDocumentModel,
): DocumentLayoutSettings;
```

Absent compatibility flags resolve to `false`. Existing specification defaults,
including the default tab interval and kinsoku behavior, resolve once here.

### Section context

Line-grid and character-grid policy are separate. This prevents paragraph
`snapToGrid` from accidentally disabling a run-level character grid, and
prevents the table-cell line-pitch exception from disabling character pitch.

```ts
export interface SectionGridContext {
  readonly kind: 'none' | 'lines' | 'linesAndChars' | 'snapToChars';
  readonly linePitchPt: number | null;
  readonly charSpacePt: number | null;
}

export interface SectionLayoutContext {
  readonly geometry: SectionGeom;
  readonly columns: readonly ColumnGeom[];
  readonly grid: SectionGridContext;
  readonly textDirection: string;
  readonly verticalAlignment: string;
  readonly lineNumbering?: LineNumbering;
}

export function resolveSectionLayoutContext(
  settings: DocumentLayoutSettings,
  section: SectionProps,
): SectionLayoutContext;
```

Vertical-page coordinate swapping stays before context resolution. The section
resolver consumes the already-normalized logical section geometry.

### Paragraph and run contexts

```ts
export interface LineGridPolicy {
  readonly active: boolean;
  readonly pitchPt: number | null;
}

export interface CharacterGridPolicy {
  readonly active: boolean;
  readonly deltaPt: number;
}

export interface ParagraphLayoutContext {
  readonly lineGrid: LineGridPolicy;
  readonly characterGrid: CharacterGridPolicy;
  readonly physicalIndentLeftPt: number;
  readonly physicalIndentRightPt: number;
  readonly firstIndentPt: number;
  readonly lineSpacing: LineSpacing | null;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly baseRtl: boolean;
  readonly tabStops: readonly TabStop[];
  readonly hasRuby: boolean;
  readonly hasEastAsianText: boolean;
  readonly kinsoku: KinsokuRules;
  readonly defaultTabPt: number;
}

export interface RunLayoutContext {
  readonly characterGrid: CharacterGridPolicy;
}

export function resolveParagraphLayoutContext(
  settings: DocumentLayoutSettings,
  section: SectionLayoutContext,
  story: StoryContext,
  paragraph: DocParagraph,
): ParagraphLayoutContext;

export function resolveRunLayoutContext(
  paragraph: ParagraphLayoutContext,
  run: DocxTextRun,
): RunLayoutContext;
```

Grid resolution follows this matrix:

| Input | Line grid | Character grid |
| --- | --- | --- |
| `docGrid` absent or `type=default` | off | off |
| `type=lines` | on | off |
| `type=linesAndChars` | on | on when `charSpace` is present |
| `type=snapToChars` | on | on when `charSpace` is present |
| paragraph `snapToGrid=false` | off | unchanged |
| paragraph `lineRule=exact` | off | unchanged |
| table cell without `adjustLineHeightInTable` | off | unchanged |
| run `snapToGrid=false` | unchanged | off for that run |

`atLeast` does not disable the line grid. Its requested minimum and the resolved
grid minimum both participate in line-box calculation. This behavior change is
delivered and verified separately from the context extraction.

Existing Office-compatible East Asian cell rounding, ruby line unification, and
font metric correction move behind this context without semantic changes during
the extraction stage. Any later change to those behaviors requires its own
evidence and test boundary.

Paragraph adjacency behavior does not belong in `ParagraphLayoutContext`.
Spacing collapse, contextual spacing, section-break spacer handling,
`keepNext`, `keepLines`, and widow/orphan decisions remain flow-placement rules
because they depend on neighboring content or available page space.

## Placement-Aware Paragraph Measurement

Paragraph measurement is not intrinsic to a paragraph and width. Floating wrap
geometry can change line width and vertical origin, so the placement is an
explicit part of the contract.

```ts
export interface LineLayoutEnvironment {
  readonly pageIndex: number;
  readonly totalPages: number;
  readonly displayPageNumber?: number;
  readonly pageNumberFormat?: NumberFormat;
  readonly currentDateMs?: number;
  readonly noteNumbers?: ReadonlyMap<string, number>;
  readonly currentNoteNumber?: number;
  readonly verticalCJK?: boolean;
}

export interface ParagraphMeasurementEnvironment extends LineLayoutEnvironment {
  readonly documentHasEastAsianText: boolean;
}

export interface WrapOracle {
  lineWindow(input: {
    readonly topYPt: number;
    readonly minimumStartWidthPt: number;
    readonly probeHeightPt: number;
    readonly paragraphXPt: number;
    readonly maximumWidthPt: number;
  }): {
    readonly topYPt: number;
    readonly xOffsetPt: number;
    readonly maximumWidthPt: number;
  };

  skipTopAndBottomBands(yPt: number): number;
}

export interface ParagraphPlacement {
  readonly startYPt: number;
  readonly paragraphXPt: number;
  readonly availableWidthPt: number;
  readonly maximumYPt: number;
  readonly suppressSpaceBefore: boolean;
  readonly wrap?: WrapOracle;
}

export interface TextMeasurer {
  readonly context:
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D;
  readonly fontFamilyClasses: Readonly<Record<string, string>>;
}

export interface MeasuredLine {
  readonly layout: LayoutLine;
  readonly topYPt: number;
  readonly advancePt: number;
}

export interface MeasuredParagraph {
  readonly lines: readonly MeasuredLine[];
  readonly markOnly: boolean;
  readonly requestedSpaceBeforePt: number;
  readonly requestedSpaceAfterPt: number;
  readonly contentStartYPt: number;
  readonly contentEndYPt: number;
  readonly placement: Readonly<ParagraphPlacement>;
}

export function measureParagraph(
  paragraph: DocParagraph,
  context: ParagraphLayoutContext,
  placement: ParagraphPlacement,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
): MeasuredParagraph;
```

The measurement coordinate system is points at scale 1. `contentStartYPt`
includes the effective leading spacing selected by the placement, while
`contentEndYPt` is the end of the measured line content and excludes trailing
paragraph spacing. The paginator owns trailing spacing collapse and records the
actual leading and trailing contributions on each fragment, preventing spacing
from being counted once in measurement and again in placement. Paint scales
measured geometry; it does not repeat text layout. A measurement is valid only
while its placement-DETERMINING inputs hold: the line partition and per-line
geometry are a function of the available width, the wrap context, and the
content — without a wrap oracle, the X/Y origin is a pure translation recorded
on the placed fragment, not a layout input. Moving a paragraph to a placement
with a different available width or a different wrap context therefore requires
deterministic remeasurement (the §17.6.4 unequal-width continuation re-wrap);
a same-width, wrap-free move to another page or column reuses the measurement
with the new origin recorded on `PlacedFragment` — the shipped continuation
contract in the fragment model section below.

`documentHasEastAsianText` preserves the existing document-level font-axis
choice for empty and anchor-only paragraph marks. Content-bearing lines continue
to use `ParagraphLayoutContext.hasEastAsianText`; the two inputs are intentionally
not interchangeable for an empty paragraph. Measurement callers populate this
required field from document layout settings.

`createFloatWrapOracle` in the paragraph-measurement module adapts the existing
float-layout pure functions to `WrapOracle`. Floating drawing discovery and
registration stay in the current anchor subsystem. A `wrapNone` object never
appears in the oracle. Table-cell measurement receives no page-level wrap
oracle.

Inline drawings remain line segments and continue to affect line width and line
height inside `measureParagraph`.

## Measured Fragment Model

Fragments belong to a layout result, not to the parsed document model.

```ts
export interface ParagraphFragment {
  readonly kind: 'paragraph';
  readonly source: DocParagraph;
  readonly measured: MeasuredParagraph;
  readonly lineStart: number;
  readonly lineEnd: number;
  readonly leadingSpacePt: number;
  readonly trailingSpacePt: number;
}

export interface CellFragment {
  readonly source: DocTableCell;
  readonly blocks: readonly FlowFragment[];
  readonly verticalMerge: 'none' | 'restart' | 'continue';
}

export interface RowFragment {
  readonly source: DocTableRow;
  readonly sourceRowIndex: number;
  readonly heightPt: number;
  readonly cells: readonly CellFragment[];
  readonly repeatedHeader: boolean;
}

export interface TableFragment {
  readonly kind: 'table';
  readonly source: DocTable;
  readonly columnWidthsPt: readonly number[];
  readonly rows: readonly RowFragment[];
  readonly continuesFromPreviousPage: boolean;
  readonly continuesOnNextPage: boolean;
}

export type FlowFragment = ParagraphFragment | TableFragment;

export interface PlacedFragment {
  readonly fragment: FlowFragment;
  readonly columnIndex: number;
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly section: SectionLayoutContext;
  readonly geometry: SectionGeom;
  readonly fragments: readonly PlacedFragment[];
}

export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
}
```

Paragraph continuation uses line ranges over one measured result only while the
placement remains valid. A continuation page with a different wrap context must
hold a separate measurement. Table continuation is recursive: a cell contains
paragraph or nested-table fragments, and each cell has an independent
continuation point.

Paint consumes `PlacedFragment` and page paint state. For migrated body/table
paths, paint must not call segment construction, line layout, text measurement,
or row measurement. Once modules are separated, an ast-grep rule will enforce
that boundary statically.

## Table Pagination Rules

The fragment paginator retains existing specification-backed table behavior:

- `cantSplit` rows move as a unit when possible;
- exact-height rows clip rather than split;
- automatic rows may split at cell block or paragraph-line boundaries;
- repeated header rows are placed on continuation pages;
- nested tables fragment recursively;
- vertical-merge geometry continues across row fragments;
- vertical-merge content remains owned by the restart cell;
- break selection never creates an invalid vertical-merge boundary;
- cell margins and paragraph spacing are measured once and included in row
  height.

The existing pure table geometry functions remain the source for column widths,
row-height rules, and vertical-merge span distribution. The migration replaces
model cloning and runtime stamps, not those validated algorithms.

`keepNext`, `keepLines`, and widow/orphan control are placement decisions over
measured line arrays. They are extracted as pure slice-selection functions and
must preserve current behavior before any semantic correction is considered.

## Data Flow

```text
Rust parser model
  -> document and section context resolvers
  -> story/container traversal
  -> paragraph and run context resolvers
  -> placement-aware paragraph measurement
  -> paragraph/table/row/cell fragments
  -> page and column placement
  -> fragment-only paint
```

Anchored drawing registration remains adjacent to page/column placement:

```text
existing anchor subsystem
  -> registered wrapping objects
  -> WrapOracle
  -> paragraph measurement
```

This boundary allows body measurement to use correct float constraints without
including anchor positioning or z-order migration in this delivery series.

## Compatibility Strategy

`paginateDocument` and existing renderer test helpers are internal but heavily
used. They remain available as compatibility adapters while the production path
moves to `DocumentLayout`. The adapters are removed only after body and table
tests consume the new result directly.

During extraction, existing layout results are the characterization baseline.
Normative behavior changes are isolated so visual differences can be attributed
to one rule. In particular, documents that previously received table-cell line
pitch unconditionally will change only in the dedicated compatibility-gating
delivery.

## Delivery Sequence

The work is delivered as a sequence of reviewable PRs. Each PR starts from the
then-current `main`, passes CI, and is merge-committed before the next dependent
stage is opened.

### PR 1: Preserve missing parser facts

- Parse `adjustLineHeightInTable`.
- Parse run-level `snapToGrid`.
- Add Rust and TypeScript wire fields.
- Do not change rendering behavior.

### PR 2: Introduce immutable layout resolvers

- Add document, section, paragraph, and run context modules.
- Replace duplicated state construction and physical bidi-indent resolution.
- Preserve existing layout semantics, including the temporarily documented
  table-cell line-grid deviation.
- Require zero visual difference.

### PR 3: Apply normative grid corrections

- Gate table-cell line pitch on `adjustLineHeightInTable`.
- Apply run `snapToGrid=false` to character pitch only.
- Treat `snapToChars` as both line and character grid.
- Keep `atLeast` in the line-grid policy and apply both minima.
- Verify each rule through synthetic matrices and local Office comparison.

### PR 4: Unify paragraph measurement

- Introduce placement-aware `measureParagraph`.
- Route body height estimation, paragraph page splitting, and cell paragraph
  measurement through it.
- Keep existing runtime stamps temporarily as compatibility adapters.
- Require old/new geometry equivalence outside PR 3's intentional changes.

### PR 5: Introduce body fragments

- Add `DocumentLayout`, paragraph fragments, and placed fragments.
- Make body pagination produce fragments.
- Make body paint consume measured lines without remeasurement.
- Remove paragraph runtime stamps from the migrated path.
- Add the static paint-boundary rule.

### PR 6: Introduce table fragments

- Add table, row, and cell fragments.
- Move row sizing, row splitting, nested-table splitting, repeated headers, and
  vertical merges to fragment ownership.
- Make table paint consume measured fragments.
- Remove table runtime stamps from the migrated path.

Text-box and anchored-drawing migrations are follow-up series. Text-box
migration must first preserve full `txbxContent` block structure instead of the
current reduced shape-text model.

## Tests and Invariants

### Parser tests

- `adjustLineHeightInTable`: absent, true by empty element, explicit true,
  explicit false.
- run `snapToGrid`: absent, explicit true, explicit false, and style inheritance.
- optional fields do not change unrelated serialized documents.

### Resolver matrix

Cover combinations of:

- grid type: absent, default, lines, linesAndChars, snapToChars;
- line pitch and character space: absent, positive, and valid negative character
  space;
- story: body, header, footer, note, text box contract;
- container stack: flow, table cell, nested table cell;
- compatibility flag: off and on;
- paragraph line rule: absent, inherited auto, explicit auto, exact, atLeast;
- paragraph and run `snapToGrid`: absent, true, false.

### Measurement invariants

1. Identical paragraph, context, placement, and measurer produce deeply equal
   measurements.
2. A measurement is never reused after page, column, width, or wrap placement
   changes.
3. Empty and anchor-only paragraphs retain one paragraph-mark line.
4. Inline drawings affect line geometry; `wrapNone` anchors do not enter the
   wrap oracle.
5. Scale-1 measured geometry scaled by `s` matches direct scale-`s` geometry
   within the established numeric tolerance.

### Pagination and paint invariants

1. Fragment height sums equal paginator cursor advancement.
2. Paint consumes exactly the fragment line ranges and row ranges selected by
   pagination.
3. Migrated paint modules cannot call measurement APIs.
4. Painting does not mutate the document model, layout result, or fragments.
5. Pagination row heights equal painted row heights.
6. Repeated table headers use equivalent measured geometry on every placement.
7. Vertical-merge content stays with the restart cell across page boundaries.
8. `cantSplit`, exact row height, nested-table continuation, `keepLines`,
   `keepNext`, and widow/orphan behavior retain their characterized semantics.

### Regression verification

- Focused resolver, line-box, pagination, table-split, and fragment tests.
- Full DOCX unit test suite.
- Rust formatting, clippy, and parser tests.
- TypeScript project build.
- Local-only visual comparison against representative private documents and
  Office-produced references without publishing their contents.
- Existing public visual references remain unchanged in behavior-preserving PRs.

## Error Handling

- Invalid or unsupported grid enum values resolve to `none`; they do not enable
  a guessed grid mode.
- Non-positive or non-finite pitches are inactive and covered by resolver tests.
- Missing Canvas font metrics continue through the existing deterministic font
  metric fallback.
- A fragment placement failure reports the responsible element and placement
  context in development diagnostics; it does not silently switch to a second
  paint-time layout algorithm.
- Compatibility behavior not defined by the standard requires explicit evidence
  and a separate decision before implementation.

## Acceptance Criteria

The body/table migration series is complete when:

1. the parser preserves the two missing OOXML settings;
2. body and table-cell paragraph policy is resolved through the same immutable
   context stack;
3. body and table-cell paragraphs use the same placement-aware measurement API;
4. table-cell line pitch follows `adjustLineHeightInTable`;
5. `snapToChars`, paragraph `snapToGrid`, and run `snapToGrid` affect only their
   specification-defined grid axes;
6. body and table pagination produce explicit fragments owned by
   `DocumentLayout`;
7. migrated paint paths do not remeasure paragraphs or rows;
8. existing non-target behavior remains covered by unit, type, Rust, and visual
   regression checks;
9. no private fixture name, content, or local environment path appears in public
   commits, documentation, or PR descriptions.
