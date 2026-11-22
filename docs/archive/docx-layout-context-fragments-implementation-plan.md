# DOCX Layout Context and Measured Fragments Implementation Plan

> **Completed historical plan (2026-08-17):** The immutable
> layout-to-paint migration described here has been delivered. The unchecked
> task boxes preserve the original execution plan and are not current backlog.
> Use [`docx-layout-engine-redesign.md`](../docx-layout-engine-redesign.md) for the
> current architecture and gates, and [`README.md`](README.md) for document
> classification.

**Goal:** Make DOCX body and table layout resolve OOXML policy through immutable contexts, share one placement-aware paragraph measurement API, and paginate and paint explicit measured fragments.

**Architecture:** Rust preserves style-resolved OOXML facts, while TypeScript resolves document, section, story, container, paragraph, and run layout policy. Measurement produces point-space geometry for a specific placement, pagination owns immutable fragments, and migrated paint paths scale those fragments without repeating line or row measurement.

**Tech Stack:** Rust 2021, serde, TypeScript 6, Canvas 2D, Vitest 4, pnpm 10, ast-grep, Storybook 10, Playwright.

## Global Constraints

- Start every PR from updated `main`; integrate only through a merge-commit PR after CI passes.
- Use TDD: focused failing test, observed failure, minimal implementation, focused pass, broader verification, refactor.
- Treat ECMA-376 Part 1 sections 17.3.1.32, 17.3.1.33, 17.3.2.34, 17.6.5, 17.15.3.1, 17.18.14, 20.4.2.3, and 20.4.2.15 as normative.
- Keep private fixtures, their filenames, their contents, local paths, and generated visual artifacts out of commits and PR text.
- Preserve non-target rendering in behavior-preserving PRs; intentional visual changes belong only to the normative grid-correction PR.
- Use optional serde fields so unrelated parser output remains unchanged.
- Do not add sample-specific thresholds, constants, or path branches.
- Every agent-authored commit includes the actual model in its co-author trailer.
- Build WASM before local Storybook or private visual verification.

---

## PR 1: Preserve Missing OOXML Facts

Branch: `codex/docx-layout-context`.

### Task 1: Parse the table line-height compatibility flag

**Files:**
- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Modify: `packages/docx/src/types.ts`

**Interfaces:**
- Produces Rust field: `DocumentSettings.adjust_line_height_in_table: Option<bool>`.
- Produces wire field: `DocSettings.adjustLineHeightInTable?: boolean`.
- Rendering behavior remains unchanged in this PR.

- [ ] **Step 1: Add failing parser tests**

Add tests beside `settings_east_asian_compat_flags_surface`:

```rust
#[test]
fn settings_adjust_line_height_in_table_surfaces() {
    let xml = format!(
        r#"<w:settings xmlns:w="{w}"><w:compat><w:adjustLineHeightInTable/></w:compat></w:settings>"#,
        w = W_NS,
    );
    let settings = parse_document_settings(&xml).expect("compat setting");
    assert_eq!(settings.adjust_line_height_in_table, Some(true));
}

#[test]
fn settings_adjust_line_height_in_table_false_surfaces() {
    let xml = format!(
        r#"<w:settings xmlns:w="{w}"><w:compat><w:adjustLineHeightInTable w:val="0"/></w:compat></w:settings>"#,
        w = W_NS,
    );
    let settings = parse_document_settings(&xml).expect("compat setting");
    assert_eq!(settings.adjust_line_height_in_table, Some(false));
}
```

- [ ] **Step 2: Run the focused test and observe the missing-field failure**

Run:

```bash
cargo test -p docx-parser settings_adjust_line_height_in_table -- --nocapture
```

Expected: compile failure because `adjust_line_height_in_table` does not exist.

- [ ] **Step 3: Add the optional Rust and TypeScript fields**

Add to `DocumentSettings`:

```rust
/// ECMA-376 §17.15.3.1: apply the section line grid inside table cells.
#[serde(skip_serializing_if = "Option::is_none")]
pub adjust_line_height_in_table: Option<bool>,
```

Add to `DocSettings`:

```ts
/** §17.15.3.1: apply section line pitch inside table cells. */
adjustLineHeightInTable?: boolean;
```

- [ ] **Step 4: Parse and materialize the compatibility flag**

In `parse_document_settings`, add:

```rust
let adjust_line_height_in_table = compat_bool("adjustLineHeightInTable");
```

Include the field in both the all-fields-absent check and `DocumentSettings` construction.

- [ ] **Step 5: Run parser verification**

Run:

```bash
cargo test -p docx-parser settings_adjust_line_height_in_table -- --nocapture
cargo fmt --all --check
```

Expected: both focused tests pass and formatting is clean.

- [ ] **Step 6: Commit the setting parser change**

Commit subject:

```text
feat(docx): preserve table line-grid compatibility setting
```

The body records the missing parser fact, ECMA-376 §17.15.3.1, unchanged rendering, focused tests, and the actual co-author model.

### Task 2: Parse run-level character-grid participation

**Files:**
- Modify: `packages/docx/parser/src/styles.rs`
- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Modify: `packages/docx/src/types.ts`

**Interfaces:**
- Produces `RunFmt.snap_to_grid: Option<bool>`.
- Produces `TextRun.snap_to_grid: Option<bool>` and wire field `DocxTextRun.snapToGrid?: boolean`.
- The field controls character-grid participation only in a later PR.

- [ ] **Step 1: Add failing direct and inherited run tests**

Use the existing `run_of` and `run_of_with_styles` helpers:

```rust
#[test]
fn run_snap_to_grid_false_surfaces() {
    let run = run_of(r#"<w:r><w:rPr><w:snapToGrid w:val="0"/></w:rPr><w:t>x</w:t></w:r>"#);
    assert_eq!(run.snap_to_grid, Some(false));
}

#[test]
fn run_snap_to_grid_inherits_from_character_style() {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="character" w:styleId="NoCharGrid"><w:rPr><w:snapToGrid w:val="0"/></w:rPr></w:style>
    </w:styles>"#;
    let run = run_of_with_styles(
        styles,
        r#"<w:r><w:rPr><w:rStyle w:val="NoCharGrid"/></w:rPr><w:t>x</w:t></w:r>"#,
    );
    assert_eq!(run.snap_to_grid, Some(false));
}
```

- [ ] **Step 2: Run the focused test and observe the missing-field failure**

Run:

```bash
cargo test -p docx-parser run_snap_to_grid -- --nocapture
```

Expected: compile failure because run formatting and `TextRun` lack the field.

- [ ] **Step 3: Extend run formatting and its canonical merge**

Add to `RunFmt`:

```rust
/// ECMA-376 §17.3.2.34: run participation in the character grid.
pub snap_to_grid: Option<bool>,
```

In `parse_run_fmt` assign:

```rust
fmt.snap_to_grid = bool_prop(rpr, "snapToGrid");
```

In `apply_run` copy only explicit values:

```rust
if src.snap_to_grid.is_some() {
    dst.snap_to_grid = src.snap_to_grid;
}
```

- [ ] **Step 4: Carry the resolved value to the wire model**

Add to Rust `TextRun`:

```rust
#[serde(skip_serializing_if = "Option::is_none")]
pub snap_to_grid: Option<bool>,
```

Populate it from the resolved `RunFmt` at every `TextRun` construction path. Add to TypeScript `DocxTextRun`:

```ts
/** §17.3.2.34: false opts this run out of character-grid spacing. */
snapToGrid?: boolean;
```

- [ ] **Step 5: Verify parser, formatting, clippy, and TypeScript**

Run:

```bash
cargo test -p docx-parser run_snap_to_grid -- --nocapture
cargo fmt --all --check
cargo clippy -p docx-parser --all-targets -- -D warnings
pnpm --filter @silurus/ooxml-docx typecheck
```

Expected: focused tests pass; formatting, clippy, and typecheck exit 0.

- [ ] **Step 6: Commit the run parser change**

Commit subject:

```text
feat(docx): preserve run character-grid participation
```

### Task 3: Verify and merge PR 1

**Files:**
- Verify: `docs/archive/docx-layout-context-fragments-design.md`
- Verify: `docs/archive/docx-layout-context-fragments-implementation-plan.md`
- Verify: parser and TypeScript files changed in Tasks 1-2.

**Interfaces:**
- PR 1 changes parser facts only; renderer output must remain unchanged.

- [ ] **Step 1: Run full PR verification**

```bash
cargo fmt --all --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test -p docx-parser
pnpm --filter @silurus/ooxml-docx typecheck
pnpm vitest run packages/docx/src/line-box-height.test.ts packages/docx/src/docgrid-char.test.ts packages/docx/src/pagination.test.ts packages/docx/src/table-split.test.ts
git diff --check main...HEAD
```

Expected: all commands exit 0.

- [ ] **Step 2: Push and create the public PR**

Use an OSS-safe title such as `Preserve DOCX layout compatibility settings`. The body describes parser information loss, relevant specification sections, unchanged renderer behavior, and sanitized commands only.

- [ ] **Step 3: Wait for all GitHub Actions checks**

Run:

```bash
gh pr checks --watch
```

Expected: every required check passes.

- [ ] **Step 4: Merge through the PR**

Run:

```bash
gh pr merge --merge
```

Expected: merge commit created on `main`; no direct push.

---

## PR 2: Introduce Immutable Layout Resolvers

Branch: `codex/docx-layout-resolvers` from updated `main`.

### Task 4: Add pure document, section, paragraph, and run resolvers

**Files:**
- Create: `packages/docx/src/layout-context.ts`
- Create: `packages/docx/src/layout-context.test.ts`
- Modify: `packages/docx/src/types.ts`

**Interfaces:**
- Produces the context types and resolver functions approved in the design.
- Does not depend on `renderer.ts` or mutable `RenderState`.

- [ ] **Step 1: Add failing resolver matrix tests**

Define complete local builders in the new test file:

```ts
const section = (overrides: Partial<SectionProps> = {}): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  titlePage: false,
  evenAndOddHeaders: false,
  ...overrides,
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 12,
  indentRight: 6,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const bodyStory: StoryContext = {
  story: 'body',
  containers: [],
  lineNumberingEligible: true,
};
const cellStory: StoryContext = {
  story: 'body',
  containers: [{ kind: 'tableCell' }],
  lineNumberingEligible: false,
};
```

Construct settings through `resolveDocumentLayoutSettings` from a minimal document, then assert:

```ts
const layoutSettings = (adjustLineHeightInTable = false): DocumentLayoutSettings => ({
  kinsoku: DEFAULT_KINSOKU_RULES,
  defaultTabPt: 36,
  documentHasEastAsianText: true,
  compat: {
    adjustLineHeightInTable,
    useFeLayout: false,
    balanceSingleByteDoubleByteWidth: false,
  },
});

const defaultGrid = resolveSectionLayoutContext(
  layoutSettings(),
  section({ docGridType: 'default', docGridLinePitch: 20 }),
);
expect(defaultGrid.grid.kind).toBe('none');

const snapGrid = resolveSectionLayoutContext(
  layoutSettings(),
  section({ docGridType: 'snapToChars', docGridLinePitch: 20 }),
);
expect(snapGrid.grid.kind).toBe('snapToChars');

const bidiParagraph = paragraph({ bidi: true });
const bidiContext = resolveParagraphLayoutContext(
  layoutSettings(),
  snapGrid,
  bodyStory,
  bidiParagraph,
);
expect(bidiContext.physicalIndentLeftPt).toBe(bidiParagraph.indentRight);

const cellContext = resolveParagraphLayoutContext(
  layoutSettings(false),
  snapGrid,
  cellStory,
  paragraph(),
);
expect(cellContext.lineGrid.active).toBe(false);

const textRun = {
  text: 'x', bold: false, italic: false, underline: false,
  strikethrough: false, fontSize: 10, color: null, fontFamily: null,
  isLink: false, background: null, vertAlign: null, hyperlink: null,
  snapToGrid: false,
} satisfies DocxTextRun;
expect(resolveRunLayoutContext(bidiContext, textRun).characterGrid.active).toBe(false);
```

- [ ] **Step 2: Run the tests and observe the missing-module failure**

```bash
pnpm vitest run packages/docx/src/layout-context.test.ts
```

Expected: import failure for `layout-context.js`.

- [ ] **Step 3: Implement immutable context types and resolvers**

Export these contracts from `layout-context.ts`:

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

export interface SectionGridContext {
  readonly kind: 'none' | 'lines' | 'linesAndChars' | 'snapToChars';
  readonly linePitchPt: number | null;
  readonly charSpacePt: number | null;
}

export interface LineGridPolicy {
  readonly active: boolean;
  readonly pitchPt: number | null;
}

export interface CharacterGridPolicy {
  readonly active: boolean;
  readonly deltaPt: number;
}

export type StoryKind =
  | 'body'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'textbox';

export type ContainerFrame = { readonly kind: 'tableCell' };

export interface StoryContext {
  readonly story: StoryKind;
  readonly containers: readonly ContainerFrame[];
  readonly lineNumberingEligible: boolean;
}

export interface SectionLayoutContext {
  readonly geometry: SectionGeom;
  readonly columns: readonly ColumnGeom[];
  readonly grid: SectionGridContext;
  readonly textDirection: string;
  readonly verticalAlignment: string;
  readonly lineNumbering?: LineNumbering;
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

export function resolveDocumentLayoutSettings(
  document: DocxDocumentModel,
): DocumentLayoutSettings;

export function resolveSectionLayoutContext(
  settings: DocumentLayoutSettings,
  section: SectionProps,
): SectionLayoutContext;

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

Implement the policy classifiers as pure functions:

```ts
export function isSectionLineGrid(kind: SectionGridContext['kind']): boolean {
  return kind === 'lines' || kind === 'linesAndChars' || kind === 'snapToChars';
}

export function isSectionCharacterGrid(kind: SectionGridContext['kind']): boolean {
  return kind === 'linesAndChars' || kind === 'snapToChars';
}
```

`resolveParagraphLayoutContext` applies paragraph `snapToGrid`, `lineRule=exact`, and table-cell compatibility gating. `resolveRunLayoutContext` applies only run character-grid gating. While legacy callers still require `DocGridCtx`, map `SectionGridContext.kind` to `DocGridCtx.type` without renaming or reinterpreting values; remove that bridge when all callers consume the resolved policies.

Move the existing recursive `documentHasEastAsian` helper into the resolver module and call it once from `resolveDocumentLayoutSettings`; do not add a second body scan at renderer entry points.

- [ ] **Step 4: Run resolver tests and typecheck**

```bash
pnpm vitest run packages/docx/src/layout-context.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
```

Expected: resolver tests and typecheck pass.

- [ ] **Step 5: Commit the pure resolver kernel**

Commit subject:

```text
feat(docx): add immutable layout context resolvers
```

### Task 5: Route body layout through resolvers without changing output

**Files:**
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/line-layout.ts`
- Modify: `packages/docx/src/layout-context.test.ts`
- Modify: `packages/docx/src/pagination.test.ts`

**Interfaces:**
- `RenderState` temporarily carries resolved document/section contexts.
- Body paragraph code reads physical indents from `ParagraphLayoutContext` and carries its grid policy for later use.
- Table cells remain on the existing path until PR 3, isolating intentional visual changes.

- [ ] **Step 1: Add a failing body resolver integration test**

Add a synthetic bidi paragraph with unequal logical indents and assert pagination and paint use the same physical width and line count with resolver instrumentation enabled.

- [ ] **Step 2: Run the focused integration tests**

```bash
pnpm vitest run packages/docx/src/pagination.test.ts packages/docx/src/measure-column-geometry.test.ts packages/docx/src/rtl-tab-stops.test.ts
```

Expected: the new resolver-use assertion fails before integration.

- [ ] **Step 3: Replace duplicate body policy resolution**

At document entry, resolve document settings once. At each active section, resolve section context once. For body paragraphs, replace direct calls that swap bidi indents or construct `DocGridCtx` with:

```ts
const paragraphContext = resolveParagraphLayoutContext(
  state.layoutSettings,
  state.sectionLayout,
  BODY_STORY_CONTEXT,
  para,
);
```

Keep current line metrics and float functions unchanged. Do not route table-cell paragraphs yet.
In particular, line-box height must continue through the existing `DocGridCtx` and `isGridLineRule` behavior in this PR. Do not consume `ParagraphLayoutContext.lineGrid.active` until PR 3, because the resolver already expresses the corrected `snapToChars` semantics while this PR is required to remain pixel-identical.

- [ ] **Step 4: Run focused and broad renderer tests**

```bash
pnpm vitest run packages/docx/src/layout-context.test.ts packages/docx/src/pagination.test.ts packages/docx/src/measure-column-geometry.test.ts packages/docx/src/rtl-tab-stops.test.ts packages/docx/src/line-box-height.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
git diff --check
```

Expected: all tests pass with unchanged expectations.

- [ ] **Step 5: Commit the body integration**

Commit subject:

```text
refactor(docx): resolve body layout policy once
```

### Task 6: Verify and merge PR 2

**Files:** all PR 2 files.

**Interfaces:** PR 2 is behavior-preserving and requires zero public VRT difference.

- [ ] **Step 1: Run unit, type, build, and local visual verification**

```bash
pnpm --filter @silurus/ooxml-docx typecheck
pnpm vitest run packages/docx/src/layout-context.test.ts packages/docx/src/pagination.test.ts packages/docx/src/line-box-height.test.ts packages/docx/src/docgrid-char.test.ts packages/docx/src/table-split.test.ts
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
git diff --check main...HEAD
```

Expected: all commands pass; behavior-preserving visual references have no changed pixels.

- [ ] **Step 2: Push, open an OSS-safe PR, wait for CI, and merge with `--merge`**

The PR describes duplicated policy resolution and immutable resolver adoption. It does not mention local-only fixture identities or content.

---

## PR 3: Apply Normative Document-Grid Corrections

Branch: `codex/docx-grid-context-fixes` from updated `main`.

### Task 7: Correct line-grid and character-grid participation

**Files:**
- Modify: `packages/docx/src/layout-context.ts`
- Modify: `packages/docx/src/layout-context.test.ts`
- Modify: `packages/docx/src/line-layout.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/line-box-height.test.ts`
- Modify: `packages/docx/src/docgrid-char.test.ts`
- Create: `packages/docx/src/table-cell-docgrid.test.ts`

**Interfaces:**
- `snapToChars` participates in both line and character grids.
- Table cells receive line pitch only when `adjustLineHeightInTable` is true.
- Paragraph `snapToGrid=false` and exact line spacing disable line pitch only.
- Run `snapToGrid=false` disables character pitch for that run only.
- `atLeast` preserves both the authored minimum and active grid minimum.

- [ ] **Step 1: Add failing normative matrix tests**

Pin the following results using public kernel functions and the builders from `layout-context.test.ts`:

```ts
expect(isGridLineRule({ type: 'snapToChars', linePitchPt: 20 })).toBe(true);

const cellWithoutCompat = resolveParagraphLayoutContext(
  layoutSettings(false),
  sectionContext({ docGridType: 'lines', docGridLinePitch: 20 }),
  cellStory,
  paragraph(),
);
expect(cellWithoutCompat.lineGrid.active).toBe(false);

const cellWithCompat = resolveParagraphLayoutContext(
  layoutSettings(true),
  sectionContext({ docGridType: 'lines', docGridLinePitch: 20 }),
  cellStory,
  paragraph(),
);
expect(cellWithCompat.lineGrid.active).toBe(true);

expect(lineBoxHeight({ value: 18, rule: 'atLeast', explicit: true }, 10, 2, 1, grid20)).toBe(20);

expect(segAdvanceWidth({ ...eaSegment, snapToCharacterGrid: true }, 20, 1, 1)).toBe(21);
expect(segAdvanceWidth({ ...eaSegment, snapToCharacterGrid: false }, 20, 1, 1)).toBe(20);
```

Define the remaining values in the test without implicit helpers:

```ts
const grid20 = { type: 'lines', linePitchPt: 20 } as const;
const eaSegment: LayoutTextSeg = {
  text: 'あ',
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontSize: 20,
  color: null,
  fontFamily: null,
  vertAlign: null,
  measuredWidth: 0,
};
```

Keep this PR's tests self-contained by defining exact local builders:

```ts
const layoutSettings = (adjustLineHeightInTable = false): DocumentLayoutSettings => ({
  kinsoku: DEFAULT_KINSOKU_RULES,
  defaultTabPt: 36,
  documentHasEastAsianText: true,
  compat: {
    adjustLineHeightInTable,
    useFeLayout: false,
    balanceSingleByteDoubleByteWidth: false,
  },
});

const section = (overrides: Partial<SectionProps> = {}): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  titlePage: false,
  evenAndOddHeaders: false,
  ...overrides,
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 12,
  indentRight: 6,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const cellStory: StoryContext = {
  story: 'body',
  containers: [{ kind: 'tableCell' }],
  lineNumberingEligible: false,
};

const sectionContext = (overrides: Partial<SectionProps> = {}) =>
  resolveSectionLayoutContext(layoutSettings(), section(overrides));
```

- [ ] **Step 2: Run and observe the four independent failures**

```bash
pnpm vitest run packages/docx/src/layout-context.test.ts packages/docx/src/line-box-height.test.ts packages/docx/src/docgrid-char.test.ts packages/docx/src/table-cell-docgrid.test.ts
```

Expected: failures identify snapToChars line activation, table-cell gating, atLeast grid minimum, and per-run character spacing.

- [ ] **Step 3: Route table cells through a table-cell story context**

Create the cell story by appending `{ kind: 'tableCell' }` to the current container stack. Pass its resolved paragraph context through row measurement, row splitting, and cell paint. Remove direct table-cell use of the section `state.docGrid`.

- [ ] **Step 4: Gate character-grid delta per text segment**

Add to `LayoutTextSeg`:

```ts
snapToCharacterGrid?: boolean;
```

Set it from `DocxTextRun.snapToGrid !== false` in `buildSegments`. Gate the delta inside the shared advance helpers so every caller observes the same rule:

```ts
const segmentGridDelta = seg.snapToCharacterGrid === false ? 0 : gridDeltaPx;
```

Use `segmentGridDelta` inside `segAdvanceWidth` and `segLetterSpacingPx` before calling `gridSegDeltaPx`; make `gridWidth` accept the already-gated delta or add the same segment-aware gate to its caller. The paint letter-spacing path must call the same helper rather than duplicate the condition.

- [ ] **Step 5: Correct line-grid activation and atLeast minimum**

Include `snapToChars` in `isGridLineRule`. For `atLeast`, return the maximum of natural height, authored minimum, and active grid-resolved single-line height. Keep exact spacing independent of a positive grid.

- [ ] **Step 6: Run focused tests, full DOCX tests, and local comparison**

```bash
pnpm vitest run packages/docx/src/layout-context.test.ts packages/docx/src/line-box-height.test.ts packages/docx/src/docgrid-char.test.ts packages/docx/src/table-cell-docgrid.test.ts packages/docx/src/pagination.test.ts packages/docx/src/table-split.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
```

Expected: unit/type checks pass. Any intentional local visual difference is attributable to one of the four normative rules and is reviewed without updating references automatically.

- [ ] **Step 7: Commit each behavior correction separately**

Use these subjects, omitting a commit only if its code is inseparable from the preceding one:

```text
fix(docx): gate table cell line pitch by compatibility
fix(docx): honor run character-grid participation
fix(docx): apply snap-to-chars line pitch
fix(docx): retain the grid minimum for at-least spacing
```

### Task 8: Verify and merge PR 3

- [ ] **Step 1: Run full verification and review intentional visual differences**

```bash
cargo test -p docx-parser
pnpm vitest run packages/docx/src/layout-context.test.ts packages/docx/src/line-box-height.test.ts packages/docx/src/docgrid-char.test.ts packages/docx/src/table-cell-docgrid.test.ts packages/docx/src/pagination.test.ts packages/docx/src/table-split.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
git diff --check main...HEAD
```

Expected: all non-visual checks pass. Review each visual difference against the four normative rule changes without updating references automatically.

- [ ] **Step 2: Gate public visual-reference changes on explicit approval**

List every changed tracked visual reference and attribute it to one of the four normative corrections. Do not regenerate, stage, or merge a changed public reference without explicit user approval. Local-only comparison artifacts remain untracked.

- [ ] **Step 3: Ask Opus 4.8 for a read-only specification review**

Provide only the public diff and specification sections. Require findings with file/line references and prohibit edits, private fixture names, and local paths.

- [ ] **Step 4: Resolve valid findings, rerun verification, push, open the PR, wait for CI, and merge with `--merge`**

The PR body explains each OOXML mismatch and its general document class.

---

## PR 4: Unify Placement-Aware Paragraph Measurement

Branch: `codex/docx-paragraph-measurement` from updated `main`.

### Task 9: Add the measurement kernel

**Files:**
- Create: `packages/docx/src/paragraph-measure.ts`
- Create: `packages/docx/src/paragraph-measure.test.ts`
- Modify: `packages/docx/src/line-layout.ts`
- Modify: `packages/docx/src/float-layout.ts`

**Interfaces:**
- Produces the placement-aware contracts below without importing renderer state.
- A result is valid only for its recorded placement.
- Coordinates and measurement use scale 1 points.

- [ ] **Step 1: Add failing placement tests**

Test no-float, float-window, empty-mark, anchor-only, ruby, bidi, inline-image, exact spacing, and changed-start-Y cases. The changed placement test asserts two calls with different `startYPt` do not share the same result object and can produce different line windows.

- [ ] **Step 2: Run and observe the missing-module failure**

```bash
pnpm vitest run packages/docx/src/paragraph-measure.test.ts
```

- [ ] **Step 3: Implement a line-layout environment independent of paint state**

Extract the fields read by segment construction and field resolution into a narrow interface:

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
```

Change `buildSegments` and `resolveFieldText` to accept `LineLayoutEnvironment`. The current `buildSegments` implementation reads only `verticalCJK`, note numbering, and `resolveFieldText`'s page/date fields from state; keep font classes on `TextMeasurer` and grid, kinsoku, tabs, and spacing on `ParagraphLayoutContext`. `ParagraphMeasurementEnvironment.documentHasEastAsianText` preserves the existing document-level font-axis choice for empty and anchor-only paragraph marks; content lines continue to use `ParagraphLayoutContext.hasEastAsianText`. Measurement callers must populate it from document layout settings. Verify this boundary with a `state.` usage scan and typecheck.

Define the measurement boundary explicitly:

```ts
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

export interface TextMeasurer {
  readonly context:
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D;
  readonly fontFamilyClasses: Readonly<Record<string, string>>;
}

export interface ParagraphPlacement {
  readonly startYPt: number;
  readonly paragraphXPt: number;
  readonly availableWidthPt: number;
  readonly maximumYPt: number;
  readonly suppressSpaceBefore: boolean;
  readonly wrap?: WrapOracle;
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

- [ ] **Step 4: Implement `WrapOracle` over existing pure float functions**

The adapter delegates to `resolveLineFloatWindow` and `skipPastTopAndBottom`. Keep object registration and one-inch compatibility behavior outside the interface.

- [ ] **Step 5: Implement `measureParagraph` by moving existing calculations verbatim**

Use `buildSegments`, `layoutLines`, `lineBoxHeight`, `paragraphMarkLineHeight`, and the resolved paragraph context. Return measured line top and advance values plus content start/end; exclude trailing paragraph spacing.

- [ ] **Step 6: Run measurement tests and typecheck**

```bash
pnpm vitest run packages/docx/src/paragraph-measure.test.ts packages/docx/src/line-box-height.test.ts packages/docx/src/layout-lines-scale-invariance.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
```

- [ ] **Step 7: Commit the measurement kernel**

```text
refactor(docx): add placement-aware paragraph measurement
```

### Task 10: Route body pagination and table-cell measurement through one API

**Files:**
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-layout.ts`
- Modify: `packages/docx/src/pagination.test.ts`
- Modify: `packages/docx/src/cell-paragraph-line-reuse.test.ts`
- Modify: `packages/docx/src/paginate-paint-line-count.test.ts`
- Modify: `packages/docx/src/float-line-start-one-inch.test.ts`

**Interfaces:**
- `estimateParagraphHeight`, `splitParagraphAcrossPages`, `measureParaHeight`, and `layoutCellParagraphForRowSplit` consume `measureParagraph`.
- Runtime line stamps remain temporarily for paint compatibility.

- [ ] **Step 1: Add old/new geometry equivalence tests**

For representative synthetic paragraphs, capture page count, line ranges, top positions, and cell heights. Add a test-only legacy toggle so the same model can run through old and new callers during migration.

- [ ] **Step 2: Run and verify the equivalence test fails before routing**

```bash
pnpm vitest run packages/docx/src/pagination.test.ts packages/docx/src/cell-paragraph-line-reuse.test.ts packages/docx/src/paginate-paint-line-count.test.ts
```

- [ ] **Step 3: Replace each mirrored calculation with `measureParagraph`**

Body estimation passes the active wrap oracle and placement. Cell measurement passes the table-cell context with no page wrap oracle. Paragraph splitting slices `MeasuredLine[]` and applies `keepLines` and widow/orphan rules without recalculating line advances.

Preserving the first line's measured top displacement can expose page-scoped square floats registered in another newspaper column. Per §20.4.2.17, apply square wrapping only when the virtual wrap rectangle intersects the paragraph's horizontal line area; `wrapText` then selects the permitted side within that intersecting area.

- [ ] **Step 4: Remove the test-only legacy toggle after equality passes**

Keep characterization assertions, remove the alternate production code path, and ensure only one line measurement implementation remains.

- [ ] **Step 5: Run focused and broad verification**

```bash
pnpm vitest run packages/docx/src/paragraph-measure.test.ts packages/docx/src/pagination.test.ts packages/docx/src/cell-paragraph-line-reuse.test.ts packages/docx/src/paginate-paint-line-count.test.ts packages/docx/src/table-split.test.ts
pnpm vitest run packages/docx/src/float-line-start-one-inch.test.ts packages/docx/src/float-table-page-fit.test.ts
pnpm vitest run packages/docx/src
pnpm --filter @silurus/ooxml-docx typecheck
git diff --check
```

- [ ] **Step 6: Commit the caller migration**

```text
refactor(docx): share paragraph measurement across flow paths
```

### Task 11: Verify and merge PR 4

- [ ] **Step 1: Run full DOCX tests, typecheck, WASM build, and local VRT**

```bash
cargo test -p docx-parser
pnpm test
pnpm typecheck
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
git diff --check main...HEAD
```

Expected: every command exits 0 and behavior-preserving visual references are unchanged.

- [ ] **Step 2: Ask Sonnet 5 for a read-only measure/paint divergence review**

Run Claude Code with `--model claude-sonnet-5`, read-only tools, and a prompt that requests findings with file/line references for duplicate measurement, placement reuse, field resolution, and missing tests. Do not grant edit or network tools.

- [ ] **Step 3: Resolve findings, rerun checks, push, open an OSS-safe PR, wait for CI, and merge with `--merge`**

Rerun Step 1 after every accepted finding. Then use `git push -u origin codex/docx-paragraph-measurement`, create the PR, `gh pr checks --watch`, and `gh pr merge --merge`.

---

## PR 5: Produce and Paint Body Fragments

Branch: `codex/docx-body-fragments` from updated `main`.

### Task 12: Add fragment types and body layout output

**Files:**
- Create: `packages/docx/src/layout-fragments.ts`
- Create: `packages/docx/src/document-layout.ts`
- Create: `packages/docx/src/document-layout.test.ts`
- Modify: `packages/docx/src/types.ts`
- Modify: `packages/docx/src/renderer.ts`

**Interfaces:**
- Produces `ParagraphFragment`, `PlacedFragment`, `LayoutPage`, and `DocumentLayout`.
- Adds `layoutDocument(doc): DocumentLayout`.
- Keeps `paginateDocument` as a temporary compatibility adapter.

- [ ] **Step 1: Add failing body fragment tests**

Assert source identity, immutable line ranges, placement coordinates, page geometry, section context, paragraph continuation, and changed-wrap remeasurement. Add an invariant that cursor advancement equals `leadingSpacePt + measured line advances + trailingSpacePt`, proving paragraph spacing is owned and added exactly once.

- [ ] **Step 2: Run and observe the missing-layout failure**

```bash
pnpm vitest run packages/docx/src/document-layout.test.ts
```

- [ ] **Step 3: Implement immutable fragment and page types**

Export the initial body-fragment contracts explicitly:

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

export type FlowFragment = ParagraphFragment;

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

PR 6 adds `TableFragment` to `FlowFragment` when its complete contract is introduced. Freeze layout results in tests and ensure no field is added to `DocParagraph`.

- [ ] **Step 4: Make body pagination emit `DocumentLayout`**

Move page, column, and section stamps from parsed elements into `LayoutPage` and `PlacedFragment`. Use `ParagraphFragment.leadingSpacePt` and `trailingSpacePt` for spacing collapse and continuation ownership.

- [ ] **Step 5: Implement the compatibility adapter**

`paginateDocument` projects fragments to the current `PaginatedBodyElement[][]` shape only for unmigrated callers. Production rendering uses `DocumentLayout`.

- [ ] **Step 6: Run body layout tests and existing pagination tests**

```bash
pnpm vitest run packages/docx/src/document-layout.test.ts packages/docx/src/pagination.test.ts packages/docx/src/paginate-column-anchor.test.ts packages/docx/src/page-anchor-prescan.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
```

- [ ] **Step 7: Commit body fragment production**

```text
refactor(docx): paginate body paragraphs as fragments
```

### Task 13: Paint body fragments without remeasurement

**Files:**
- Create: `packages/docx/src/fragment-paint.ts`
- Create: `packages/docx/src/fragment-paint.test.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/layout-lines-reuse-identity.test.ts`
- Create: `rules/no-docx-measurement-in-fragment-paint.yml`
- Create: `rule-tests/no-docx-measurement-in-fragment-paint-test.yml`
- Modify: `sgconfig.yml` only if the existing rule directory configuration requires no automatic discovery.

**Interfaces:**
- `paintParagraphFragment(fragment, pageState)` scales stored geometry only.
- The body fragment paint module cannot import or call `buildSegments`, `layoutLines`, `measureParagraph`, `measureText`, or row measurement.

- [ ] **Step 1: Add a failing paint-purity test**

Use a Canvas stub whose `measureText` throws. Paint a premeasured paragraph fragment and assert text draw calls and coordinates complete without invoking measurement.

- [ ] **Step 2: Run and observe current paint-time measurement failure**

```bash
pnpm vitest run packages/docx/src/fragment-paint.test.ts
```

- [ ] **Step 3: Implement fragment-only paragraph paint**

Move line drawing inputs needed by body paint into the fragment or resolved page paint context. Draw only `lineStart..lineEnd`; draw paragraph-level anchors once on the first fragment while retaining the existing anchor subsystem.

- [ ] **Step 4: Add the ast-grep boundary rule and rule tests**

The rule matches prohibited imports or calls inside `packages/docx/src/fragment-paint.ts`. Its valid fixture paints supplied geometry; invalid fixtures import `layoutLines` or call `measureText`.

- [ ] **Step 5: Remove paragraph runtime stamps from production body paint**

Delete migrated use of `layoutLinesInputs` and stamped `LayoutLine[]`. Keep compatibility fields only where table paint still needs them before PR 6.

- [ ] **Step 6: Run paint, lint, type, and identity tests**

```bash
pnpm vitest run packages/docx/src/fragment-paint.test.ts packages/docx/src/layout-lines-reuse-identity.test.ts packages/docx/src/paginate-paint-line-count.test.ts
pnpm lint
pnpm lint:test
pnpm --filter @silurus/ooxml-docx typecheck
```

- [ ] **Step 7: Commit body fragment paint and static enforcement**

```text
refactor(docx): paint body fragments without remeasurement
```

### Task 14: Verify and merge PR 5

- [ ] **Step 1: Run full DOCX tests, lint, typecheck, WASM build, and local VRT**

```bash
pnpm test
pnpm typecheck
pnpm lint
pnpm lint:test
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
git diff --check main...HEAD
```

Expected: every command exits 0; body fragment paint invokes no measurement and visual references remain unchanged.

- [ ] **Step 2: Ask Opus 4.8 for a read-only fragment ownership and immutability review**

Run Claude Code with `--model claude-opus-4-8`, read-only tools, and a prompt focused on model mutation, stale placement reuse, hidden paint-time measurement, and missing section/page geometry. Require file/line findings.

- [ ] **Step 3: Resolve findings, rerun checks, push, open an OSS-safe PR, wait for CI, and merge with `--merge`**

Rerun Step 1 after every accepted finding. Then use `git push -u origin codex/docx-body-fragments`, create the PR, `gh pr checks --watch`, and `gh pr merge --merge`.

---

## PR 6: Produce and Paint Table Fragments

Branch: `codex/docx-table-fragments` from updated `main`.

### Task 15: Fragment table rows and cell content

**Files:**
- Modify: `packages/docx/src/layout-fragments.ts`
- Create: `packages/docx/src/table-fragments.ts`
- Create: `packages/docx/src/table-fragments.test.ts`
- Modify: `packages/docx/src/document-layout.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/table-split.test.ts`

**Interfaces:**
- Produces `CellFragment`, `RowFragment`, and `TableFragment`.
- Reuses `resolveTableRowHeights`, `resolveSingleRowHeight`, and `findMergeEndRow`.
- Cell content recursively contains paragraph or nested-table fragments.

- [ ] **Step 1: Add failing table fragment tests**

Cover ordinary row breaks, `cantSplit`, exact height, auto-height line splits, block splits, nested tables, repeated headers, vertical merges, and independent per-cell continuation points.

- [ ] **Step 2: Run and observe missing table fragment behavior**

```bash
pnpm vitest run packages/docx/src/table-fragments.test.ts packages/docx/src/table-split.test.ts
```

- [ ] **Step 3: Implement recursive cell and row measurement**

Measure each cell with the table-cell story context and no page wrap oracle. Construct immutable cell blocks, row heights, and vertical-merge state without cloning `DocTableRow` or adding `lineSlice` fields.

- [ ] **Step 4: Implement page slicing over row fragments**

Preserve `cantSplit`, exact-height clipping, repeated-header placement, nested-table continuation, and vertical-merge-safe boundaries. Each continuation table records both continuation flags.

- [ ] **Step 5: Route `layoutDocument` table elements to table fragments**

Place table fragments in the same page and column cursor used for paragraph fragments. Floating tables remain on the existing anchored table path in this PR.

- [ ] **Step 6: Run focused table and document layout tests**

```bash
pnpm vitest run packages/docx/src/table-fragments.test.ts packages/docx/src/table-split.test.ts packages/docx/src/document-layout.test.ts packages/docx/src/table-row-height.test.ts packages/docx/src/table-layout-reuse.test.ts
pnpm --filter @silurus/ooxml-docx typecheck
```

- [ ] **Step 7: Commit table fragment pagination**

```text
refactor(docx): paginate table rows as measured fragments
```

### Task 16: Paint table fragments and remove table stamps

**Files:**
- Modify: `packages/docx/src/fragment-paint.ts`
- Modify: `packages/docx/src/fragment-paint.test.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/types.ts`
- Modify: `rules/no-docx-measurement-in-fragment-paint.yml`
- Modify: `rule-tests/no-docx-measurement-in-fragment-paint-test.yml`

**Interfaces:**
- Table paint consumes stored column widths, row heights, cell fragments, and border jobs.
- Parsed model and fragments remain immutable during paint.
- `tableColWidthsPt`, `tableRowHeightsPt`, `tableLayoutInputs`, and migrated cell line stamps are removed.

- [ ] **Step 1: Add a failing table paint-purity test**

Paint a premeasured table fragment with a Canvas stub whose `measureText` throws. Assert cell text, backgrounds, borders, repeated headers, and vertical merge geometry render without measurement.

- [ ] **Step 2: Run and observe current table paint measurement failure**

```bash
pnpm vitest run packages/docx/src/fragment-paint.test.ts
```

- [ ] **Step 3: Paint stored table geometry**

Use fragment column widths and row heights directly. Render each cell's stored paragraph or nested-table fragments. Preserve existing border conflict resolution and vertical alignment using measured content height.

- [ ] **Step 4: Remove runtime table stamps and fallback recomputation**

Delete the migrated runtime fields from `PaginatedBodyElement` and remove paint-time table layout fallback from the production `DocumentLayout` path. Retain compatibility adapters only until their tests are migrated in the same commit.

- [ ] **Step 5: Extend static enforcement and run all focused checks**

```bash
pnpm vitest run packages/docx/src/fragment-paint.test.ts packages/docx/src/table-fragments.test.ts packages/docx/src/table-layout-reuse.test.ts packages/docx/src/cell-paragraph-line-reuse.test.ts packages/docx/src/table-split.test.ts
pnpm lint
pnpm lint:test
pnpm --filter @silurus/ooxml-docx typecheck
git diff --check
```

- [ ] **Step 6: Commit table fragment paint**

```text
refactor(docx): paint table fragments without remeasurement
```

### Task 17: Final regression verification and PR 6 merge

**Files:** all body/table layout files and tests from PRs 1-6.

- [ ] **Step 1: Run complete Rust and TypeScript verification**

```bash
cargo fmt --all --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
pnpm typecheck
pnpm test
pnpm lint
pnpm lint:test
git diff --check main...HEAD
```

Expected: every command exits 0.

- [ ] **Step 2: Build WASM and run visual verification**

```bash
pnpm build:wasm
pnpm --filter @silurus/ooxml-docx vrt
```

Review public and local-only documents without updating references automatically. Confirm the targeted document class uses the expected table line pitch and pagination, and representative unrelated DOCX layouts do not regress.

- [ ] **Step 3: Start Storybook for user verification**

```bash
pnpm storybook
```

Report the selected local port and keep the server running until the user finishes verification.

- [ ] **Step 4: Request a final read-only review**

Use Claude Fable 5 or Opus 4.8 to review the public diff for specification compliance, fragment ownership, hidden remeasurement, table split invariants, and missing tests. Resolve valid findings and rerun Steps 1-2.

- [ ] **Step 5: Push, open the final OSS-safe PR, and wait for CI**

The PR describes the shared measurement/fragment implementation and specification gaps. It contains no private fixture identity, content, local path, or personal environment detail.

- [ ] **Step 6: Merge through the PR**

```bash
gh pr merge --merge
```

Expected: all required checks pass and the merge commit lands on `main`; no direct push and no squash.
