# `@maxgent/ooxml-pptx-editor`

Optimistic PPTX editing for Maxgent's `@maxgent/ooxml` fork: mutate an
in-memory `Presentation`, translate commands to OfficeCLI batches, and paint
through a `PptxEditorViewerHost` backed by a loaded `PptxPresentation`.

This package is published **separately** from `@maxgent/ooxml` so upstream
viewer syncs stay thin. Prefer the high-level `PptxEditorSession` +
`PptxEditorViewBinding` surface unless you are extending the editor itself.

| Package | Role |
| --- | --- |
| `@maxgent/ooxml` | Viewer / parser SDK (peer dependency) |
| `@maxgent/ooxml-pptx-editor` | Editor session, mutations, view binding |

## Architecture

```text
UI / host app
  │  submit(command) / undo() / redo()
  ▼
PptxEditorSession          ← snapshot, history, sync state, listeners
  ├─ PptxEditorStore       ← optimistic Presentation + pending commands
  ├─ UndoRedoStack         ← invert + stack; issues undo/redo command ids
  └─ SerialOfficeCliSubmitter
        │  sendBatch(OfficeCliBatch)
        ▼
     your transport        ← confirmed | rejected | unknown

PptxEditorViewBinding
  │  session.subscribe → coalesced apply
  ▼
PptxEditorViewerHost       ← PptxViewer + PptxPresentation (mode: 'main')
  └─ replaceSlides + redraw
```

Data ownership:

| Layer | Owns |
| --- | --- |
| Session | Optimistic `Presentation`, undo/redo, submission queue, sync halt |
| View host | Canvas, package media/theme plumbing, paint |
| Transport | Persistence / OfficeCLI side effects |

The binding never owns the canvas. It only pushes the session’s presentation
into a host that already loaded the same package.

## Install

```bash
pnpm add @maxgent/ooxml @maxgent/ooxml-pptx-editor
```

```ts
import { PptxPresentation, PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  PptxEditorViewerHost,
  UpdateTextMutation,
  createElementRef,
  OFFICECLI_BATCH_SEND_STATUSES,
  COMMAND_SUBMISSION_STATUSES,
} from '@maxgent/ooxml-pptx-editor';
```

Requires `@maxgent/ooxml >= 0.77.0-0`, which provides the internal main-thread
slide replacement hook. Inside this monorepo, depend on the workspace package:

```json
{
  "dependencies": {
    "@maxgent/ooxml-pptx-editor": "workspace:*"
  }
}
```

## Quick start

Minimal loop: load a viewer in **main** mode, export editor JSON with
`toEditorPresentation()`, open a session on that model, bind them, then submit
commands.

```ts
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  PptxEditorViewerHost,
  UpdateTextMutation,
  createElementRef,
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
  type OfficeCliBatch,
  type OfficeCliBatchSendResult,
} from '@maxgent/ooxml-pptx-editor';
import { PptxPresentation, PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';

async function openEditor(args: {
  canvas: HTMLCanvasElement;
  source: string | ArrayBuffer;
  sendBatch: (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>;
}) {
  // The viewer borrows this main-mode presentation. The caller owns it.
  const loadedPresentation = await PptxPresentation.load(args.source, {
    mode: 'main',
  });
  // Detached editor JSON from the same loaded package (not a second parser source).
  const presentation = await loadedPresentation.toEditorPresentation();
  const viewer = PptxViewer.fromPresentation(args.canvas, loadedPresentation);

  let commandSeq = 0;
  const session = new PptxEditorSession({
    presentation,
    sendBatch: args.sendBatch,
    createCommandId: ({ direction, sourceCommandId }) => {
      commandSeq += 1;
      return `${direction}:${sourceCommandId}:${commandSeq}`;
    },
  });

  const host = new PptxEditorViewerHost(viewer, loadedPresentation);
  const binding = new PptxEditorViewBinding({
    session,
    host,
    onRenderError: (cause) => {
      console.error('view apply failed', cause);
      // Host may be stale; recover with a full sync:
      binding.requestRender();
    },
  });
  await binding.whenIdle();

  return { viewer, loadedPresentation, session, binding };
}

async function editFirstShapeText(
  session: PptxEditorSession,
  presentation: Presentation,
  nextText: string,
) {
  const slide = presentation.slides[0] as Presentation['slides'][number];
  const element = slide.elements[0] as Presentation['slides'][number]['elements'][number];
  const target = createElementRef(slide, element, 0);

  const submission = session.submit({
    id: 'edit-text-1',
    mutations: [new UpdateTextMutation({ target, value: nextText })],
  });

  // Optimistic model is already updated.
  const snapshot = session.getSnapshot();
  console.log(snapshot.presentation.slides[0].elements[0]);

  const result = await submission.settled;
  if (result.status !== COMMAND_SUBMISSION_STATUSES.CONFIRMED) {
    throw new Error(`edit did not confirm: ${result.status}`);
  }
}

async function sendBatch(batch: OfficeCliBatch): Promise<OfficeCliBatchSendResult> {
  // Call your OfficeCLI / backend. Return one of:
  //   { status: 'confirmed' }
  //   { status: 'rejected', cause }
  //   { status: 'unknown', cause }  → session sync halts until resync()
  void batch;
  return { status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED };
}
```

Teardown:

```ts
binding.dispose();
session.dispose();
viewer.destroy();
```

Dispose the binding before (or with) the session so an in-flight drain does not
call `getSnapshot()` on a disposed session.

## Presentation model requirements

The session operates on `@maxgent/ooxml/pptx` `Presentation` JSON, not on the
zip package itself. Bootstrap that model from the loaded package:

```ts
const loadedPresentation = await PptxPresentation.load(source, { mode: 'main' });
const presentation = await loadedPresentation.toEditorPresentation();
```

Editable slides must expose complete `elementSources` parallel to `elements`
(same length). Text, transform, and style mutations support direct slide shapes
(`origin: 'slide'`). `RemoveElementMutation` also maps slide-origin pictures,
tables, and charts to OfficeCLI paths from their frontend element type.
OfficeCLI `zorder` is derived as the ordinal among
slide-origin entries before `presentationElementIndex` (1-based). This matches
true spTree position for top-level 1:1 shapes; groups / hidden nodes that expand
or skip break that equivalence. Master/layout decorations are
visible in the model but reject edit attempts with
`element.unsupportedOrigin`.

Stable identity:

| Concept | Id source |
| --- | --- |
| Slide | `slide.partName` when present, else `String(slide.index)` via `getSlideMutationId` |
| Element | OOXML `cNvPr` id when present, else `index:<n>` via `getElementMutationId` |

Build refs with `createElementRef(slide, element, elementIndex)` rather than
hand-writing ids.

```ts
import { createElementRef, ELEMENT_ORIGINS } from '@maxgent/ooxml-pptx-editor';

const target = createElementRef(slide, element, elementIndex);
// target.origin === ELEMENT_ORIGINS.SLIDE for editable slide elements
```

## Commands and mutations

A **command** is the atomic unit of optimistic update, history, and transport:

```ts
import type { Command } from '@maxgent/ooxml-pptx-editor';

const command: Command = {
  id: 'cmd-1',                 // unique per submission
  mutations: [/* at least one */],
  mergeKey: 'title-typing',    // optional history coalescing key
};
```

Built-in mutations:

| Class | Effect | OfficeCLI |
| --- | --- | --- |
| `UpdateTextMutation` | Replace shape plain text, whole-shape styles, or incremental paragraph/span edits (`text` and/or `style`) | `set` path + `{ text, bold, … }` and/or `range=` |
| `UpdateShapeMutation` | Patch shape position, size, rotation, flips, fill, or outline | `set` path + changed shape props |
| `InsertSlideMutation` | Insert an empty slide at a 0-based index | `add` under `/` with `type: 'slide'` and `index` |
| `RemoveSlideMutation` | Remove a slide; direct removal is not undoable | `remove` at the current slide path |
| `AddElementMutation` | Insert a slide element at indexes | `add` under slide path |
| `RemoveElementMutation` | Remove a slide-origin shape, picture, table, or chart; only shape removal is undoable | type-based stable `remove` path |

Low-level apply without a session:

```ts
import { applyCommand, applyMutation } from '@maxgent/ooxml-pptx-editor';

const { presentation, changedSlideIds, changedElements } = applyCommand(
  currentPresentation,
  command,
);
```

## Session API

```ts
const session = new PptxEditorSession({
  presentation,
  sendBatch,
  createCommandId,
  onListenerError, // optional; defaults to console.error
});

session.getSnapshot();
session.subscribe((change) => { /* UI / telemetry */ });
session.submit(command);
session.undo();
session.redo();
session.resync(authoritativePresentation);
session.dispose();
```

### Snapshot

```ts
interface PptxEditorSessionSnapshot {
  presentation: Presentation;       // optimistic current model
  syncState: EditorSyncState;       // ready | halted
  pendingCommandIds: readonly string[];
  isSubmitting: boolean;
  undoDepth: number;
  redoDepth: number;
  canUndo: boolean;
  canRedo: boolean;
}
```

### Change events

`subscribe` receives a `PptxEditorSessionChange` after store or history updates.
Useful fields:

- `reason` — `command.dispatched` / `command.confirmed` / `command.rejected` /
  `submission.halted` / `presentation.resynced` / `history.changed`
- `snapshot` — post-change session snapshot
- `commandId`, `invalidatedCommandIds`
- `changedSlideIds`, `changedElements` — for incremental UI (the view binding
  already consumes these)

Dispatch is optimistic: `submit` / `undo` / `redo` update the local presentation
before transport settles. `submission.settled` resolves with the final
submission status.

### Undo / redo

```ts
if (session.getSnapshot().canUndo) {
  await session.undo().settled;
}
if (session.getSnapshot().canRedo) {
  await session.redo().settled;
}
```

`createCommandId` must mint a **new** id for every undo/redo submission:

```ts
createCommandId: ({ direction, sourceCommandId }) =>
  `${direction}:${sourceCommandId}:${crypto.randomUUID()}`,
```

Undo and redo history advances optimistically with the local presentation.
Pending invertible commands can be undone or redone immediately; the resulting
commands remain serial in the OfficeCLI submission queue. A pending
non-invertible command temporarily disables both operations. If a command is
rejected, it and its invalidated optimistic tail are removed from history.

### Submission outcomes

Your `sendBatch` must return one of:

| Status | Meaning | Session effect |
| --- | --- | --- |
| `confirmed` | Server accepted the batch | Command leaves pending; optimistic history becomes confirmed |
| `rejected` | Server rejected with known cause | Optimistic change rolled back for that command |
| `unknown` | Outcome unclear (timeout, network ambiguity) | Sync **halts**; further submits blocked until `resync` |

Settled `CommandSubmissionResult` statuses:

| Status | Meaning |
| --- | --- |
| `confirmed` | Applied and acknowledged |
| `rejected` | Rolled back |
| `invalidated` | Dropped because an earlier command in the serial queue failed |
| `halted` | Queue stopped after an `unknown` send |

### Halt and resync

When transport returns `unknown`, the session enters
`syncState.status === 'halted'`. Do not keep submitting. Fetch an authoritative
presentation and reset:

```ts
import { EDITOR_SYNC_STATUSES } from '@maxgent/ooxml-pptx-editor';

const { syncState } = session.getSnapshot();
if (syncState.status === EDITOR_SYNC_STATUSES.HALTED) {
  const authoritative = await fetchAuthoritativePresentation();
  session.resync(authoritative);
  // History and pending commands are cleared; sync returns to ready.
}
```

If `resync` changes the slide count, the standard view binding replaces the
in-memory slide list without reloading the PPTX package.

## View binding

`PptxEditorViewBinding` connects a session to any host that implements:

```ts
interface PptxEditorViewHost {
  applyPresentation(
    presentation: Presentation,
    options?: { readonly changedSlideIndexes?: readonly number[] },
  ): void | Promise<void>;
}
```

Create the standard host from a viewer that borrows the same main-mode
presentation:

```ts
const host = new PptxEditorViewerHost(viewer, loadedPresentation);
const binding = new PptxEditorViewBinding({
  session,
  host,
  syncOnBind: true, // default: push current session state immediately
  onRenderError: (cause) => {
    console.error(cause);
    binding.requestRender();
  },
});

await binding.whenIdle();
binding.requestRender(); // force a full apply
binding.dispose();
```

Behavior:

- Subscribes to slide changes and slide-count changes.
- Coalesces rapid mutations: while one apply is in flight, later changes merge
  into the next revision (latest snapshot wins).
- Passes `changedSlideIndexes` for incremental patches; uses a full apply when
  `requestRender()` is called or after a failed apply (host state unknown).
- Isolates host failures: errors go to `onRenderError`, the binding stays usable.
- Does **not** auto-retry. After a failure, call `requestRender()` or wait for
  the next session change (which escalates to a full apply).

`PptxEditorViewerHost` behavior:

- Keeps the loaded package’s media / theme plumbing; only swaps in-memory slide
  JSON used by the next paint.
- Invalidates find geometry; clears leftover highlight overlays even when the
  visible slide is not redrawn.
- Replaces the complete in-memory slide list when the slide count changes.
- Throws if the presentation is in `mode: 'worker'`.
- Does not own resources. Dispose the binding, viewer, and presentation
  explicitly.

## Element selection

`PptxEditorSelectionController` maps canvas pointer coordinates into slide EMUs,
hit-tests slide-origin elements from front to back, and exposes the selected
`ElementRef` used by mutations:

```ts
const selection = new PptxEditorSelectionController({
  session,
  host: viewer,
});

selection.subscribe(({ snapshot }) => {
  const selected = snapshot.selection;
  if (!selected) return hideSelectionOverlay();
  showSelectionOverlay(selected.element, selected.target);
});
```

Selection is transient UI state: it is not part of `Presentation`, command
history, or OfficeCLI transport. The controller follows optimistic element
updates, clears itself when the selected element disappears, and reconciles the
selection against the host's current slide before snapshot reads and pointer
input. Pass the viewer itself so `slideIndex` remains a live getter; do not copy
`viewer.slideIndex` into a plain host object, because that freezes its initial
numeric value. Dispose the controller before the session:

```ts
selection.dispose();
binding.dispose();
viewer.destroy();
loadedPresentation.destroy();
session.dispose();
```

The MVP hit test uses each rotated element frame, reverse render order, and a
4-CSS-pixel tolerance for lines. It skips layout/master decorations. Media is
not selectable and blocks selection of elements underneath it. Other elements
without a stable numeric OOXML id can be selected for UI feedback but report
`isOfficeCliTargetable: false`. In
edit mode, keep viewer text selection and hyperlink overlays disabled unless
the app forwards their pointer events into `selectAtClientPoint`.

## OfficeCLI transport

Mutations translate to an `OfficeCliBatch` via `toOfficeCliBatch` /
per-mutation `toOfficeCli`. The product envelope looks like:

```ts
{
  schemaVersion: /* OFFICECLI_BATCH_SCHEMA_VERSION */,
  officecliVersion: /* OFFICECLI_VERSION */,
  commandId: 'edit-text-1',
  commands: [
    { command: 'set', path: '/slide[1]/shape[@id=7]', props: { text: 'Hello' } },
  ],
}
```

Paths use stable slide ordinals and OOXML `cNvPr` ids. Shape updates target
direct slide shapes. Element removal selects `/shape`, `/picture`, `/table`, or
`/chart` from the frontend element type.

You can translate without submitting:

```ts
import { toOfficeCliBatch } from '@maxgent/ooxml-pptx-editor';

const batch = toOfficeCliBatch(presentation, command);
```

## Lower-level building blocks

Most apps should stay on `PptxEditorSession`. These are exported for tests and
custom pipelines:

| API | Role |
| --- | --- |
| `PptxEditorStore` | Optimistic presentation + pending commands + sync state |
| `UndoRedoStack` | Invert commands and drive undo/redo submissions |
| `SerialOfficeCliSubmitter` | Serial queue over `sendBatch` |
| `applyCommand` / `applyMutation` | Pure local apply |

## Current limitations

Document these in product code rather than papering over them:

1. **Main-thread presentation only.** The internal slide replacement hook throws
   in `mode: 'worker'`. Load `PptxPresentation` with `{ mode: 'main' }`.
2. **Slide content only.** The host installs slide models; presentation theme /
   size fields on the session snapshot are not pushed into the viewer.
3. **Slide-origin elements only.** Master/layout elements are not editable.
   Shape text, transform, and style mutations remain shape-only. Removal also
   maps pictures, tables, and charts from their frontend type. Media is rejected
   at translate time (`target.unsupportedElement`) because OfficeCLI has no
   stable `@id` selector for video/audio. A grouped,
   wrapped, or projected element may produce a path that OfficeCLI rejects; a
   rejected submission rolls the optimistic deletion back.
   Removal of a non-shape element is not undoable because its binary parts and
   relationships cannot be restored from the projected presentation model.
   A confirmed non-undoable deletion clears command history.
   OfficeCLI `zorder` is derived from `origin: 'slide'` ordinals before
   `presentationElementIndex`; this matches spTree position for top-level 1:1
   shapes, not for groups / hidden nodes that expand or skip.
   `UpdateTextMutation` can patch whole-shape text (`value` + `style`) and
   incremental `edits` (paragraph `text` and/or `style`, or span `style`).
   Paragraph text replacement and span edits require separate mutations.
   Selection-scoped text replacement (span rewrite) remains out of scope.
   Clear-to-inherit (`null`) style keys are
   resolve-then-set for OfficeCLI using paragraph/body/presentation defaults
   (explicit values, not true OOXML attribute removal).
   Character offsets use run-concatenated plain text (OfficeCLI `range` rules).
4. **Complete `elementSources` required** for any editable slide.
5. **Bootstrap via `toEditorPresentation()`.** Prefer
   `await loadedPresentation.toEditorPresentation()` (`mode: 'main'` only) so
   the session JSON comes from the same loaded package the viewer paints.
   Passing a separately parsed `Presentation` remains possible for tests, but
   product hosts should not maintain a second bootstrap source.

## Publishing / fork sync

- Published as `@maxgent/ooxml-pptx-editor` (independent semver, currently
  `0.1.0`).
- Peer-depends on `@maxgent/ooxml` — not wired into the umbrella `exports` map,
  so upstream sync of the viewer SDK does not fight editor packaging.
- Keep the Maxgent-only internal `replaceSlides` hook as the only editor patch in
  `packages/pptx`; viewer composition remains inside `packages/pptx-editor`.

```bash
pnpm --filter @maxgent/ooxml-pptx-editor build
pnpm --filter @maxgent/ooxml-pptx-editor publish --access public
```

## Testing

```bash
pnpm --filter @maxgent/ooxml-pptx-editor test
pnpm --filter @maxgent/ooxml-pptx-editor typecheck
```

Focused suites live under `test/` (`session/`, `rendering/`, `history/`,
`submission/`, `transport/`) and exercise optimistic dispatch, halt/resync,
view coalescing, and OfficeCLI translation.
