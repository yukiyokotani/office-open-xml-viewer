# Controlled PowerPoint ruler experiments

These source conditions investigate how PowerPoint serializes paragraph tabbing
into `TextRulerAtom` and `StyleTextPropAtom`. They do not encode a converter rule
or an expected Office result.

```sh
node --test scripts/legacy-ppt-ruler-probes.test.mjs
node scripts/legacy-ppt-ruler-probes.mjs
```

The second command prints a JSON manifest only. It does not generate a deck or
launch Office. Use the manifest to author a macro-free OOXML control presentation
with one slide per condition. Keep generated decks, manifests and all inspection
reports local; do not commit them.

## Source design

Each condition has two paragraphs containing the same literal tab-separated
text. The first always retains the baseline. The second changes one paragraph
property, except for the baseline and unchanged repeat. Keep font, text-box
position, width and text insets fixed. Use the same no-bullet paragraph format
for both paragraphs. Put the condition ID outside the text box under test.

The 29 conditions cover signed/zero default intervals, intervals around a binary
master-unit boundary, nonzero left margins, signed first-line offsets, custom
tab positions, all four tab alignments, an empty tab list, two tab stops, indent
level and paragraph direction. A master unit is 1/576 inch, or 1587.5 EMU. The
authored integer-EMU values deliberately straddle some half-EMU boundaries;
they are not assertions about PowerPoint's rounding rule.

This is not an exhaustive factorial design. It does not yet cover absent versus
explicit values, conflicting master defaults, all indent levels, RTL scripts,
or arbitrary combinations of these parameters. Do not generalize an observed
precedence rule beyond tested conditions. In particular, a source `a:tabLst`
does not prove that PowerPoint retained distinct per-paragraph binary tab lists.

## Required Office and converter sequence

1. Validate the authored OOXML, freeze its hash and record the Office version.
   Verify no macros, action links, external relationships or embedded active
   content exist. Do not relax the ordinary corpus export safety checks for
   arbitrary user files based on these synthetic inputs.
2. Open a disposable authored copy in local PowerPoint and save it as PPT.
   Close it, then reopen the saved PPT. Do not inspect only the in-memory state
   before binary serialization. Record any repair or compatibility messages.
3. Export the reopened PPT directly to PDF. Separately save a PPTX copy for the
   binary-to-OOXML mapping evidence. Preserve the authored source unchanged.
4. Resolve the binary current edit and live persist references. Record each
   shape's text ruler and each paragraph's direct and inherited properties.
   Include default intervals, tab arrays, indent level and coordinate origin.
   Distinguish no tab array from an explicitly empty array. Physical record
   order alone is not proof that an object is live.
5. Determine which source conditions survived as distinct binary conditions.
   If PowerPoint merges two source cases, they cannot establish the behavior
   of the distinction that was lost. Any binary counterfactual must preserve
   container/record validity and be labeled separately from Office-authored
   binary output.
6. Compare the direct-binary PDF, Office-reconverted OOXML, converter output and
   previous converter output. Check text alignment and tab placement as well as
   package structure. Keep Office-fidelity evaluation separate from renderer
   self-VRT. A structural pass is not visual-fidelity approval.

The relevant definitions are MS-PPT 2.2.29, 2.9.20, 2.9.23-24, 2.9.28-30,
2.9.41 and 2.13.32, and ECMA-376 21.1.2.2.7 and 21.1.2.2.13-14. A tab offset's
binary origin depends on whether it belongs to a ruler or a paragraph
exception. Do not add a margin correction or choose conflicting property
precedence solely to improve a private sample.

### Container constraints

Read the constraints of the containing record as well as `TextPFException`.
MS-PPT 2.9.45 requires `leftMargin`, `indent`, `defaultTabSize` and `tabStops`
to be absent in a `TextPFRun` inside `StyleTextPropAtom`. A synthetic direct
paragraph run containing these fields is not a conforming serialization oracle,
even if the converter's shared exception reader accepts it. The master-level
exception in 2.9.35-36 does not have that restriction; placeholder inheritance
from the main master is specified in 2.9.44.

Inspect ruler records and master-level exceptions separately. An observation of
zero direct-run tab arrays is not evidence that the document has no custom tabs.
If Office produces an exception to a container constraint, record that separately
as observed Office behavior before using it to justify compatibility code.

## Verification status

The first paired source presentation has completed a local PowerPoint roundtrip:
save as PPT, close, reopen the saved PPT, export a local print-quality PDF, and
save a separate OOXML copy. All 29 PDF pages were inspected. This establishes
the reference route, not converter fidelity.

In the initial same-text-body arrangement, varying the second paragraph's tab
position, alignment, empty/list state or default interval did not produce a
distinct PDF placement from the baseline in the tested cases. For selected tab
cases, the reconverted OOXML retains the baseline stop in the text body's
`a:lstStyle`, not the direct `a:pPr`. Do not mistake an absent direct tab list
for an absent effective tab list, or infer that the converter should ignore
those properties. The observation does not isolate OOXML import from PPT
serialization as the point where the distinction was lost.

Before inferring the behavior of these lost distinctions, repeat the affected
conditions with control and treatment in separate text bodies, retaining the
original same-body experiment and its reference artifacts. Verify which values
actually survive in the binary. No ruler precedence rule or new production
converter behavior follows from the first roundtrip alone.
