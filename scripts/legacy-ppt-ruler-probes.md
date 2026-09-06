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

## Verification status

The source-condition generator is tested. The paired source presentation has
been authored and structurally checked locally, but its controlled local Office
roundtrip has not completed. No new compatibility precedence rule, production
converter behavior or visual-fidelity claim follows from this experiment yet.
