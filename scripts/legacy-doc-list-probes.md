# Controlled Word list indentation experiments

These probes investigate list paragraph formatting during DOCX-to-DOC conversion
and native DOC display. They contain synthetic text only. They do not encode an
expected precedence rule, and generating them does not establish Office behavior.

## Generate and verify the inputs

Use a Python environment with `python-docx` and `lxml` installed:

```sh
python scripts/legacy-doc-list-probes.test.py
python scripts/legacy-doc-list-probes.py /tmp/doc-list-probes-new
```

The output directory must not already exist. The generator creates one DOCX
with 64 controlled cases and a separate JSON manifest. The intended layout is
one case per page, with two numbered paragraphs per case. Verify actual page
counts after rendering; the intent is not an observed result. Each case uses
a distinct abstract numbering definition and instance to isolate list counters.

The manifest records parameters in twips, parent cases, changed parameters and
the DOCX hash. Except for two baseline cases and two unchanged repeats, each
case differs from its parent in exactly one parameter. The second half repeats
the experiment in RTL paragraph layout. Latin probe text deliberately isolates
paragraph geometry from complex-script shaping; it does not cover Arabic text
shaping or mixed-direction runs.

The first phase covers absent versus explicit-zero paragraph indents, signed
left/right indents, first-line versus hanging indentation, conflicting style
properties, style-inherited numbering, a level's paragraph-style association,
list-level bidi settings, and tab suffix/stop/clear behavior. It is not an
exhaustive factorial design. Expand interactions and boundary cases based on
observed results rather than fitting a private sample.

The optional second phase uses a nonzero list-level right indent so direct
right indents actually conflict with a retained list value. It also examines
one-twip boundaries around zero and matching list indents, cumulative left/right/
first-line overrides, and conflicting list/paragraph bidi in both directions:

```sh
python scripts/legacy-doc-list-probes.py /tmp/doc-list-probes-interactions --phase interactions
```

This creates another 64 cases, identified by `Q` rather than `P`. Its manifest
records the experiment phase. The default `P` cases and authored DOCX remain
unchanged. One-twip differences may be below PDF text-coordinate precision;
inspect binary values and quantify export rounding instead of treating a
visually indistinguishable pair as proof of equivalent formatting semantics.

The third phase varies the paragraph styles of two paragraphs sharing one list:
the same custom style, Normal, custom followed by Normal, or two distinct custom
styles with identical properties. Each pattern is exercised with absent direct
indentation, direct-left zero/1440, direct-first-line zero, and combined zeros,
in both directions, with unchanged repeats:

```sh
python scripts/legacy-doc-list-probes.py /tmp/doc-list-probes-style-association --phase style-association
```

This creates 48 `R` cases without changing the authored `P` or `Q` files. It
attempts to obtain unlinked binary controls; the source style pattern alone
does not establish whether Word will create or retain a list-style association.

## Required local Office sequence

1. Record Word version/build, platform, installed/substituted fonts and relevant
   compatibility settings. Preserve the authored DOCX and manifest unchanged.
2. Open a disposable copy in local Word. Save it as Word 97-2003 DOC. Record any
   compatibility warnings; do not silently approve an unexpected operation.
3. Close that saved document. Reopen the saved DOC, not the original DOCX or an
   unsaved in-memory conversion. Do not execute macros or update external links.
4. Export the reopened DOC directly to PDF. This is the binary display oracle.
5. Reopen the same saved DOC separately and save an OOXML copy. This is evidence
   for the binary-to-XML mapping, not the display oracle. An original-DOCX PDF
   can additionally distinguish serialization changes from display changes.
6. Hash every artifact. Inspect the saved DOC's effective PAPX/PCD properties,
   signed iLfo, iLvl, selected LFO/LSTF/LVL, linked styles, and physical/logical
   indentation SPRMs. Record both authored and observed values. If Word merges
   two source conditions into the same binary condition, that pair does not
   establish the effect of those parameters on DOC rendering.
7. Compare marker, first-line, continuation-line and wrapped-text coordinates
   in the direct-DOC PDF, generated converter output, and previous converter
   output. Inspect page images as well as extracted geometry. Keep source-DOCX
   display, Word-converted DOCX display, and binary display results distinct.

Only infer a rule after the affected condition, controls, unchanged repeats,
boundary cases and counterexamples have been examined. Distinguish normative
MS-DOC/ECMA rules from a bounded observed Office behavior. Do not alter renderer
thresholds or add filename checks to improve a probe score.

Relevant references are MS-DOC 2.4.6.3, 2.4.6.6 and 2.6.2, and ECMA-376
17.3.1.12, 17.3.1.38, 17.7.2 and 17.9. In particular, source XML alone cannot
prove which binary list formatting properties Word retained.

Generated DOCX/DOC/PDF files, local manifests and visual comparisons stay outside
version control. Existing private references must not be replaced automatically.
This script changes no shipped converter, parser, renderer or opt-in contract.

## Observed scope and unresolved precedence

Both phases have been exercised with local Word 16.112.3 on macOS. Each saved
DOC was closed, reopened, and exported directly to a 64-page PDF. Separate
Word-reconverted DOCX files were retained for mapping comparisons. These are
observations of that Word version, not a normative amendment to MS-DOC.

The 256 numbered paragraphs across the two phases all have positive `iLfo`,
empty piece-property records, and a Word-generated paragraph-style association
in the selected binary list. Consequently, an input DOCX without `w:pStyle`
inside its list level did **not** produce an unlinked binary control. Do not
claim coverage of unlinked lists, negative `iLfo`, or nonempty PCD from these
experiments.

The saved DOC retains conflicting direct indentation properties, including
explicit zero, but Word removes some direct properties equal to the selected
list value. In the interaction phase, the list retains logical left/right
indentation of 720 twips and first-line indentation of -360 twips.

For the LTR continuation line, the observed PDF x positions are:

| Direct left indent (twips) | Direct-DOC PDF x (points) |
| ---: | ---: |
| -1 | 71.949936 |
| 0 | 72.000000 |
| 1 | 72.050064 |
| 719 | 107.949936 |
| 720 | 108.000000 |
| 721 | 108.050064 |

The converter under investigation places all these continuations at 108 points.
Word-reconverted OOXML rendered by the same unchanged diagnostic renderer
preserves the corresponding direct-left positions. First-line boundary cases
also move the marker while leaving the continuation start unchanged. This
rules out an explicit-zero-only explanation for the tested conflict.

This is not evidence for preserving every property present before list
formatting: the baseline style-indentation controls behave differently from
direct PAPX conflicts. Nor does the left edge of an LTR continuation establish
the effective right indent. Right-indent changes require line-wrap/right-edge
analysis; RTL starts also depend on text widths and wrapping.

Unchanged repeats have small differences in extracted interior word boxes
(observed examples include 0.000672 and 0.013752 points), even when the selected
line-start anchors agree. No fitted acceptance tolerance or pixel-equivalence
claim follows from these observations. Occupied-region contact sheets were
reviewed, and both converter-output and Word-OOXML diagnostic renders produced
64 pages per phase. These Node/Skia runs are not browser self-VRT approval.

MS-DOC 2.4.6.6 part 2 and 2.4.6.3 part 3 describe applying list paragraph
properties after direct properties. Section 2.6.2 explicitly protects logical
left and first-line indentation for negative `iLfo`; it does not explain the
positive-reference observations above. `LVLF.fIndentSav` describes removing
indentation when numbering is removed, not a general direct-format override.
Do not present an inferred positive-reference precedence rule as mandated by
these sections. A compatibility implementation needs an explicitly reviewed
scope, counterexamples, and approval; no such override is introduced here.

### Additional style-association controls

The third phase was saved and reopened in the same local Word version, then
exported to a 48-page direct-DOC PDF. All 96 numbered paragraphs were mapped
through their physical paragraph marks, PAPX, positive `iLfo`, selected LFO and
LSTF. They have level zero, no LFO formatting overrides, and empty piece
properties. Of these paragraphs, 84 use `rgistdPara[0] = 0x0fff` (unlinked),
and 12 use a linked style. All selected lists are single-level lists.

The unlinked LTR alternating-style controls retain list-left 720 and first-line
-360 twips in both physical and logical LVL properties. Conflicting direct
PAPX values remain present. Their measured direct-DOC PDF positions are:

| Direct left / first-line (twips) | Marker x (points) | Continuation x (points) |
| --- | ---: | ---: |
| absent / absent | 90 | 108 |
| 0 / absent | 54 | 72 |
| 1440 / absent | 126 | 144 |
| absent / 0 | 108 | 108 |
| 0 / 0 | 72 | 72 |

Both paragraphs in each case exhibit these anchors. Word-reconverted OOXML
also retains the conflicting direct indents. Thus the observed direct-indent
effect is not limited to linked lists. This does not establish a general
precedence rule for other properties, levels, overrides, or piece properties.

Source style patterns did not uniquely determine the saved association: an
unchanged custom-style baseline repeat became unlinked although the first
baseline was linked, with the same selected PDF anchors. Consequently the
generator does not promise a particular association from a particular pattern,
and this run does not establish Word's association-creation algorithm. Inspect
the saved binary on every run. All 48 occupied page regions were reviewed as
contact sheets; this is not a full-resolution visual equivalence or browser
self-VRT claim. No converter or renderer compatibility rule was changed.
