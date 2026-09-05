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
