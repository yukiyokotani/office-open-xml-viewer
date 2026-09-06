# Legacy DOC counter instance probes

This generator creates nine passive DOCX controls for one unresolved question:
when two `w:num` instances reference the same `w:abstractNum`, does native Word
advance one counter or a counter per instance? ECMA-376 defines the inheritance
of numbering properties but does not explicitly assign mutable counter identity.
The generator therefore records every outcome as `UNOBSERVED`.

## Generate the source controls

Run the script with the bundled document Python runtime and a new output
directory:

```bash
python scripts/legacy-doc-counter-probes.py /new/output/directory
```

Existing output paths are rejected. The result contains a deterministic
`counter-instance-probes.docx` and a manifest with the source hash and exact
settings. It contains no macros, fields, embedded objects, or external
relationships.

## Cases

Each case starts on a separate page, owns fresh numbering IDs, and contains four
numbered paragraphs. `1` and `2` below denote the two `w:num` instances, not
literal marker text.

| Case | Instance sequence | Second instance |
| --- | --- | --- |
| C01 | 1111 | unused baseline instance |
| C02 | 1212 | same abstract, no override |
| C03 | 1122 | same abstract, no override |
| C04 | 1212 | identical full level replacement |
| C05 | 1212 | full level replacement with bold marker only |
| C06 | 1212 | `startOverride=7` |
| C07 | 1212 | full level replacement whose `start=7`, without `startOverride` |
| C08 | 1212 | separate but equivalent abstract definitions |
| C09 | 1212 | unchanged repeat of C02 with fresh IDs |

Markers remain dynamic through `w:lvlText w:val="%1."`; paragraph text does not
contain expected numbers.

## Native Office protocol

1. Export a PDF directly from the untouched source DOCX. This is the primary
   oracle for how Word interprets the emitted OOXML.
2. Save the source as legacy DOC, close it, and reopen that DOC.
3. Export a PDF directly from the reopened DOC.
4. Save a roundtrip DOCX and inspect its numbering XML as well as its PDF.
5. Record each stage separately. A DOC down-save or roundtrip change must not be
   attributed to the source OOXML behavior.

For each case, record all four displayed markers exactly. Do not classify the
counter as shared or independent from the source structure, the generator's
labels, or another case. C02 is the minimal alternating alias control, C03 checks
ordering, C08 is the independent-definition control, and C09 detects order- or
save-dependent drift.

This protocol requires native Word. LibreOffice is not a substitute for the
Office observation and is not part of this probe.

## Tests

```bash
python scripts/legacy-doc-counter-probes.test.py
```

The tests inspect the package in memory. They verify input passivity,
deterministic bytes, numbering schema order, ID/reference isolation, exact case
differences, title styling, dynamic markers, and rejection of an existing output
directory. They do not create the probe artifact or assert a Word counter result.
