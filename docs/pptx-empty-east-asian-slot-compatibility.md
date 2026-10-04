# PPTX empty East Asian font slot compatibility

A run whose East Asian (`a:ea`) font slot is empty now draws its East
Asian-slot characters the way PowerPoint does. Fonts missing from the reference
catalogue use a line-local fallback in the ordinary baseline model; the
structural whole-body guards below still apply. No migration is required.

## Specification boundary

ECMA-376 Part 1 §21.1.2.3 assigns characters to the Latin, East Asian, complex
script and symbol slots. §21.1.2.3.1/.3, §20.1.4.1.16–.18/.24/.25 and
§20.1.10.81 define the slot fonts, theme tokens and empty typefaces.
§21.1.2.2.5/.11 make line size and spacing line-local.

Two things below are **observed PowerPoint compatibility**, not normative
rules:

- which face draws a glyph when the East Asian slot is empty;
- where the baseline splits inside a line.

## Office evidence

PowerPoint for Mac 16.113 opened synthetic control decks and exported tagged
PDFs with "best for electronic distribution". The controls held each factor
fixed while varying one at a time:

- 64 control/test pairs over four empty-slot decks;
- 29 fallback-scope pairs.

Faces were read from the embedded fonts' name tables, not from PDF font labels.
Baselines were read from the text state. The decks and PDFs are local
verification artifacts, not redistributable baselines.

- **Selected face.** With an empty, omitted or theme-empty `a:ea`, the selected
  face S is:
  - the run's complex-script (`a:cs`) face, when one is named (directly or
    through a theme token);
  - otherwise the run's Latin face.

  S draws every East Asian-slot glyph it covers, including Han and kana when it
  maps them. Language (en/ja/zh/ko), the theme Latin face and the Latin face's
  own coverage play no part when a cs face is named.
- **Symbol fallback.** A non-CJK glyph that S lacks is drawn in Calibri (§),
  even when no Calibri appears in the deck, or in Cambria Math (◆ ■).
- **CJK fallback.** For a CJK glyph that S lacks:
  - A Far-East face that maps basic CJK uses its own script chain, whichever
    slot it fills. A Japanese face sends simplified Han to Microsoft JhengHei
    and Hangul to Malgun Gothic.
  - A face that maps no CJK Unified Ideograph falls back by its PANOSE serif
    style. Styles 1–10 go to MS Mincho; 11–15 and unknown faces go to MS Gothic.
  - Whether a face maps basic CJK comes from its installed cmap, recorded in the
    reference catalogue.
- **Synthetic styles.** A family without an italic face is drawn upright with a
  synthetic slant. Its vertical metrics are those of the upright resource.
- **Line metrics.** A line is sized by the OS/2 metrics of the resources that
  draw it: usWinAscent / usWinDescent, or the typo metrics plus line gap under
  USE_TYPO_METRICS. A face changes only its own line; earlier and later lines
  keep their baselines.

## Owner decisions (accepted platform differences)

- **(B) Second CJK fallback.** After the first fallback of a Far-East face that
  maps no basic CJK, the next face depends on the font itself. SimSun-ExtB and
  MingLiU-ExtB differ, and no general font property explains the difference.
  It is left to the browser's own glyph fallback. No face name is encoded.
- **(c) Unknown drawing face.** The browser does not report which installed
  face draws a fallback glyph.
  - Such a glyph adds no face to its line, and the line keeps the metric model
    of its known faces. Only a line with no known face at all uses the ordinary
    model.
  - Synthetic italic uses the browser's own oblique, not PowerPoint's 0.3333
    shear.
  - Symbol coverage is recorded from each concrete cut's Unicode cmap in a
    bounded domain derived from the slot sweeps: Latin-1 symbols, General
    Punctuation, Letterlike and Number Forms, Arrows, Math, Misc Technical,
    Enclosed, Box/Block/Geometric Shapes, Misc Symbols and Dingbats. Identical
    repertoires share inclusive ranges. Outside this domain the previous
    selected-face model remains.
  - Within the domain the first covering face in the painting stack owns the
    metric, including Calibri, Cambria Math and the CJK tier. Unknown coverage
    stops attribution: an earlier unknown face might draw the glyph.
  - Office repertoire profiles take precedence over same-name system cmaps,
    following the existing CJK coverage policy. Real and synthetic cuts use
    the same style resolution as line metrics; metric-source precedence keeps
    the separately measured line-table rule.
  - The installed Cambria Math cmap omits U+25C6 although the exported PDF
    resource draws it. The catalogue does not invent coverage for that resource
    difference. The next covering CJK face owns the installed stack's metric;
    the PDF comparison remains a resource mismatch.
  - CJK attribution also checks each concrete cut's cmap per glyph in a bounded
    CJK domain, including supplementary Han and kana. Family-wide basic Han
    coverage selects the fallback chain but cannot establish glyph ownership.
    Unrecorded CJK scalars and unknown earlier resources stop attribution.
  - Deck-embedded fonts are distinct resources even when a catalogue name
    matches. Their own bounded cmap coverage is retained at registration and
    follows the same real/synthetic resource selection as their OS/2 metrics.
    Attribution is `complete`, `absent`, `partial`, or `unknown` under the
    fixed internal `canonical-static-v1` library profile. This adopts pinned
    canonical normalization, modern Hangul fallback units, and bounded simple
    Indic/Myanmar syllables; it does not detect the host's shaping engine or
    guarantee arbitrary Canvas font-resource identity. A partial resource
    stops sole-owner attribution to an older subset.
  - Scalar presence requires agreement across every eligible Unicode cmap;
    absence requires exclusion from their union. GSUB/GDEF analysis separately
    proves nonzero outputs, missing-glyph isolation, and freedom from direct or
    transitive erasure for the normalized glyphs. Unsupported mechanisms,
    malformed tables, exhausted budgets, and unsupported script preprocessing
    stay unknown. Context uncertainty invalidates the actual shaping span,
    including authored run seams, without creating extra paint boundaries.
    Installed same-name facts cannot replace embedded-resource uncertainty.
    Established installed named-slot metrics retain their catalogue policy
    and its provenance limits.
  - Metric contributions remain attached to their source ranges on each
    wrapped line. Horizontal measurement and paint share one shaping policy:
    visual style plus the established installed named-family route when every
    unit in the original CSS span has a resolved installed route and its paint
    stack contains no embedded candidate. An unresolved or embedded unit
    disables that extra boundary for the whole span; resource metrics never
    create shaping boundaries. Here "installed" denotes the established named
    slot compatibility route, not proof of host installation or complete cmap
    coverage. Stacked text retains its grapheme-cell model. Explicit
    `fontAlgn` uses the existing fallback when resource fragments inside one
    paint span require different or unknown baseline offsets.
  - A font-slot boundary inside one original grapheme remains attribution
    metadata. Canvas cannot attach a combining mark across separate calls with
    different fonts, so the complete grapheme keeps its base run's paint style
    for measurement, wrapping and painting, including authored run seams.
    Observed Office cross-resource mark selection remains recorded; this
    backend policy does not claim to reproduce its cross-font attachment.
    Separate-grapheme spacing marks retain their existing routing. This does
    not infer a larger script-syllable boundary.
  - Lines containing secondary CJK misses can still differ from Office because
    those drawing resources are intentionally unknown under (B). Recording a
    symbol fallback does not establish the missing CJK resource's metrics.

## Implementation

`packages/pptx/src/east-asian-default.ts` builds the stack in this order:
the selected face, then the symbol fallback, then the CJK fallback faces. It
also names the drawing face when the renderer can know it.

`renderer.ts` uses that stack for measuring, wrapping, painting and stacked or
vertical text. It sizes each line from its known faces only.

The catalogue honors the font's USE_TYPO_METRICS bit even when its OS/2 table
predates version 4, matching the resource parser and the measured symbol
resources. A version gate previously discarded the declared typo metrics.

Deck-embedded fonts are sized from their own font parts' OS/2 tables. These
are parsed when the font is registered, in the window and in the render worker.

Structural cases outside the measured controls still use the ordinary model
for the whole body:

- equations;
- markers taller than the text;
- `compatLnSpc="0"` without Excel tables for every glyph face, including
  unknown-only lines and unknown glyphs mixed with known faces;
- unresolved `fontAlgn` offsets.
