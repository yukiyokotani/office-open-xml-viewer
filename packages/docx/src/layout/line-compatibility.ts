import { isCjkBreakChar } from '@silurus/ooxml-core';
import { defineCompatibilityRule } from './compatibility.js';
import { OFFICE_FAR_EAST_SINGLE_LINE_FACTOR, officeOpenTypeAutoLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';
import type { LineSpacing, TabStop } from '../types.js';
import type { ParagraphLayoutSource } from './text.js';

export const WORD_TAB_DISPLACED_READING_FRAME = defineCompatibilityRule({
  id: 'word-float-tab-reading-frame',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'Stop eligibility stays margin-relative (ECMA-376 §17.3.1.37). With the compatibility mode absent or below 15, an authored stop moves with a float-displaced leading line edge, and a cell whose displaced target lies outside the float window is placed there anyway: it overlaps the float or runs past the margin (and off the page) instead of breaking the line. From mode 15 stops are not displaced; an end/center/decimal cell is limited to the window by WORD_ALIGNED_TAB_CELL_FIT_BACK and a left/start stop past the window breaks the line. Evidence: issue #1672 controls (mode absent; ordinary left/right/center/decimal tabs, one/two tabs, both directions, either float edge, no-float counterexamples; PR #1698, Windows re-export for #1896), and the blk-0898-tabs-float controls of #2047 (Windows Word 16.0.20527, modes absent, 14 and 15; Courier New 12; a 150 x 10 pt square-wrapped float at either column edge narrowing line 1; right stops at 200, 300, 360, 400 and 481.9 pt, center, decimal and left stops at 360/400 pt, a single word and a six-word cell, LTR rows and bidi end/start rows; jc left, both, distribute, lowKashida). Modes absent and 14 are identical: every cell sits at stop + displacement (for example right 360 pt beside a left float ends at 510 pt, past the 481.9 pt margin; right 481.9 pt ends off the page), or over the float when the float is on the trailing side. Mode 15 places every row at its undisplaced stop or limited to the window end. Two pre-15 rows (jc left six-word cell, jc distribute bidi start stop) were instead set below the float with undisplaced stops; their siblings with identical geometry under other justifications were not, and this unexplained exception is not emulated. The renderer reproduces 76 of the 84 #1672 positions exactly (all single-tab and in-window cases); in the 8 two-tab cases the second tab starts past the right margin, where Word advances it 14 pt at 10 and 20 pt (tabs-floats-overflow-font20) while the library keeps the automatic grid (10 pt further), an unexplained residual. Library scope: the overflow applies only when the rest of the paragraph (every later cell to the paragraph end or line break) is no wider than the paragraph, and not in a second gap on the same baseline (WORD_FLOAT_GAP_FLOW); longer content (an empty cell before a long one included) and continuation gaps were not measured and keep the window containment.',
});

/** The compatibility modes in which {@link WORD_TAB_DISPLACED_READING_FRAME}
 * displaces stops and lets out-of-window cells overflow. */
export function wordTabDisplacedReadingFrameApplies(compatibilityMode: number | undefined): boolean {
  return compatibilityMode === undefined || compatibilityMode < 15;
}

/** Projection of {@link WORD_TAB_DISPLACED_READING_FRAME}. The caller has already
 * selected the next eligible stop in margin coordinates. */
export function wordFloatTabStopPosition(
  stopPosition: number,
  custom: boolean,
  leadingShift: number,
): number {
  return custom ? stopPosition + Math.max(0, leadingShift) : stopPosition;
}

export const WORD_TAB_GLUED_TO_FOLLOWING_TEXT = defineCompatibilityRule({
  id: 'word-tab-glued-to-following-text',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'A tab character offers a line break before itself but not after itself: it stays with the text that immediately follows it. When the first unbreakable unit after a tab does not fit behind the tab\'s stop, the line breaks before the tab, which moves to the next line and resolves its stop there; a tab that starts a line keeps its unit, which is then broken by characters in the room left behind the stop. A space after the tab is an ordinary break opportunity (the tab and space end the line). Evidence: #2051, blk-0898-tabs-move controls (Windows Word 16.0.20527, compatibility modes absent, 14 and 15 identical; Courier New 12; jc left, both, distribute, lowKashida and highKashida; LTR left stops and bidi start stops): label + tab + a 24-character word at a 324 pt stop (the tab moves and the word is broken after 21 characters, 8 under the highKashida line-fit reserve); two stops 144/324 pt with a 180 pt label (the word lands whole at 144 pt on line 2); a space before the tab (same); a space after the tab (the word starts line 2 at the margin); a leading tab (the word is broken after the stop on line 1); an automatic 36 pt stop (the word lands at 36 pt); an 80-character word (moved, then broken behind the stop on line 2); a 20-character word that fits (stays). Before, the tab ended line 1 and the word started line 2 at the margin. Library scope (wordTabGlueUnit in line-breaker/tabs.ts): every measured word was one plain run; a tab followed by a word that continues across a formatting seam, holds an optional hyphen, a hyphen/URL or SEA/CJK opportunity, a no-break hyphen, fitText, ruby or tate-chu-yoko, or that lies on a snapToChars grid, a word in a run against the paragraph direction (a left-to-right run in a bidi paragraph: the #1672 r4 controls end such a line with its tabs), a float-narrowed line, and a positional tab, keep the earlier behaviour (the tab may end the line). Intrinsic minimum widths apply the rule to the same input class (a leading stop needs its pen position plus the glued word).',
});

export const WORD_ALIGNED_TAB_CELL_FIT_BACK = defineCompatibilityRule({
  id: 'word-aligned-tab-cell-fit-back',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'ECMA-376 §17.3.1.37/§17.18.84 align the text after an end (right), center or decimal stop to that stop; neither says what happens when the aligned cell does not fit the line. Word fits such a cell as if the tab had no gap and then lets the gap take the line\'s leftover room, up to the stop-aligned gap, so the cell slides back toward the pen and ends at the line (or float window) end. On a justified line that still has justification opportunities the gap stays empty and the slack is justified instead. A start (left) stop never slides: its gap is fixed at the stop and the cell wraps at the line end, in left-to-right and bidi paragraphs alike. Evidence: #2050, blk-0898-tabs-rtl controls (Windows Word 16.0.20527, compatibility modes absent, 14 and 15; Courier New 12; jc left, both, distribute and lowKashida; every row in a bidi paragraph with Hebrew text and as a left-to-right twin): start stops at 72/216/360 pt, a leading tab, an automatic stop and two stops keep their gaps (line 1 holds 11/7/3/7/8/5 four-character tokens, against 12/12/12/13/9/7 when the last gap shrank); a 216 pt start cell overflowing by 0.6 pt wraps its last token; a center stop at 360 pt whose 280.8 pt cell would end at 500.4 pt ends at the 481.9 pt margin instead (start 201.1 pt), and a decimal stop at 360 pt with an overflowing cell starts its cell at 42.6 pt, ending line 1 at the margin, under jc left; under both, distribute and lowKashida that line starts its cell at the pen and is justified. End and center cells wider than the room before the stop start at the pen. #2047 blk-0898-tabs-float controls in mode 15 limit end, center and decimal cells beside a float to the window end (right/center/decimal 400 pt beside a right float end at 331.9 pt). Right- and center-aligned paragraphs with such a cell are unmeasured and treated like left-aligned ones. Because tab stops stay put, justification widens only gaps after a line\'s last tab, soft-wrapped lines included: with two start stops at 144/288 pt the first cell keeps its natural spacing under both, distribute and lowKashida (before, its gap took a share of the slack and the second cell started 4.2 pt late). Library scope: the measured bidi rows held right-to-left (w:rtl) runs after the tab, so in a bidi paragraph only such a cell keeps its start gap or defers its fitted-back gap, while a cell starting with a left-to-right run keeps the earlier provisional gap that the bidi walk shrinks; a decimal cell without a real separator (only the numeric fallback point) keeps the earlier allocation in either direction (the in-place bidi walk, the stop-aligned left-to-right gap). Every measured fitted-back cell was the last on its line, so a cell ended by a later tab keeps the earlier stop-aligned gap (left-to-right) or the bidi walk\'s in-place fit, and only a line\'s last tab takes slack. Bidi lines defer the fitted-back gap like left-to-right ones (the D2 bidi rows: zero gap under both, distribute and lowKashida, the leftover room under left). A line whose slack goes into Arabic kashida elongation counts as justified and keeps the gap empty (unmeasured: the lowKashida rows were Latin and Hebrew). Under a WORD_KASHIDA_LINE_FIT reserve a fitted-back cell is admitted against the band, as an aligned cell is in both directions (centre cells under a reserve are unmeasured there).',
});

export const WORD_POSITIONAL_TAB_AVAILABLE_BAND = defineCompatibilityRule({
  id: 'word-positional-tab-available-band',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'tabs-floats-round3',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Issue #1672 Word exports cover all 36 positional-tab combinations of left/center/right alignment, margin/indent reference, LTR/RTL paragraph direction and no/left/right float. Word intersects the normative reference box (ECMA-376 §§17.3.3.23, 17.18.71, 17.18.73) with the available float band and aligns the following cell in reading order; no-float references remain unchanged. A target at the current pen is reachable without a line break; only a target behind it requires the next line. RTL must retain the positional descriptor through the bidi post-pass rather than resolving it as an ordinary tab. Evidence: issue #1672 controls (36 combinations, PR #1698), re-exported in Windows Word 16.0.20430 for #1896 with every in-band line identical to the earlier export.',
});

/** Reference-box projection of {@link WORD_POSITIONAL_TAB_AVAILABLE_BAND}. */
export function wordPositionalTabReferenceBox(
  referenceStart: number,
  referenceEnd: number,
  bandStart: number,
  bandEnd: number,
  narrowed: boolean,
): Readonly<{ start: number; end: number }> {
  return narrowed
    ? { start: Math.max(referenceStart, bandStart), end: Math.min(referenceEnd, bandEnd) }
    : { start: referenceStart, end: referenceEnd };
}

export const WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS = defineCompatibilityRule({
  id: 'word-fixed-paragraph-auto-spacing-stored-margins',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'paragraph-auto-spacing-stored-margin-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'With doNotUseHTMLParagraphAutoSpacing enabled, Word for Mac retains stored paragraph before/after spacing instead of imposing Part 4 §14.8.3.15 fixed 5pt/10pt automatic margins. Eighteen fixed-setting documents and eighteen HTML-setting counterexamples cover direct and inherited automatic flags, explicit false, missing/zero/5pt/20pt stored values, line-unit conflicts, adjacent automatic paragraphs and page edges. Three faces at 8/12/24pt and a separate 6/18/36pt Normal-style-size sweep show no face- or inline-size-dependent amount. The fixed-setting projection preserves the already resolved stored numerical margins; it does not invent a second line-unit interpreter. Without that setting, the separate consumer HTML-em policy remains unchanged. This is an approved Word compatibility choice, not the normative fixed-pair rule or a claim that other Office versions/HTML consumers share this behavior. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export const LIBRARY_ABSENT_FACE_FALLBACK_LINE_BOX = defineCompatibilityRule({
  id: 'library-absent-face-fallback-line-box',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/reference-font-layout.test.ts#sizes an absent authored face from its painting fallback, not the authored catalog',
  },
  description: 'Library policy (user decision 2026-10-16), replacing the e3286ecb9 absent-face rule: when the authored face is not present on the viewing host, a native CSS family list paints with a browser fallback, and the span is sized from that painting face (its Canvas font box), never from the absent face\'s catalog line box. Word lays out with the authored face it has, so pagination differences from Office caused by the missing font are accepted. Present local faces use their measured host geometry first and the catalog only as a fallback; registered substitutes use their own geometry. The painting face of an absent face is the first available family of the native CSS list; its normal line (Canvas font box plus the CSS normal-line leading, measured per exact list and tuple in the document realm before layout and re-verified by the laying-out realm) sizes the clusters that face is proven to draw (native-route-line-metrics.ts). Clusters drawn by a later fallback, unidentified or unclassified faces, and Far East faces (whose Word allocation uses usWin, which CSS does not expose) keep the Canvas box. No viewer-OS or font-name branch is involved. Library policy, no Office measurement: deliberate project decision (2026-10-16, PR #1844) to lay out with the face that actually paints; explicitly not an Office rule.',
});

export const LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY = defineCompatibilityRule({
  id: 'library-synthesized-face-line-geometry',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/synthesized-face-line-geometry.test.ts#sizes a bold empty mark from the proved face: identical=$sameRaster',
  },
  description: 'Library evidence boundary (user decision 2026-10-18, issue #1848): the reference catalog holds some installed faces only at weight 400 (MS Mincho, for example), so a bold or italic run of such a family resolves to a native route without a catalog tuple. Catalog absence does not show that the host synthesizes the tuple: a genuine styled face may exist with different vertical tables, and Canvas cannot expose OS/2 usWin, which the Far-East 1.3 x usWin line needs. Instead the span is proved to paint a loaded face. Each installed Office face loaded through an exact local() name sits under a private family holding only that FontFace, so by CSS Fonts Level 4 section 5.2 font matching any weight/style request to that family selects it and the browser synthesizes the missing bold/oblique from it. Fixed probe glyphs of the span\'s script class (国あ for East Asian text, Hxg otherwise; 64px) are rasterized once per route and tuple through the native route and through that private family on this realm\'s Canvas. Every grapheme of the span and of the probe must also be owned by that exact source, shown by zero advance through its size-adjust:0% registration (the #1842 coverage probe, loaded for regular faces whose catalog family lacks a bold or italic face), since CSS font matching still falls back per glyph. Pixel-identical non-empty rasters of owned glyphs prove the native route paints the same synthesized face, so the catalog geometry of the CSS-matched loaded tuple applies, including its Far-East projection. Any raster or advance difference (a genuine styled face, another file, glyphs from a fallback face) is no proof and the span keeps its own Canvas box, as before. Main thread and worker each rasterize with their own Canvas and FontFaceSet over the same host fonts. Residual risk: a genuine styled face whose glyphs are pixel-identical to the synthesized regular face but whose vertical tables differ. Scope: only the span\'s own painted face and its paragraph mark are proved; an anchor host (the floating-object anchor character) and the East Asian floor of a run painted in another slot are not, because a Word-exported page keeps a bold Latin heading\'s line when its East Asian slot names a regular-only face.',
});

export const WORD_RUBY_LINE_ASCENT = defineCompatibilityRule({
  id: 'word-ruby-line-ascent',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/ruby-hpsraise-render.test.ts#hpsRaise $raise pt with an 8pt guide over 12pt base text',
  },
  description: 'Windows Word 16.0 build 16.0.20430 #1885 controls: Latin (Arial 10pt), MS Mincho 10.5pt and ruby-only lines with MS Mincho guides at hps 6/10/14/20 (3-10pt), hpsRaise 0/10/18/20/24/28/30/38 and base 10/20pt, under single, auto 1.5, exact 18 and atLeast 18 spacing. ECMA-376 section 17.3.3.12 defines hpsRaise as the distance between base and guide text; Word places the guide baseline hpsRaise above the base baseline and the line ascent is max(base ascent, hpsRaise + guide size), over the base line descent (an MS Mincho base keeps its Far-East 1.3 em box sides). Automatic multiples pin the baseline to the top like text lines, exact 18pt lines keep four fifths, and atLeast lines keep the line descent above the bottom. Every measured line matches within 0.3pt. Yu Mincho controls agree where the host has the face; a missing face uses its painting fallback (LIBRARY_ABSENT_FACE_FALLBACK_LINE_BOX). Ruby without hpsRaise is unmeasured and keeps the retained ink-derived additive reserve.',
});

export const WORD_KERN_THRESHOLD_AUTHORITY = defineCompatibilityRule({
  id: 'word-kern-threshold-authority',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'kern-threshold-and-general-punctuation-slots',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'ECMA-376 §17.3.2.19 makes the resolved style-cascade w:kern threshold authoritative: absence at every level disables kerning and size below the threshold disables it. In 56 informative mode-15 controls (42 boundary rows and 14 calibrations), Word additionally disables a zero threshold; positive thresholds below/equal/above 8–20pt run sizes follow the size comparison, independently of enableOpenTypeFeatures. Twelve null-adjustment controls cannot identify the switch. In mode 15, same-face T + U+0020 across identically formatted source runs retains the font pair adjustment. Windows Word 16.0.20527 controls (Times New Roman and Arial 12pt lines of AV/To/Ta/Wa/LT/Yo letter pairs and of space pairs, explicit w:kern 0 against 2 and absent, styles part absent, without w:kern, or with a docDefaults w:kern 2) show the same zero-disables behaviour with the compatibility mode omitted as in mode 15: every explicit-zero line equals its unkerned control while the threshold-2 lines kern. Modes 11, 12 and 14 were not measured separately for a zero threshold; because the omitted mode and mode 15 agree, an explicit zero disables kerning for content text in every mode. Numbering glyphs and paragraph marks have no measured zero-threshold evidence and retain the previous comparison in every mode. There is no new Office observation for direct letter-pair splits, changed formatting/fonts, explicit w:spacing, other compatibility modes, or unadjusted positive-threshold flag comparisons. Separately, library policy acquires formatting-only source splits as one text sequence in every mode (run-split-invariance.test.ts), preserving semantic run units and independently scoped substitute faces; this invariance is not a broader Office compatibility claim. Three fit contradictions and one flag-dependent break do not establish a different kerning switch or a shaping-table preference. The zero-disables extension is observed Word behaviour beyond the normative positive-threshold rule; no font-specific amount is inferred. Pair kerning that involves U+0020 follows WORD_KERNED_SPACE_PAIRS.',
});

/** Threshold is already resolved by the parser (including explicit zero).
 * WORD_KERN_THRESHOLD_AUTHORITY supplies the zero-disables extension for
 * content text in every compatibility mode (measured with the mode omitted
 * and in mode 15). Generated numbering glyphs and paragraph marks pass
 * `zeroThreshold: 'size-comparison'`: they have no measured zero-threshold
 * evidence and keep the previous size comparison.
 * The positive size comparison and absence default are ECMA-376 §17.3.2.19.
 * Compare declared size, before small-caps/super/subscript paint transforms. */
export function wordKerningApplies(
  fontSizePt: number,
  thresholdPt: number | null | undefined,
  zeroThreshold: 'disables' | 'size-comparison' = 'disables',
): boolean {
  return thresholdPt != null && fontSizePt >= thresholdPt
    && (thresholdPt > 0 || zeroThreshold === 'size-comparison');
}

export const WORD_KERNED_SPACE_PAIRS = defineCompatibilityRule({
  id: 'word-kerned-space-pairs',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'kerned-space-pairs',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'When kerning applies (WORD_KERN_THRESHOLD_AUTHORITY), Word applies the font pair adjustments of pairs that contain U+0020 exactly like letter pairs. Windows Word controls, Times New Roman and Arial at 12 and 30pt, w:kern 2 against absent and 0, compatibility mode omitted and 15, left-aligned single lines of twelve (12pt) or four (30pt) repetitions of "x A ", "A x ", "x Y ", "x T ", "L x ", "\u2019 x " and "x B ": with kerning the space+A, A+space, space+T, space+Y, L+space and \u2019+space pairs all tighten (a 12pt Times New Roman "x A " line spans 229.4pt against 244.9pt unkerned, Arial 229.3pt against 244.6pt, both by the fonts\' -113/2048 em pairs), the "x B " lines, which hold no space pair, are unchanged, and an explicit zero threshold removes every pair. Library projection: a word piece is measured together with its trailing spaces, so the pair between a glyph and the following space is already part of that measurement. Two seams are not: the last space of one piece before the first glyph of the next piece on the same line, acquired by measuring the two pieces joined through the text measurement service, and a glyph and space that §17.3.2.26 font slots split into separate pieces of one face (U+2019 in the highAnsi slot before an ascii space), which the service shapes slot by slot, so the pair is measured from that face as the two glyphs against their isolated advances. Either amount comes from the selected face (never a constant) and is charged to the following piece\'s leading edge wherever that piece\'s advance is used: line fitting, the visible prefix before hanging punctuation, hyphen and optional-hyphen prefix admission, the width a kinsoku retraction keeps, right/center/decimal and positional tab cells, max-content sizing, committed width and painted glyph origins. The pair adjusts the advance of the seam\'s first character, so under WORD_CONDENSED_SPACING_ADVANCE_FLOOR it is folded into that character\'s committed contextual advance before its zero floor (max-content sizing keeps such seams, font-slot seams, every seam bounded lines leave unkerned and every kerned change of source sequence between pieces, so it gives each pair the same owner and kerns no excluded seam), and no piece is charged more than its own advance. Scope, library policy where unmeasured: both pieces in one identically formatted source text sequence, horizontal non-complex left-to-right Latin text of one face, size, weight and style with kerning applying to both; a change of face, size or any other formatting across the space, a size below the threshold, a character-grid snapped cell, fitText, any piece of a ruby run, a tate-chu-yoko cell, vertical or right-to-left text takes no cross-space pair (Word with differing faces or sizes across a space was not measured). A piece that starts a line takes no pair with the space ending the previous line; the line-end space keeps its pair with the glyph before it inside the hanging space (WORD_LINE_END_SPACE_HANGS), also when a font slot split them, so the visible prefix keeps its natural advance for fitting and alignment. Justification distributes its slack on top of the kerned natural advances; mode-15 justified compression already acquired this pair at its word boundaries (WORD_JUSTIFIED_INTERWORD_COMPRESSION). Modes 11, 12 and 14 were not measured; letter-pair kerning is not mode-gated either, so the pair applies in every mode. The amount is whatever the host face and the browser shaping yield: on macOS, Chromium shaping does not apply the space pairs of the Times New Roman 5.01 face (they are present in its legacy kern table), so those lines remain unkerned there, while macOS Arial reproduces the Word spans.',
});

export const WORD_ARABIC_PAIR_KERNING = defineCompatibilityRule({
  id: 'word-arabic-pair-kerning',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/arabic-kerning.test.ts#kerns Arabic text whatever w:kern says',
  },
  description: 'Windows Word 16.0 build 16.0.20430 #1850 controls with the Arabic supplemental fonts installed: "Z <Arabic run> X" paragraphs at 14pt in Sakkal Majalla, Traditional Arabic, Simplified Arabic, Arabic Typesetting, Arial, Times New Roman and Tahoma, three texts, w:kern absent/0/2/40 and enableOpenTypeFeatures off/on. Every Arabic run width is identical across all eight settings, and for the four faces whose GPOS kern feature has Arabic pairs it equals HarfBuzz shaping with kerning (within 0.15pt) rather than without (1.4-4.8pt wider), measured on the same Windows font files. Word therefore applies pair kerning to Arabic text unconditionally; ECMA-376 section 17.3.2.19 w:kern governs the other scripts (WORD_KERN_THRESHOLD_AUTHORITY). Arial, Times New Roman and Tahoma carry no Arabic kern pairs, so they cannot distinguish the rules. Only the Arabic script was measured; other complex scripts keep w:kern.',
});

/** WORD_ARABIC_PAIR_KERNING: Arabic-script text is always pair-kerned. The
 * decision is per measured group or segment (Arabic text sits in runs of its
 * own); a group that mixes in letters of another script is unmeasured. */
export function wordArabicTextKerned(text: string): boolean {
  void WORD_ARABIC_PAIR_KERNING;
  // Only Arabic-script letters were measured: text mixing in letters of
  // another script keeps the w:kern decision.
  return /\p{Script=Arabic}/u.test(text)
    && !/[^\p{Script=Arabic}\p{Script=Common}\p{Script=Inherited}]/u.test(text);
}

export const WORD_NUMBERING_MARKER_FIRST_LINE_UNION = defineCompatibilityRule({
  id: 'word-numbering-marker-first-line-union',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-font-size-line-box-matrix',
    application: 'Microsoft Word',
    version: '16.113.2 and 16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'In a non-grid paragraph with 1.15 automatic spacing, a text marker participates in the first-line ascent/descent union. Relative to a marker-free control, 8, 14, and 20 pt markers added 0, 1.68, and 7.68 pt to the line advance; changing the marker alone shifted subsequent paragraph baselines by the same amount. Added automatic leading follows the body text single-line height rather than scaling the taller marker box. A further 108-case Word matrix used three Latin body faces at 9 and 12 pt with smaller, larger, same-size, bold, and symbol-face markers. All 36 paired omitted-spacing and explicit-auto1 controls had identical marker-added advance; a 9 pt body with a same-size symbol-face marker gained about 0.72 pt in both modes. These controls support implicit marker participation, including a different marker face. They do not establish portable face identity or exact PDF baseline quantization. Library allocation uses the same selected-resource/reference vertical admission as body text, separately from Canvas advance and ink; pixel-rounded Canvas marker sides cannot inflate an otherwise precise same-face line. Exact spacing, grids, ruby, wrapping floats, picture markers, and continuation lines are outside the measured scope. Windows Word 16.0 build 16.0.20430 #1884 controls supersede the descent part for the selected Windows rule: with 6-20pt Symbol bullets and Arial numbers over Arial 10pt and Times New Roman 12pt text, auto 1/1.5/2 and atLeast 18/30/60, the marker raises the first line by its ascent (with line gap) only and the next line follows one text line below, so the marker descent and single-line height do not enter the union. atLeast lines use the same union, and wrapping floats do not suspend it (#1883 controls). An empty item\'s text marker joins its mark line the same way and the whole box follows an automatic multiple.',
});

/** WORD_NUMBERING_MARKER_FIRST_LINE_UNION: Windows Word 16.0 controls
 * (2026-10-23) give identical body baselines in omitted/12/14/15 modes across
 * 78 cases: Ubuntu 12pt, Arial 10pt and Times New Roman 12pt bodies; Symbol,
 * Wingdings, Courier New and Arial markers at 6/12/20pt; auto 1.15/1.5, with
 * atLeast/exact 18pt counterexamples. A marker can raise the first ascent,
 * but its descent does not increase the step to the continuation/successor.
 * This extends the earlier mode-15 observation, not an ECMA line-box formula.
 * The caller retains eligibility: only mode 15 admits floats and atLeast;
 * legacy atLeast participation remains unsupported despite the new controls.
 * Other modes and grid/ruby/fixed-line exclusions keep their existing policy.
 */
export function wordTextMarkerKeepsBodyDescent(compatibilityMode?: number): boolean {
  void WORD_NUMBERING_MARKER_FIRST_LINE_UNION;
  return compatibilityMode == null || compatibilityMode === 12
    || compatibilityMode === 14 || compatibilityMode === 15;
}

export const WORD_PICTURE_BULLET_LINE_PARTICIPATION = defineCompatibilityRule({
  id: 'word-picture-bullet-line-participation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'picture-bullet-line-height-matrix',
    application: 'Microsoft Word',
    version: '16.0',
    platform: 'Windows',
  },
  description: 'Issue #1882 controls: DOCX picture bullets (15 and 2.25 pt displayed, and 30, 7.5, 20 and 83 pt size variants) and their Word-saved DOC (40 and 6 pt), Arial 10 and 24 pt text, single, 1.5 and 2.0 automatic, atLeast 18/60 pt and exact 18 pt spacing, 18 and 24 pt line grids with 10.5 pt MS Mincho text, a ruby line, empty list items, and a square-wrapped float. ECMA-376 §17.9.20 defines the picture as the bullet but not its line allocation. The bullet bottom stands on the first baseline. It joins the first line as an inline picture of its displayed height: the ascent is the larger of text ascent and bullet height; automatic multiples add only the text single-line leading (natural + single times (m - 1)); atLeast keeps the larger of the authored minimum and the natural box; a ruby line forms one union with it. Exact spacing ignores it: tall and short bullets, text and empty items all keep the same 18 pt box and a baseline 14.1-14.3 pt below its top, and the bullet overflows above. On an East Asian grid line the bullet counts cells together with the text below the baseline: a 15 pt bullet over 10.5 pt text takes two 18 pt cells (baseline centred at 96.0 pt) and one 24 pt cell; a Latin line keeps its natural box (15 pt bullet, Arial 10 pt: one 18 pt cell). An empty item takes the bullet height as its mark ascent and the whole box follows the line rule (40 pt bullet, 10 pt mark: 42.1, 63.3 and 84.3 pt at 1, 1.5 and 2; atLeast 42.1 and 60 pt; 3 cells of 18 pt). In a document with East Asian text the mark beside a bullet keeps its non-East-Asian face: a 15 pt bullet over a 10.5 pt Arial/MS Mincho mark fills one 18 pt cell with the Arial descent (baseline 87.35 pt), where the 1.3 em East Asian box would take two. Under 1.5 and 2.0 automatic multiples an empty item draws the bullet at the top of the box (15 and 40 pt bullets start at the line top; 2.25 and 6 pt ones stand at the mark ascent), the paint rule of word-auto-multiple-baseline-pin; pagination keeps the centred trailing-mark metric. Numbered-line baselines under atLeast/exact, float-wrapped list items and ruby line height are separate pre-existing differences (#1884, #1883, #1885).',
});

export const WORD_EAST_ASIAN_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-east-asian-grid-line-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'word-font-metrics-resource-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Across eight East-Asian resources, six sizes, and no-grid/grid/useFELayout variants, Word allocated grid cells from the selected face line height, rounding up to whole pitches. Historical Word for Mac evidence from a later controlled font-table intervention found that OS/2 code-page bits 17–20, rather than cmap/script alone, selected the 1.3-times-hhea line box. The grid-cell rule takes the resulting height as input; it does not itself classify the face. Mixed rFonts slots also prevent a family-wide Latin override. Windows Word 16.0 build 16.0.20430 retest: The whole-pitch rounding pattern is retained, but Yu Gothic single-line height differs from the recorded Mac hhea-based formula. At size/pitch 8/12, 12/20 and 20/30pt, that formula predicts 12/20/30pt grid allocations; Windows produces 24/40/60pt. Snap-disabled Yu Gothic steps are about 13.4/20.1/33.4pt at 8/12/20pt, versus Mac-formula heights 11.46/17.19/28.65pt. MS Mincho, MS Gothic and Meiryo allocations agree with the Mac formula in this sweep. Adjudication (2026-10-16): use each selected host face\'s OS/2 usWin ascent/descent for the bits 17–20 class, with 0.15 times their sum added to each side; use Windows output when identical vertical data differs across Office platforms. The tested Mac Office and Windows PDF Yu Gothic faces have hhea=(1802, -455, 1024), usWin=(2017, 619), UPM=2048: 1.6732421875em yields 24/40/60pt cells at size/pitch 8/12, 12/20, 20/30pt. MS Mincho, MS Gothic and Meiryo have equal hhea/usWin boxes and retain their allocations. Mac hhea evidence remains historical; mixed slots and absent exact-source Mac PDFs still limit the platform comparison. No viewer-OS or font-name branch is introduced. Missing usWin declines projection, and other faces retain signed hhea leading.',
});

export const WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT = defineCompatibilityRule({
  id: 'word-table-cell-ignores-grid-right-indent-adjustment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'table-cell-adjust-right-indent-width-position-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed linesAndChars matrix, paragraphs inside fixed-width table cells retain the same line breaks for omitted (default true) and explicit-false w:adjustRightInd across four boundary widths and both left/right cell positions. Scope this Word-only exception to table-cell containers; ordinary body paragraphs retain the ECMA-376 §17.3.1.1 adjustment.',
});

export const WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT = defineCompatibilityRule({
  id: 'word-snap-to-chars-east-asian-cell-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'snap-to-chars-east-asian-cell-fit-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For snapToChars, Word centers each East-Asian grapheme independently in the smallest whole number of character-pitch units that contains its natural advance. A grapheme that fits uses the one-unit placement described by [MS-OI29500] §2.1.534; an undersized authored pitch expands only that grapheme to additional units.',
});

export const WORD_SNAP_TO_CHARS_SCRIPT_BLOCK_ALLOCATION = defineCompatibilityRule({
  id: 'word-snap-to-chars-script-block-allocation',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.534',
  },
  description: 'Allocate snapToChars Latin text in contiguous blocks centered across the required grid units, complex-script blocks from their leading edge, and East-Asian graphemes independently by character cell.',
});

/** Word compatibility projection governed by
 * {@link WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT}. */
export function wordSnapToCharsEastAsianCellCount(
  naturalAdvancePt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(naturalAdvancePt)) return 1;
  return Math.max(1, Math.ceil(Math.max(0, naturalAdvancePt) / pitchPt - 1e-9));
}

/** Compatibility projection governed by
 * {@link WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT}. */
export function wordContainerAllowsGridRightIndentAdjustment(
  insideTableCell: boolean,
): boolean {
  return !insideTableCell;
}

export const WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT = defineCompatibilityRule({
  id: 'word-grid-right-indent-pitch-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'grid-right-indent-character-pitch-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For body paragraphs whose ECMA-376 §17.3.1.1 adjustment is enabled on a linesAndChars character grid, Word reduces the physical line width to the greatest whole character-pitch multiple not exceeding the available width. The observed matrix covers exact and non-exact widths, zero and negative charSpace, explicit opt-out, line-only control, both physical indent sides, and the separately registered table-cell exception.',
});

/** Word compatibility projection governed by
 * {@link WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT}. */
export function wordGridRightIndentAdjustmentPt(
  availableWidthPt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(availableWidthPt) || availableWidthPt <= 0) {
    return 0;
  }
  const remainder = ((availableWidthPt % pitchPt) + pitchPt) % pitchPt;
  const epsilon = 1e-9;
  return remainder <= epsilon || pitchPt - remainder <= epsilon ? 0 : remainder;
}

export const WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE = defineCompatibilityRule({
  id: 'word-hanging-tab-same-position-precedence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'hanging-indent-authored-tab-collision-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When the implicit tab created by a hanging indent shares its coordinate with an authored center, end, or start stop, Word resolves one advancing stop at that coordinate using the authored alignment. An authored bar remains an independent drawing rule, so the implicit advancing stop survives beside it. If center/end alignment would place following text before the current pen, the tab contributes zero advance.',
});

/** Compatibility projection governed by
 * {@link WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE}. */
export function wordAuthoredTabReplacesImplicitHangingStop(
  alignment: TabStop['alignment'],
): boolean {
  return alignment !== 'bar' && alignment !== 'clear';
}

export const WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT = defineCompatibilityRule({
  id: 'word-rtl-decimal-tab-physical-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'rtl-decimal-tab-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For LTR numeric cells embedded in a bidi paragraph, Word aligns the physical left edge of the first halfwidth period to the decimal stop across source-run boundaries. When no period exists, it aligns the numeric cell\'s physical right edge to the stop.',
});

export const WORD_DECIMAL_TAB_SEPARATOR_RESOLUTION = defineCompatibilityRule({
  id: 'word-decimal-tab-separator-resolution',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.556',
  },
  description: 'Use the first explicit halfwidth period as the decimal-tab alignment point; when absent, use the implicit separator after the final digit of the first Unicode decimal-number sequence.',
});

export const WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM = defineCompatibilityRule({
  id: 'word-use-fe-layout-inherited-grid-minimum',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-visible-script-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a visible Latin line with a resolved eastAsia font axis participates in Far East grid metrics even when w:rFonts@hint is absent; inherited automatic spacing keeps the larger of its whole-cell design allocation and one grid pitch multiplied by the inherited spacing value.',
});

export const WORD_USE_FE_LAYOUT_EMPTY_MARK_GRID_ALLOCATION = defineCompatibilityRule({
  id: 'word-use-fe-layout-empty-mark-grid-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-empty-mark-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a content-less paragraph mark participates in Far East whole-cell document-grid allocation even when the document contains no literal East Asian text. Its face-specific Far East design height governs the cell count; exact spacing and snapToGrid=false remain the document-grid overrides named by ECMA-376 §17.6.5. Observed Word output gives signed atLeast spacing a discontinuous boundary on an active grid: negative values use their absolute magnitude as the mark advance, zero keeps the ordinary atLeast-zero advance regardless of inheritance source, and positive values retain whole-cell allocation.',
});

/** Compatibility projection governed by the useFELayout empty-mark allocation
 * and {@link WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED}. The caller supplies the
 * ordinary line-spacing result and the mark's whole-cell grid allocation.
 * Exact spacing is the normative §17.6.5 override. Observed Word output gives
 * signed atLeast values an empty-mark-specific negative/zero/positive boundary. */
export function wordUseFeLayoutParagraphMarkGridAdvancePx(
  input: Readonly<{
    ordinaryAdvancePx: number;
    allocatedGridAdvancePx: number;
    atLeastZeroAdvancePx: number;
    lineSpacing: LineSpacing | null;
    gridAllocationActive: boolean;
    scale: number;
  }>,
): number {
  const {
    ordinaryAdvancePx,
    allocatedGridAdvancePx,
    atLeastZeroAdvancePx,
    lineSpacing,
    gridAllocationActive,
    scale,
  } = input;
  if (!gridAllocationActive) return ordinaryAdvancePx;
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value < 0) {
    return Math.abs(lineSpacing.value) * scale;
  }
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value === 0) {
    return atLeastZeroAdvancePx;
  }
  return lineSpacing?.rule === 'exact'
    ? ordinaryAdvancePx
    : Math.max(ordinaryAdvancePx, allocatedGridAdvancePx);
}

export const WORD_CONTIGUOUS_UNDERLINE_GEOMETRY = defineCompatibilityRule({
  id: 'word-contiguous-underline-geometry',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/paragraph.test.ts#keeps a solid underline continuous across floating-precision retained run seams',
  },
  description: 'Adjacent compatible underlined source runs share one safe baseline and continuous authored cadence while style, color, and thickness boundaries remain distinct. Library policy, no Office measurement: numerical robustness of contiguous underline seams (PR #1463) that keeps the ECMA-376 §17.3.2.40 geometry; not an Office claim.',
});

export const WORD_INLINE_OBJECT_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-inline-object-grid-line-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'inline-picture-grid-line-allocation',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Word 16.0 build 20430 on Windows, compatibility mode 14, allocates a single-spaced inline-picture line to ceil(natural height / line pitch) whole cells and centers the object in the extra space. ECMA-376 §20.4.2.8 makes inline objects affect line height like glyphs; §17.6.5 and §17.3.1.32 define grid participation but not this tall-object rounding. Library policy reuses the existing grid allocation and centered baseline when inline pictures, charts or shapes supply the natural ascent and the font design height does not exceed that natural box, with automatic or omitted spacing. Smaller objects retain the taller text line\'s grid policy. The Office evidence establishes the single-spaced picture class only; automatic multiples and other inline payloads reuse that allocation without claiming separately measured Office parity. Off-grid, snap-disabled, exact and atLeast lines retain their existing paths; floating-anchor reference origins are independent. Evidence: issue #1674 / PR #1796 (a 221.2pt picture takes 240pt of grid).',
});

export const WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED = defineCompatibilityRule({
  id: 'word-grid-at-least-tall-line-unsnapped',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/line-box-height.test.ts#does not round tall East Asian content up to an additional grid cell',
  },
  description: 'An explicitly authored atLeast line on an active document grid keeps the maximum of its natural height, authored minimum, and one pitch instead of rounding tall content to another whole cell. Classification: normative reading of ECMA-376 §§17.3.1.33, 17.18.48 (atLeast) and 17.6.5; not a Word measurement.',
});

export const WORD_DEGENERATE_LINE_SPACING_SINGLE = defineCompatibilityRule({
  id: 'word-degenerate-line-spacing-single',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-DOC] §2.9.146',
  },
  description: 'Preserve a non-collapsing single-line fallback for exact or automatic line spacing at or below zero, consistent with the native LSPD representation.',
});

export const WORD_AUTO_MULTIPLE_BASELINE_PIN = defineCompatibilityRule({
  id: 'word-auto-multiple-baseline-pin',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'auto-multiple-baseline-pin',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Paint a positive automatic line-spacing multiplier with its glyph baseline pinned inside the single design line, placing extra leading or compressed overflow toward block-end; this is draw-only and does not replace the centered trailing-mark pagination metric.',
});

export const WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS = defineCompatibilityRule({
  id: 'word-mixed-anchor-visible-line-metrics',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/anchor-host-metrics.test.ts#reserves host line height without using its zero-ink box for a visible run baseline',
  },
  description: 'A zero-ink drawing anchor host reserves its line and grid height while visible neighboring glyphs retain their own ascent, descent, and design-line baseline. Library policy, no Office measurement: line metrics choice for mixed anchor and visible content (PR #1035); nothing measured.',
});

export const WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION = defineCompatibilityRule({
  id: 'word-justification-leading-indent-exclusion',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/text-distribute.test.ts#forwards (segs, slack, firstContentSi, lastDrawnSi) positionally',
  },
  description: 'Keep leading whitespace used as a first-line text indent fixed while distributing justified-line slack across content in a left-to-right line. Library policy, no Office measurement: kept the existing behaviour in PR #483; nothing measured.',
});

export const WORD_COLLAPSIBLE_LINE_EDGE_VISIBLE_FIT = defineCompatibilityRule({
  id: 'word-collapsible-line-edge-visible-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'visible-prefix-fit-sweep',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Library fitting uses the visible prefix, excluding a collapsible U+0020 edge separator. ST_Jc (§17.18.44) defines inter-word justification, not an edge-space admission charge. Word mode-14 mixed-format Latin output retains a naturally fitting prefix when only its edge separator exceeds the band; mode-15 compression controls use the same visible prefix. Canonical joining must not turn formatting-only source seams into different admission decisions. This replaces an unsupported separator-charge policy, without adding compression to older modes. RTL separately retains the complete advance because its right-edge origin can otherwise move visible LTR cells outside the band, except on the justified lines WORD_RTL_JUSTIFIED_VISIBLE_FIT measured. Other glyphs and authored fixed-width/atomic units retain their existing fit contracts. Evidence: Word compatibility-mode 14 output in PR #1703, and the #1896 Windows sweep (Arial 10pt, a bold run before the line-end space): a line keeps its last word at 110.25pt and not at 110.0pt, the 110.06pt visible prefix, in mode 14 and in a mode-15 left-aligned line.',
});

export const WORD_LINE_END_SPACE_HANGS = defineCompatibilityRule({
  id: 'word-line-end-space-hangs',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'justify-1934',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'The U+0020 and U+3000 spaces that end a left-to-right line (soft-wrapped, ended by a manual break, or paragraph-final) hang past it: they take no part in justification or in centre/right alignment, and Word paints no underline, double underline, strike-through, double strike-through, highlight or shading over them; with ulTrailSpace (§17.15.3.8) the hanging spaces keep all of these, clipped at the right edge of the band of the line (decorated wrapped and paragraph-final U+0020/U+3000 tails, both orders, space-only runs; right, centre and justified lines; with and without a linesAndChars grid, which changes nothing). Windows Word controls in compatibility modes unset, 11, 12, 14 and 15 (Times New Roman 12pt and Arial 11pt, plain, bold and larger runs, one or three trailing spaces, ideographic-space separators, a tab earlier on the line, spaces in their own decorated runs, a text box) end every justified line flush at the right margin (both, distribute, lowKashida and highKashida within 0.6pt; mediumKashida and thaiDistribute lines also end flush but Word breaks those Latin controls at other words, a separate line-breaking difference) and place centre/right-aligned lines by their visible text, wrapped, break-ended and final alike (one, three and ideographic trailing spaces before a break: right +3.2/+9.2/+12.1pt, centre half that, before this rule). Mode-15 justified lines fit their visible width (WORD_JUSTIFIED_INTERWORD_COMPRESSION), which drops line-end U+0020 and U+3000 alike: a line-end U+3000 that overflows the margin hangs and the visible text ends flush (blk-0898-u3000-mode15: Latin words with U+3000 separators 538.6pt against Word 538.5pt, CJK lines after Latin prefixes within 1.5pt; before, 11-12pt short); the whole U+0020/U+3000 tail hangs and the U+0020 gaps opening into it take no slack, while wordWrap-off paragraphs and lines with right-to-left text keep their U+3000 in the visible width. Paragraphs with wordWrap off (§17.3.1.45) keep every trailing space in the line, measured for U+0020, three U+0020 and U+3000 on right and centre lines, bordered or not, Latin and CJK, in modes unset and 15 (right: 3.0, 9.0, 12.0pt short of the margin). Right-to-left lines and other white space characters (NBSP, tab) keep their previous geometry. Underlines under a snapToChars grid keep WORD_SNAP_TO_CHARS_TERMINAL_UNDERLINE, observed on Word for Mac to cover an authored trailing space.',
});

export const WORD_LINE_END_JUSTIFICATION = defineCompatibilityRule({
  id: 'word-line-end-justification',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'How a justified paragraph (§17.18.44) justifies a line that does not end at a soft wrap. A line ended by a manual break (§17.3.3.1 br textWrapping, also clear=all, and cr) is justified like a soft-wrapped line unless doNotExpandShiftReturn (§17.15.3.5) is set; a single word stays at its natural width; under lowKashida, mediumKashida, highKashida and, before mode 15, thaiDistribute a space-free (CJK) line stays ragged. A line ended by a page or column break is justified like a soft-wrapped line before compatibility mode 15 and like a final line from mode 15, whatever doNotExpandShiftReturn says. Final lines stay unjustified, except distribute: before mode 15 (mode unset, 11, 12, 14) its final, break-ended and page/column-break-ended lines are spread by an equal advance after every grapheme cluster but the last, spaces included (fit residual 0.1-0.2pt), with hanging line-end spaces outside; from mode 15 they are justified at their gaps and centred when they have none (Latin single words; CJK lines widen their inter-character gaps). From mode 15 a thaiDistribute line ended by a manual break is spread per character too, single words included. Only gaps after the last tab are justified on these lines (kashida and automatic East Asian spacing included), and a line with a tab is not centred (unmeasured). Right-to-left lines keep word justification. Existing mode-15 Word controls with unmarked Hebrew runs and a manual break spread ordinary source-space gaps flush across the band. Validated acquired source-space cells preserve those opportunities and their original run scopes; the prior segment-seam restriction was a library representation limitation, not a Word rule closing these spaces. Unsupported character-style, mixed-visible, terminal-budget and rewrite geometry keeps the original parent view and its alignment limits; a compressed mode-15 line keeps the opportunities its fit used. Fixed-pitch cells (fitText, character grid) stay whole in the character spread, and hanging ones take no part in it; one fitText region (§17.3.2.14) split over several runs stays whole, and so does a grapheme cluster continuing across a formatting seam (UAX #29 over the contiguous text: combining marks, regional-indicator pairs, emoji ZWJ sequences), while a seam that only forbids a line break still opens a gap. doNotExpandShiftReturn does not affect distribute (blk-0898-brk-nexp controls). Synthetic Windows Word controls blk-0898-brk-* and -hard-* in modes unset, 11, 12, 14 and 15, with and without doNotExpandShiftReturn, for both, distribute, low/medium/highKashida and thaiDistribute: Latin multi-word, single-word, CJK, mixed, tab, RTL Hebrew, zero-margin table cell, br clear=all, cr, break at the paragraph end, column and page breaks. All lines but mixed Latin/CJK lines match per character within 2.4pt; mixed lines keep the separate CJK/space weighting of ordinary justification.',
});

export const WORD_KASHIDA_LINE_FIT = defineCompatibilityRule({
  id: 'word-kashida-line-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-kfit',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Line fit and soft-line justification of the kashida and Thai justification types (§17.18.44). A mediumKashida line takes the next word only while its natural width (hanging line-end spaces excluded) stays within its available width less a 10% reserve, and a highKashida line less a 20% reserve; lowKashida, both, distribute and thaiDistribute fit the whole width. From compatibility mode 15 the reserve is that share of the line\'s own width, so the line fits into 0.9 / 0.8 of the width between its indents, the first line after its first-line or hanging offset. In modes unset, 11, 12 and 14 (identical in every control) the reserve is that share of the width from the line\'s start edge (the paragraph\'s start indent, w:left / w:start in either direction, or a float edge narrowing that side; first-line and hanging offsets do not count) to its far edge (a float edge narrowing that side, else the column or table cell edge past the trailing indent), taken from the line\'s own width: a trailing indent and a first-line indent do not shrink it. An aligned (right) tab cell takes no reserve: it may reach its stop anywhere in the band, while the text before the tab keeps the reduced fit width. The reserve narrows the fit width of the line breaker only, the force-break of an over-wide word or CJK run included; paint still justifies every soft-wrapped line to the full width, so the reserve is spent by stretching. A soft-wrapped lowKashida, mediumKashida, highKashida or thaiDistribute line opens no inter-CJK gaps: its slack goes to its spaces, and a space-free CJK line keeps its natural width at the line start (both and distribute still spread CJK). Evidence: synthetic Windows Word 16.0.20527 controls blk-0898-kfit-mode{none,14,15} (K1-K6 Latin with short and long words, double spaces, Arial 10pt and Times New Roman 20pt; K7 MS Mincho CJK; K8 Tahoma Thai; full 481.9pt width), blk-0898-longword-mode{none,15}, blk-0898-kfitn-mode{none,15} (Times New Roman 12pt and Arial 10pt at full width, a 113.4pt right indent, 141.75pt indents on both sides and a 300pt right indent) and blk-0898-kfit5-mode{none,11,12,14,15} (right, both-side and left-only indents, a 70.9pt first-line indent, a 70.9pt hanging indent, left/right indents with a first-line indent, a 300pt cell of a table indented 50pt, and right-to-left Arabic paragraphs with no indent, a w:right and a w:left indent) and blk-0898-kfit6-mode{none,14,15} (Latin and Arabic bidi paragraphs, Arabic in left-to-right paragraphs, no / end / start indent) and blk-0898-kfit7-mode{none,15} (a tbRl section, vert and eaVert text boxes, 300pt cells with 5.4pt and 20pt margins and a first-line indent, 150pt square-wrapped floats at either column edge, right tabs, Thai). Admissible fit widths at full width are [0.89909, 0.90179) and [0.79931, 0.80018) of 481.9pt in every mode. Before mode 15, for example: right indent 113.4pt medium [320.2, 322.4) against 320.3; left indent 113.4pt high [283.6, 301.9) against 294.8; both-side indents high [129.6, 130.6) against 130.4; hanging indent medium first line [426.5, 441.5) against 440.8 and later lines [367.9, 388.9) against 369.9; first-line indent medium first line [336.2, 367.9) against 362.8. In mode 15: right indent 113.4pt high [286.9, 301.9) against 294.8; first-line indent medium first line [367.9, 406.2) against 369.9; both-side indents high [155.7, 159.0) against 158.7; the cell high [237.3, 242.2) against 240 in every mode. Cells with margins in mode 15: 0.9 / 0.8 of the cell text width (5.4pt margins: medium [258.9, 265.9) against 260.3, high [224.6, 235.9) against 231.4; 20pt margins: high [208.0, 208.9) against 208.0), and the first-line indent in a cell likewise. Float-narrowed lines (a 322.9pt window), in both modes: medium [289.2, 294.9) against 290.6 and high [251.3, 265.9) against 258.3, where a reserve of the column (274.7, 226.5) is contradicted. Right tabs: a highKashida cell ending at 425pt (past the 385.5pt fit) and a mediumKashida cell ending at the margin stay on their lines, while highKashida text before a margin tab wraps at the fit width. Vertical text: the tbRl Latin lines (728.5pt) and the mode-15 CJK lines break as horizontal lines of that length do. A fixed fraction of the indented width before mode 15, a reserve of the column for a float-narrowed line, a constant amount, or a reserve per space, character or em is contradicted. The renderer reproduces every Latin, CJK and Thai medium/high paragraph of these controls line for line except the mode-unset K7 CJK paragraphs, and K7 mixed label+CJK lines put all slack into the space, within 1.7pt per character. The kfit5 cell matches once its geometry matches Word: without a styles part Word gives that cell no left/right margins while the renderer applies 5.4pt (a separate table-geometry difference). Right-to-left paragraphs take w:left / w:start as the start indent and fit by their visible prefix (WORD_RTL_JUSTIFIED_VISIBLE_FIT); with it every kfit5 Arabic paragraph matches before mode 15, while in mode 15 the R0 paragraphs carry a separate Arabic measurement difference. The renderer reproduces the kfit7 tbRl Latin paragraphs, mode-15 tbRl CJK paragraphs, cells in mode 15, the right-floated paragraphs, the vert text box mediumKashida paragraph, the right tabs not beside a float, and the spaced Thai paragraphs line for line (Thai base characters within 0.4pt after the label line under both and the kashida types; thaiDistribute and final distribute lines keep their separate spreading differences). Differences there that are not this rule: before mode 15 Word sets cell text at another position and width than mode 15 (both included); vertical text boxes are laid out with a longer line than the renderer gives them (a one-character difference in both, a one-word difference in highKashida); a right tab beside a float, the label line of a Thai paragraph, a long Thai chunk and the pagination of the left-floated paragraph differ for every justification type. Left and leading tabs (blk-0898-kfit8-mode{none,15}: left stops at 72, 216 and 360pt after a label and before the text, both/medium/highKashida) keep the fit width for the text after the stop and match Word line for line within 0.6pt, except the highKashida label line with the 360pt stop, where Word carries the tab to the next line with its first word and breaks that word inside (a tab-wrap behaviour not reproduced). Right-to-left start stops differ for every justification type, both included (the renderer shrinks the last start-stop gap as its cell fills the line, where Word keeps the gap at the stop and wraps), a separate right-to-left tab difference. A float edge on the far side close to the band edge (blk-0898-kfit9-mode{none,14}: Courier New, a 57pt right indent, medium and high, the float edge 1-10pt inside the band edge) still ends the reserve base: every first line beside it holds the word whose end lies between the float-edge fit and the fit of the band edge less the trailing indent, as this rule gives. The renderer admits those lines by this rule; its float search after a rejected head unit asks the unchanged float solver on the real band for the smaller of the interior and far-edge widths, taking a window that holds its own kind\'s width, else steps on the line grid up to where the larger width is found; past 32 such steps on one line it takes the window of the larger width (library policy for an unmeasured geometry: a long run of far-edge windows that each fail their own width). Before mode 15 any negative trailing indent removes the reserve: the line fits its whole width (blk-0898-floats-negind-mode{none,11,12,14,15}: a 100-character Courier New word force-broken at the fit, medium and high, start indents 0, 176.35 and 300pt each also 3.6pt further in, trailing indents -5 to -56.7pt: all 76 first lines in modes unset, 11, 12 and 14 hold floor(W / advance) characters, where a reserve of the width to the column edge (the rule for a non-negative trailing indent, which the same controls confirm at 0 and +40pt in 24 of 24 lines), of the line\'s own width or of the width to the page edge is contradicted in every line; mode 15 keeps its reserve of the line\'s own width in all 76; a float 1pt past the band edge and word-fit lines change nothing; kfit9, -40pt: a highKashida line of a 345.5pt window holds 316.9pt). Past the page edge (-80 and -120pt) Word ends the line about 5.6pt past the page edge for every justification type, both included, a separate negative-indent limit the renderer does not apply. Unmeasured: centre tab cells under a reserve and first-line indents beside a float. Unresolved (blk-0898-floats-strip, issue #2059: a left float makes the far-edge window too narrow for its reserve, a right float edge 1/4 or 3/4 of the far-edge strip inside the band follows on the next line): in mode 15 every line matches; before mode 15 a paragraph that is the first on its page, with a tight left float, sets the word below that float rather than in the strip window this rule admits (by 1.4pt medium, 2.9pt high), while a square left float breaks it on its first line, and the other paragraphs show the separate pre-15 behaviour of a paragraph anchoring its own floats; the renderer keeps its search there. Unreproduced: in mode unset Word sets the Latin label of a low/medium/highKashida CJK paragraph alone on its first line and breaks the CJK run without kinsoku; lowKashida and thaiDistribute CJK lines push a line-end U+3002 to the next line where the renderer hangs it. Thai cluster spreading of thaiDistribute, kashida elongation of Arabic words, East Asian automatic spacing on these lines and the mode-15 interword compression of both, distribute, lowKashida and thaiDistribute lines keep their separate models.',
});

export const WORD_RTL_LEFT_SOFT_FIT = defineCompatibilityRule({
  id: 'word-rtl-left-soft-fit',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/kashida-line-fit.test.ts#fits a soft-wrapped left-aligned RTL paragraph with $name by its visible prefix',
  },
  description: 'Observed soft-wrap admission for horizontal baseRTL paragraphs with literal left alignment and resolved wordWrap enabled in compatibility profiles absent, 11, 12, 14 and 15. The existing visible-prefix numerical primitive excludes collapsible source U+0020 from prospective bounded word fitting, including the final source candidate, while keeping the charged acquired advance and complete source range. Arabic and Latin run-direction counterclasses establish that paragraph bidi and run rtl are independent; scripts and fonts are evidence, not selectors. This arithmetic is not independently measured physical-prefix or face authority and does not establish Word equivalence: the measured Latin partition residual remains, and unchanged natural RTL paint can place visible Latin text outside a narrowed band; this admission rule grants no containment or new hanging adjustment. The capability is pass-local and bounded-layout only, outside tabs, fixed/atomic regions and unknown flattened direction provenance. It does not change paragraph-final paint, intrinsic layout, other alignments, justified reserves/allocation or the separate unmarked Arabic justified source-space rule.',
});

/** Line-fit policy of {@link WORD_KASHIDA_LINE_FIT}: a line's text may fill
 * `fraction` of its available width, less `reserveShare` of the width from the
 * line's start edge to its far edge (the float edge on a narrowed side, else
 * the column or cell edge past the trailing indent); undefined keeps the whole
 * width. */
export function wordJustifiedLineFit(
  alignment: string | null | undefined,
  compatibilityMode: number | undefined,
  baseRtl: boolean,
  wordWrap: boolean,
  trailingIndentPt = 0,
): { readonly fraction: number; readonly reserveShare: number; readonly rtlVisiblePrefix: boolean } | undefined {
  // Before mode 15 a negative trailing indent, from -5pt, removes the reserve
  // (blk-0898-floats-negind: 76 of 76 lines; blk-0898-kfit9).
  const reserve = compatibilityMode !== 15 && trailingIndentPt < 0 ? 0
    : alignment === 'mediumKashida' ? 0.1 : alignment === 'highKashida' ? 0.2 : 0;
  const rtlVisiblePrefix = wordRtlJustifiedVisibleFitApplies(alignment, baseRtl, wordWrap);
  if (reserve === 0 && !rtlVisiblePrefix) return undefined;
  // From mode 15 the reserve is a share of the line's own width; before it, a
  // share of the width from the start edge to the far edge.
  return compatibilityMode === 15
    ? { fraction: 1 - reserve, reserveShare: 0, rtlVisiblePrefix }
    : { fraction: 1, reserveShare: reserve, rtlVisiblePrefix };
}

/** Observed Word 16.0.20527 compatibility, not an ECMA separator-charge rule.
 * Paragraph bidi (§17.3.1.6) and resolved run rtl/cs (§17.3.2.30/7) are distinct.
 * Mode15 homogeneous Arabic-bearing runs (including a same-run Latin label)
 * in baseRTL medium/highKashida charge the actual soft-edge U+0020. Sole-run
 * rtl=true countercontrols keep visible fitting; absent and explicit false
 * agree. Medium-line separator6/8/8.5/12pt controls cross the existing fit
 * band at the full measured advance. Both reserves remain unchanged.
 * Arial12 establishes the class; face names are not admission predicates.
 * Thirty source-bound accepted/next boundaries across three indent classes
 * agree. Non-Arabic counterexamples require the Arabic-bearing gate. Other
 * modes, mixed formatting/directions, fields/tabs/generated content and
 * vertical contexts retain their existing policy. This never changes final
 * line hanging or paint, and never remeasures source substrings. */
export function wordUnmarkedArabicKashidaChargesSourceSpace(
  paragraph: ParagraphLayoutSource,
  compatibilityMode: number | undefined,
  baseRtl: boolean,
  wordWrap: boolean,
  horizontal: boolean,
): boolean {
  if (compatibilityMode !== 15 || !baseRtl || !wordWrap
    || (paragraph.alignment !== 'mediumKashida' && paragraph.alignment !== 'highKashida')
    || !horizontal
    || (paragraph.numbering !== null && paragraph.numbering !== undefined)
    || paragraph.runs.length !== 1 || paragraph.complexFieldBoundaries !== undefined) return false;
  const run = paragraph.runs[0]!;
  if (run.type !== 'text' || run.rtl === true || run.cs === true
    || run.ruby || run.noteRef || run.fitTextVal !== undefined
    || ('optionalHyphen' in run && run.optionalHyphen)
    || ('noteSeparatorCharacter' in run && run.noteSeparatorCharacter)
    || /[\t\r\n\u00ad\ufffc]/u.test(run.text)) return false;
  return /(?=\p{Script=Arabic})\p{Letter}/u.test(run.text);
}

export const WORD_RTL_JUSTIFIED_VISIBLE_FIT = defineCompatibilityRule({
  id: 'word-rtl-justified-visible-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-kfit6',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A justified right-to-left paragraph (§17.3.1.6 bidi) with jc both, mediumKashida or highKashida and wordWrap on (§17.3.1.45; with it off nothing hangs and the space is charged) fits each line by its visible prefix: the U+0020 that ends the line does not count against the width, as on left-to-right lines (WORD_COLLAPSIBLE_LINE_EDGE_VISIBLE_FIT), replacing the library policy that charged it on right-to-left lines. The space then hangs: a justified line spreads its slack over the visible text, which fills the band, and the space stands outside it on the side it is drawn (left of right-to-left text; right of a Latin run drawn rightmost, past the start edge); an unjustified final line keeps the space inside the band while it fits there and otherwise starts at the band; line-end spaces split across runs are resolved run by run, each hanging only when nothing but other line-end spaces lies between it and the line\'s edge on its drawn side; a space drawn between segments cannot hang, so any overflow it leaves is absorbed into it and no visible glyph leaves the band (library policy for cases no control measured). Measured: the kfit6 Latin bidi lines match Word within 0.8pt per character for both and highKashida and 1.8pt for mediumKashida (final lines included; about 3pt before), and the justified kfit5 Arabic lines start within 0.4pt of Word at the left margin (3.3pt before). The WORD_KASHIDA_LINE_FIT reserve applies on top, taking w:left / w:start as the start indent. Evidence: synthetic Windows Word 16.0.20527 controls blk-0898-kfit6-mode{none,14,15}, Arial 12pt, both / mediumKashida / highKashida, each with no indent, a 113.4pt w:right and a 113.4pt w:left: Latin text in bidi paragraphs, Arabic rtl runs in left-to-right paragraphs, and Arabic in bidi paragraphs. With exact Arial advances, the Latin bidi mediumKashida paragraph with the start indent admits fit widths [330.9, 344.2) by visible width but [334.2, 347.5) with the line-end space charged, against the 331.7pt the reserve gives, before mode 15 and in mode 15 alike; only the visible fit is consistent. Arabic both paragraphs with either indent (no kashida) and the blk-0898-kfit5 Arabic medium/high paragraphs before mode 15 break as Word does only with the visible fit; with it every bidi paragraph of these controls matches before mode 15. The later source-owned mode-15 rule wordUnmarkedArabicKashidaChargesSourceSpace supersedes visible admission only for its homogeneous unmarked Arabic-bearing baseRTL medium/high class; controlled run-RTL and separator-size counterexamples establish that boundary. Explicit RTL Arabic, non-Arabic and mixed-source classes retain this visible-prefix rule. Whole-source identity and embedded-font cmap audits show the historical explicit-RTL and LTR Arabic character-count residuals were PDF extraction artifacts; those counterclasses have no visible source-character wrap mismatch and retain their existing policy. Mode-15 interword compression of right-to-left both lines (Word fits Latin bidi lines past the width) stays outside WORD_JUSTIFIED_INTERWORD_COMPRESSION, which keeps right-to-left lines out. Unmeasured, so keeping the charged line-end space: lowKashida, distribute and thaiDistribute right-to-left paragraphs, and left, centre and right aligned right-to-left lines.',
});

/** Scope of {@link WORD_RTL_JUSTIFIED_VISIBLE_FIT}: right-to-left paragraphs
 * of the measured justification types whose line-end spaces can hang
 * (wordWrap on, §17.3.1.45). The single gate of fitting and line planning. */
export function wordRtlJustifiedVisibleFitApplies(
  alignment: string | null | undefined,
  baseRtl: boolean,
  wordWrap: boolean,
): boolean {
  return baseRtl && wordWrap && (alignment === 'both' || alignment === 'justify'
    || alignment === 'mediumKashida' || alignment === 'highKashida');
}

/** Whether a soft-wrapped justified line opens no inter-CJK gaps, so a
 * space-free CJK line keeps its natural width ({@link WORD_KASHIDA_LINE_FIT}). */
export function wordSoftLineJustifiesAtSpacesOnly(alignment: string | null | undefined): boolean {
  return alignment === 'lowKashida' || alignment === 'mediumKashida'
    || alignment === 'highKashida' || alignment === 'thaiDistribute';
}

/** Scope of {@link WORD_LINE_END_SPACE_HANGS}. */
export function wordLineEndSpaceHangs(bidi: boolean, wordWrap: boolean): boolean {
  return !bidi && wordWrap;
}

/** Visible-prefix projection of {@link WORD_COLLAPSIBLE_LINE_EDGE_VISIBLE_FIT}. */
export function wordVisiblePrefixFitWidthPx(
  widthPx: number,
  trailingSpacePx: number,
  baseRtl: boolean,
): number {
  return widthPx - (baseRtl ? 0 : trailingSpacePx);
}

export const WORD_JUSTIFIED_INTERWORD_COMPRESSION = defineCompatibilityRule({
  id: 'word-justified-interword-compression',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'linefit-justified-interword-compression-matrix',
    application: 'Microsoft Word',
    version: '16.113',
    platform: 'macOS 27.0',
  },
  description: 'Word 16.113 mode-15 both/justify LTR horizontal bounded lines without a character grid fit by their visible width. A U+0020 run supplies gaps only when its nearest non-space cells on both sides are text glyphs; every standalone consecutive space counts, NBSP does not. Grapheme clusters are atomic: a space with a combining extension remains fixed, including source seams. Tabs, inline objects, ruby, fixed-width fitText cells and line edges keep adjacent spaces fixed. The unmeasured fitText interaction preserves its existing fixed pitch and atomic wrapping (§17.3.2.14). Run face, size, fields, note references, symbol glyphs, soft/no-break hyphens and zero-width text do not gate eligibility. WORD_EAST_ASIAN_AUTO_SPACE gaps between text cells join this family with their natural width. With overflow C and natural opportunity sum S, accept C <= S/4 and C/(S + candidate line-end space width) <= 0.5 * E/Sprime, where E is the expansion without the candidate and Sprime includes that alternative line-end space. Fit, retained paint and justification use the breaker advances; kerning follows WORD_KERN_THRESHOLD_AUTHORITY. Paint allocates positive or negative slack proportionally: delta_i = slack*w_i/S; final lines receive accepted compression too. Ideographic/CJK/SEA expansion has no proportional evidence and keeps its separate unweighted opportunity family, except that a U+3000 whose nearest non-space cells on both sides are not CJK (Latin words separated by ideographic spaces) is no opportunity: Windows Word 16.0.20527 keeps it at its natural 12pt in mode 15 (blk-0898-u3000-mode15; mode unset and 11-14 expand it, 12 to 12.4pt) and leaves such a final line at its natural width, while a U+3000 next to CJK keeps its opportunity (blk-0898-u3cjk-mode15: CJK and mixed CJK/Latin lines with U+3000 separators within 2.2pt, 12pt before); lines with right-to-left text keep every opportunity, and a supplementary ideograph counts as CJK; sparse contextual space acquisition avoids shaping full overlong words during candidate fitting. Evidence: 240 Word controls bracket zero/one/many gaps, repeated spaces, separator runs, NBSP, mixed sizes/faces, fields, notes, symbols, hyphens, tabs, first indents, drop caps, floats, final lines, ruby and one/two inline objects of 0.5/1/2 em. Mixed-width second-boundary brackets agree. Independent 734- and 382-probe remeasurements preserve all uniform arithmetic verdicts; these overlapping sets substitute for the unavailable historical 758-set. Normative ST_Jc (§17.18.44) specifies inter-word justification, not this observed arithmetic. The final positive-threshold sweep agrees on 63/66 admissions and all 68 kerning-switch observations. Three Times New Roman admissions remain mismatches: K1-tnr-ot0-3, K1-tnr-ot1-2 and K1-tnr-ot1-3. Native Canvas measures both the complete source and its segmented prefix at 205.83984375pt, while the pinned-font HarfBuzz authority measures 204.5390625pt; no source-seam correction or empirical amount bridges that native shaping difference. A matched K1-tnr-ot0-2/K1-tnr-ot1-2 pair at 201.45pt has identical measured inputs (C=4.38984375pt, S=13.5pt, E=57.1248046875pt, Sprime=13.5pt, separator=4.5pt), yet Word rejects flag-off and accepts flag-on. The quarter bound rejects both in the library: flag-on remains a known mismatch. Together with corpus paragraph [213], PDF lines 376/377, and paragraph [843], PDF lines 1433/1434 (four lines), and the earlier flag-dependent observation, this is an accepted unresolved Word predicate, not evidence for a flag guard or a fitted coefficient; a 44-control investigation did not establish a general suppression rule (15/44 decision-balance and 30/44 grid-guard disagreements). Corpus paragraph [577], PDF lines 966/967, is left-aligned mode 14: the existing OpenType gate disables space fitting after the kerning correction. This accepted separate-mode limitation belongs to issue #1660/#1702 and its owner; this mode-15 rule does not extend into it. Font-table alternate-name/substitute-font selection differences are outside same-font fidelity acceptance. Direct letter seams and unequal typography remain shaping evidence gaps, with previous boundaries retained. Object/text origins also show a slack-independent export residual, which is not corrected with a fitted advance. The distribute/lowKashida extension has the bounded evidence recorded beside the admission gate below. Other alignments, modes, grids, RTL/vertical and unbounded widths retain their previous policy. Windows Word (Microsoft 365, ExportAsFixedFormat) controls of a justified 16pt bold Times New Roman line at 470/467.8/467pt in mode 15, each exported with and without lineWrapLikeWord6 (six controls of the IEEE line-fit matrix), give identical line ends and identical compressed gaps in every pair, so that switch does not exclude a line from this rule; the renderer reproduces all six.',
});

/** Gate governed by {@link WORD_JUSTIFIED_INTERWORD_COMPRESSION}.
 * ECMA-376 §17.18.44 specifies justification classes, not this fit arithmetic.
 * Windows Word 16.0.20527 also admits the same near-overflow word under
 * distribute/lowKashida as under both: Courier New 12pt LTR start-tab final
 * cells and end-/center-tab wrapped cells in a 481.9pt band. All three classes
 * have identical first-line source glyph origins/advances; modes unset/14
 * instead move that word down. This bounded cross-alignment observation reuses
 * the existing proportional comparison and quarter-space cap, without claiming
 * that their thresholds were independently swept for the two new alignments.
 * The added alignments enable admission/contraction only; uncompressed
 * expansion retains its existing projection until bidi gaps have source-owned
 * physical edges. Paragraph acquisition additionally bounds the new alignments
 * to base-LTR content without strong RTL text or authored RTL marks; the scalar
 * gap model does not own mixed/RTL source edges. Existing both/justify behavior
 * and the grid, width and shaping exclusions remain unchanged.
 */
export function wordJustifiedInterwordCompressionApplies(
  alignment: string | null | undefined,
  compatibilityMode: number | undefined,
): boolean {
  return (alignment === 'both' || alignment === 'justify'
    || alignment === 'distribute' || alignment === 'lowKashida')
    && compatibilityMode === 15;
}

/** The existing quarter-space cap, shared with conservative refusal proofs.
 * An upper bound on S can certify refusal only; it cannot admit a line. */
export function wordJustifiedCompressionExceedsCap(overflow: number, naturalGapSum: number): boolean {
  return overflow > naturalGapSum / 4;
}

/** Observed proportional comparison. All lengths share one unit. Line-end
 * separators participate in the comparison but never receive paint slack. */
export function wordJustifiedInterwordCompressionFactor(input: Readonly<{
  overflow: number;
  naturalGapSum: number;
  candidateLineEndSeparator: number;
  previousOpportunitySum: number;
  expansionWithoutCandidate: number;
}>): number | undefined {
  const { overflow: C, naturalGapSum: S, previousOpportunitySum: Sprime,
    expansionWithoutCandidate: E, candidateLineEndSeparator: t } = input;
  if (!(C > 0) || !(S > 0) || !(Sprime > 0) || wordJustifiedCompressionExceedsCap(C, S)) return undefined;
  return C / (S + t) <= 0.5 * E / Sprime ? C / S : undefined;
}

export const WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS = defineCompatibilityRule({
  id: 'word-overflow-punctuation-language-sets',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.56',
  },
  description: 'Apply the language-specific punctuation sets documented for Word in [MS-OE376] §2.1.56 to Chinese, Japanese, and Korean language runs, and let overflowPunct override kinsoku when both rules affect the same character. When an effective East Asian language is absent, content that actually selects the East Asian script path retains the union as a bounded fallback.',
});

export const WORD_JUSTIFICATION_PUNCTUATION_HANG = defineCompatibilityRule({
  id: 'word-justification-punctuation-hang',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-cjk-kin',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Which paragraphs let an overflowPunct (§17.3.1.21) closing mark hang past the line end, and how many. lowKashida, mediumKashida, highKashida and thaiDistribute paragraphs never hang one: the mark stays inside the line and kinsoku moves the character before it down with it (from compatibility mode 14; before mode 14 the kashida types instead break their CJK text without kinsoku, a separate unimplemented rule, and thaiDistribute already pushes). At most one mark hangs: when the character after a hangable mark may not start a line either (。」), nothing hangs and kinsoku moves the character before the pair down. Evidence: synthetic Windows Word 16.0.20527 controls blk-0898-cjk-kin-mode{none,11,12,14,15}, MS Mincho 12pt at the full 481.9pt width, every jc (left, both, distribute, low/medium/highKashida, thaiDistribute): 。 that fits at the last slot; 。, 、, 」, 々 and 。」 one slot past the margin; the same with eastAsia ja-JP, with overflowPunct off and with kinsoku off; and after a Latin label whose remainder leaves 4.5-6.3pt. Table cells, headers and footnotes: blk-0898-cjk-kincell-mode{none,11,12,14,15} repeat every row above inside a one-cell lrTb table with cell margins 0, with left/right margins 108 twips, and nested one level (each cell content width equal to the body width), and eight O1/O5 rows (every jc) in the default header and in a footnote. The header and footnote rows break exactly as the body rows in every mode. In every cell variant and every mode a table cell never hangs a mark: every row that hangs in the body (O1, O2, O3, O1j and PO1-PO3 under left, both and distribute; left only before mode 15) pushes the character before the mark instead, and with kinsoku off (N2) the mark starts the next line; every other row (kashida and thaiDistribute, 々, 。」, fitting marks) breaks exactly as in the body. Scope (admitted positively, everything else keeps the previous overflowPunct behaviour): CJK marks in ordinary (non-frame), left-to-right paragraphs with overflowPunct on and a measured alignment (left, both, distribute and the kashida and Thai types, with their Strict spellings start and justify; additionally center/right/end in the measured kinsoku-on version profiles below), in an lrTb section without a document grid, at the top level of the main body, a header or a footnote, or in a table cell of the main body whose authored text direction (and every enclosing cell) is horizontal, where no mark hangs. Footers, endnotes, note separators, text boxes, frames, cells of headers and notes, rotated or other authored-vertical cells, upright tables of vertical sections, centred and right-aligned paragraphs outside those kinsoku-on profiles, Latin and complex-script marks, right-to-left, vertical and grid paragraphs are unmeasured and keep their hang. The character after the mark is taken across run seams. Every kashida and thaiDistribute row from mode 14 (and thaiDistribute in every mode) pushes (line 1 one character short of the capacity, line 2 starting with the character and the mark); both and distribute hang 。、」 in every mode and push 。」 and 々. Additional kinsoku-on controls in modes unset, 11, 12, 14 and 15 establish left/center/right and Strict start/end: nonjustified body text hangs one mark before 15 and pushes it in 15; horizontal table cells never hang. Explicit compressPunctuation and doNotCompress countercontrols retain this version/alignment result, so character-spacing compression is not its selector. This extension retains the previous paragraph-off and document-off behavior, including the separate unresolved paragraph-off line-breaking limitation (#2079), and does not extrapolate unmeasured modes.',
});

export const WORD_OVERFLOW_PUNCTUATION_LATIN_PARENT_RUN = defineCompatibilityRule({
  id: 'word-overflow-punctuation-latin-parent-run',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'latin-ascii-punctuation-advance-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Although [MS-OE376] §2.1.56 presents concrete punctuation sets by CJK language, Office-produced Latin boundary controls admit `.` and `,` beyond the ordinary word-fit extent. Controls at 9 and 14 points in Calibri and Arial instead wrap `)` and `}` with the word at the measured advance boundary; 10-point controls also wrap `!`, `%`, `:`, `;`, `>`, `?`, and `]`. Complex-script segments with no explicit RTL-primary bidi language retain their separate observed fallback; an explicit `ar-SA` bidi language is its counterexample.',
});

export const WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE = defineCompatibilityRule({
  id: 'word-full-width-character-spacing-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.562',
  },
  description: 'Interpret ST_CharacterSpacing as applying whitespace compression to full-width punctuation characters. This rule establishes only which characters are eligible; it does not define a universal compression amount.',
});

export const WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL = defineCompatibilityRule({
  id: 'word-japanese-punctuation-compression-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'japanese-fullwidth-punctuation-compression-cell',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed Japanese compatibility matrix, 、。 ，． and the closing forms 」』】）］｝ on a full ideographic-cell advance retain at least half of that cell. U+3017 and full-width !/? remain full-cell. A fontTable w:pitch value classifies the authored face for font selection; it is not a switch for document-level characterSpacingControl. Punctuation that the selected face already exposes on a smaller proportional advance is retained as measured rather than compressed a second time. Tight adjacent glyph ink can require a larger retained extent to prevent collision. This is an Office-observed compression amount, not a normative interpretation of ST_CharacterSpacing. Windows Word 16.0 build 16.0.20430 retest: Enabled and disabled outputs have identical advances in these unconstrained lines. MS Mincho and Meiryo retain 12pt cells; MS PMincho retains natural 7.92pt comma/period and 6pt closing-bracket advances, with 12pt lenticular/exclamation/question cells. No line-fit pressure is authored and no exact-source Mac PDF exists. These are natural-width counterexamples, not a measurement of the compressed half-cell minimum; do not report a Windows/Mac compression difference.',
});

export const WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY = defineCompatibilityRule({
  id: 'word-authored-character-spacing-pitch-priority',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'authored-character-spacing-punctuation-pitch',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When a run authors a positive w:spacing character pitch, Word preserves that expanded pitch instead of additionally applying the document-level punctuation whitespace compression. Omitted, zero, or overlapping run spacing leaves characterSpacingControl active.',
});

export const WORD_SOURCE_RUN_SPACE_SEQUENCE = defineCompatibilityRule({
  id: 'word-source-run-space-sequence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'source-run-space-sequence-wrap-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'At a source-run boundary, Word keeps a space-only continuation attached when the preceding run already ends in a space. A single leading space in a distinct run without a preceding space remains a break opportunity. This isolates source-boundary compatibility from the ordinary UAX #14 LB7 handling within one authored run.',
});

export const WORD_BALANCED_CONSECUTIVE_SPACE_CELL = defineCompatibilityRule({
  id: 'word-balanced-consecutive-space-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-space-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With ECMA-376 §17.15.3.3 balanceSingleByteDoubleByteWidth enabled, Word retains one ordinary inter-word U+0020 at its proportional natural advance, while a sequence of two or more authored U+0020 spaces advances each space by half of the selected East-Asian ideographic cell. The observed matrix covers one, two, four, and eight spaces; same-run and source-run boundaries; proportional and fixed-pitch faces; linesAndChars with negative/zero charSpace; and a line-only grid.',
});

/** Compatibility projection governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. */
export function wordBalancedConsecutiveSpaceCellApplies(spaceCount: number): boolean {
  return Number.isInteger(spaceCount) && spaceCount >= 2;
}

/** Evidence-bounded grid scope governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. `snapToChars` has a separate
 * Microsoft-documented block/cell allocator and is outside this observation. */
export function wordBalancedSpaceCellAdjustmentApplies(
  gridType: string | null | undefined,
): boolean {
  return gridType !== 'snapToChars';
}

export const WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA = defineCompatibilityRule({
  id: 'word-balanced-lines-and-chars-grid-delta',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-grid-observation-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With balanceSingleByteDoubleByteWidth enabled on linesAndChars, Word applies half of the authored charSpace delta to ASCII SBCS text and to U+0020/U+3000 space characters, while applying the full delta to CJK ideographs and full-width ASCII forms. The Word-output evidence covers ASCII digits, letters, punctuation, spaces, CJK, full-width ASCII, mixed text, proportional/fixed-pitch faces, negative/zero/positive charSpace, and line-only controls. Non-ASCII high-ANSI and complex-script text are outside the observed matrix and retain the preexisting grid behavior.',
});

export const WORD_LATIN_INTERWORD_XAVG_FLOOR = defineCompatibilityRule({
  id: 'word-latin-interword-xavg-floor',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'latin-space-os2-xavg-one-twip-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For left-aligned homogeneous Latin words using one U+0020 separator and characterSpacingControl=compressPunctuation, Word reduces each natural inter-word space only as far as half the selected static face\'s positive OS/2 xAvgCharWidth. First-fit and beta-origin controls varied only U+0020 hmtx, then only xAvg, with fixed outlines and non-space advances at 8/11/16pt; a distinct Carlito outline provided a counterexample where natural U+0020 was narrower than the floor. Two-gap controls showed equal required-deficit allocation. A one-twip linesAndChars control preserved the xAvg-dependent deficit while moving both control and natural boundaries by the separate character-grid pitch. [MS-OE376] §2.1.472 documents lineWrapLikeWord6 as an explicit uncompressed-fit override. Mixed faces, explicit run spacing or scaling, justification, and snapToChars remain outside the observed scope. Windows Word (Microsoft 365, ExportAsFixedFormat) controls of one 16pt bold Times New Roman line with w:kern 28 at text widths 470/467.8/467pt, jc left/center/right/both, compatibility modes 11 and 15 with and without lineWrapLikeWord6 (48 cases): in mode 11 left, centered and right lines compress alike, with or without lineWrapLikeWord6 and despite the explicit kerning threshold (the 470pt line fits, the narrower ones do not, as the floor predicts); justified mode-11 lines never compress; mode-15 lines with the kerning threshold keep natural spaces with or without lineWrapLikeWord6. So below mode 15 neither lineWrapLikeWord6 nor the kerning threshold gates the floor (lineWrapLikeWord6 measured in mode 11, kerning in modes 11 and 14); enabled OpenType features still exclude it outside mode 14, where alone their independence is measured; mode 15 and an omitted mode keep the earlier lineWrapLikeWord6/OpenType/kerning exclusions. The one 467.8pt boundary case lies within 0.01pt per space of the floor computed from the macOS face (xAvgCharWidth 874); the Windows face is not available to resolve it.',
});

/** Document and run gate of {@link WORD_LATIN_INTERWORD_XAVG_FLOOR} beyond
 * the line-content conditions the breaker checks. */
export function wordLatinInterwordXavgFloorApplies(input: Readonly<{
  compatibilityMode: number | undefined;
  lineWrapLikeWord6: boolean | undefined;
  enableOpenTypeFeatures: boolean | undefined;
  hasExplicitKerningThreshold: boolean;
}>): boolean {
  if (input.compatibilityMode !== undefined && input.compatibilityMode < 15) {
    // OpenType independence is measured in mode 14 only.
    return input.compatibilityMode === 14 || input.enableOpenTypeFeatures !== true;
  }
  return input.lineWrapLikeWord6 !== true
    && input.enableOpenTypeFeatures !== true
    && !input.hasExplicitKerningThreshold;
}

export const WORD_EAST_ASIAN_AUTO_SPACE = defineCompatibilityRule({
  id: 'word-east-asian-auto-space',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'east-asian-latin-autospace-1851-matrix',
    application: 'Microsoft Word',
    version: '16.0',
    platform: 'Windows (build 16.0.20430)',
  },
  description: 'ECMA-376 §17.3.1.2 autoSpaceDE and §17.3.1.3 autoSpaceDN (both on unless set false) ask for adjusted spacing between East Asian text and Latin text or numbers without an amount. Issue #1851 Windows controls (5376 single-line pairs: 14 adjacency strings, each flag on/off, MS Mincho+Century, Yu Mincho, MS PMincho+Times New Roman and MS Gothic+Arial at 10.5/20pt, no grid, lines and linesAndChars charSpace 0, modes 14/15 with doNotCompress/compressPunctuation) place a quarter em on both sides wherever a letter of the Han, Hiragana, Katakana (including half-width) or Hangul script meets a Latin letter (DE) or an ASCII digit (DN), identically in every mode, spacing control and grid; each flag governs its own class only. East Asian punctuation (、。「」), ASCII parentheses, symbols and punctuation, an authored U+0020 and full-width Latin letters are neutral; a Latin letter set in the East Asian face is spaced, so the classes follow the characters, as §17.3.1.2 states, not the face or run. The amount is a quarter of the East Asian side\'s font size: the space glyph Word exports for a Latin-to-East-Asian gap belongs to the East Asian run (the controls pair equal sizes). The gap exists only between partners on one line: wrapped lines never start with it, and a left-aligned line ends before a Latin word with less room left than the gap (2.16pt). Line fitting: mode-15 doNotCompress lines keep the natural gap (E-mail wrap sweep 70-99pt in 1pt steps, two face pairs, three grids, deficits up to 0.15pt PDF quantization); under the WORD_COMPRESSED_SPACE_LINE_FIT gate the gap shrinks with the line\'s U+0020 down to an eighth em (accepted deficits up to 1.40pt, rejected from 1.55pt at 10.5pt after the 0.15pt export offset, within 0.12pt position quantization), each gap giving up the same share of its room, which the issue #1660 Mac fixed-cell controls with four U+0020 and one gap (92 widths, modes 14/15, one-twip mode-14 sweep) also reproduce; the gap keeps shrinking on a linesAndChars grid and with balanceSingleByteDoubleByteWidth (mode-14 linesAndChars charSpace 0 controls; a Word-exported linesAndChars document reaches 2.15pt gaps at 11pt), whose U+0020 keep that rule\'s grid and width-balance exclusions; mode-15 justified lines admit overflow with the gaps in the WORD_JUSTIFIED_INTERWORD_COMPRESSION family (every admitted line and every rejected next character of six measured lines agree; without the gaps none would be admitted), and a final ideograph of a run is admitted by the same comparison as a whole unit, while the unit of a Latin word ends at its gap seam. Justification (both/distribute) gives positive slack to the gaps first, in proportion to their room, until each reaches half an em, and only the rest to the ordinary opportunities (ideographs keep their pitch while the gaps grow 2.77-3.78pt; linesAndChars lines reach 5.30-5.56pt with the remainder spread over the ideographs). Limits: snapToChars grids allocate whole cells by script block (WORD_SNAP_TO_CHARS_SCRIPT_BLOCK_ALLOCATION) and apply no gap; a nonzero linesAndChars charSpace, vertical and right-to-left text, a size difference across the boundary and Latin letters outside ASCII follow the same rule unmeasured; mode-15 linesAndChars justified lines, which Word also compresses, keep the existing grid exclusion of WORD_JUSTIFIED_INTERWORD_COMPRESSION; a positive authored character pitch keeps the gap at its natural width (WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY).',
});

/** Gap of {@link WORD_EAST_ASIAN_AUTO_SPACE} for an East Asian character of
 * `eastAsianFontSize` points. */
export function wordEastAsianAutoSpacePt(eastAsianFontSize: number): number {
  return eastAsianFontSize / 4;
}

/** Narrowest a {@link WORD_EAST_ASIAN_AUTO_SPACE} gap of `naturalPt` becomes
 * under {@link WORD_COMPRESSED_SPACE_LINE_FIT}. */
export function wordCompressedAutoSpaceFloorPt(naturalPt: number): number {
  return naturalPt / 2;
}

/** Widest a {@link WORD_EAST_ASIAN_AUTO_SPACE} gap grows under justification
 * before the remaining slack reaches the ordinary opportunities. */
export function wordEastAsianAutoSpaceJustifiedMaximumPt(naturalPt: number): number {
  return naturalPt * 2;
}

export const WORD_COMPRESSED_SPACE_LINE_FIT = defineCompatibilityRule({
  id: 'word-compressed-space-line-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'mixed-script-space-fit-1660-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'Issue #1660 controls (1773 Word-exported fixed-cell lines: 576 coarse, 595 fine one-twip, 602 hypothesis-targeted) establish U+0020 fitting on lines mixing East Asian and Latin text, which WORD_LATIN_INTERWORD_XAVG_FLOOR (homogeneous Latin lines, unchanged) never compresses. In compatibility modes 12 and 14 with characterSpacingControl compressPunctuation or compressPunctuationAndJapaneseKana, every U+0020 on such a line shrinks by the same amount (single, consecutive and source-run-split spaces alike) while ideographs, kana and other glyphs keep their natural advances; enableOpenTypeFeatures and explicit w:kern thresholds do not gate it, and left, both and distribute alignment behave alike. Each space keeps at least min(xAvgCharWidth / 2, font size / 4) of its own face: Arial and Times New Roman spaces follow half their OS/2 xAvgCharWidth, while BIZ UDGothic, Meiryo and Yu Gothic (xAvgCharWidth 0.84-0.96 em) and BIZ derivatives with only post.isFixedPitch or only the far-east code-page bits changed all stop at a quarter em, rejecting fixed-pitch, code-page and weighted-lowercase-average selectors. A candidate whose last character before trailing closing punctuation is East Asian is admitted only while its natural overflow, measured without that trailing punctuation, is at most half the font size: a final ideograph or half-width kana admits 0.5 em of overflow at 8.5/10.5/16 pt, a final kana followed by a full-width closing parenthesis keeps that limit before the parenthesis (which keeps half its cell at the line end), and a final Latin word has no limit beyond the space floors. characterSpacingControl omitted or doNotCompress and compatibility mode 15 keep natural spaces on mixed lines; lineWrapLikeWord6 ([MS-OE376] §2.1.472) and an omitted compatibility mode are outside the projection. Mode-15 justified mixed lines (Word admits 2.60pt of overflow over four 4.25pt spaces but not 2.65pt) are unexplained and keep natural fitting. No control contains U+3000, whose line-end hanging (WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE) this observation does not define, so paragraphs holding U+3000 are ineligible for this compression rule; formatting-only source seams are still normalized before ordinary fitting. Automatic East Asian/Latin gaps (WORD_EAST_ASIAN_AUTO_SPACE) on such a line shrink with its spaces down to an eighth em, every gap giving up the same share of its room. One measured input the renderer does not reproduce also keeps it, paragraph-wide and read on the joined text: a compressible closing mark directly followed by U+0020 (Word keeps the full cell; the renderer compresses it under WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL). Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

/** Document gate of {@link WORD_COMPRESSED_SPACE_LINE_FIT}: an authored
 * compatibility mode below 15 (12 and 14 measured) with a compressing
 * characterSpacingControl. The line breaker further restricts it to lines
 * holding East Asian text. */
export function wordCompressedSpaceLineFitApplies(
  compatibilityMode: number | undefined,
  characterSpacingControl: string | undefined,
): boolean {
  return compatibilityMode !== undefined && compatibilityMode < 15
    && (characterSpacingControl === 'compressPunctuation'
      || characterSpacingControl === 'compressPunctuationAndJapaneseKana');
}

/** Per-space minimum advance of {@link WORD_COMPRESSED_SPACE_LINE_FIT} and
 * {@link WORD_LATIN_INTERWORD_XAVG_FLOOR}, in the caller's unit. */
export function wordCompressedSpaceFloor(fontSize: number, averageWidthRatio: number): number {
  return Math.min((fontSize * averageWidthRatio) / 2, fontSize / 4);
}

/** Natural-overflow limit of {@link WORD_COMPRESSED_SPACE_LINE_FIT} for a
 * candidate ending (before trailing closing punctuation) in an East Asian
 * character; undefined means no limit beyond the space floors. */
export function wordCompressedSpaceEastAsianOverflowLimit(
  lastCharacterIsEastAsian: boolean,
  fontSize: number,
): number | undefined {
  return lastCharacterIsEastAsian ? fontSize / 2 : undefined;
}

export const WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE = defineCompatibilityRule({
  id: 'word-ideographic-space-line-end-allowance',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ideographic-space-line-end-count-and-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word keeps a single U+3000 immediately following visible East-Asian text on that line when the visible glyph is force-fitted into a narrow table cell. A paragraph-final sequence of two or more U+3000 characters remains authored width-bearing content and may form blank continuation lines. The observed matrix covers single and trailing multiple spaces, linesAndChars with negative/positive charSpace, line-only grids, and snapToGrid opt-out.',
});

/** Compatibility projection governed by
 * {@link WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE}. */
export function wordIdeographicSpaceLineEndAllowanceCount(
  hasEastAsianVisiblePredecessor: boolean,
  consecutiveSpaceCount: number,
): 0 | 1 {
  return hasEastAsianVisiblePredecessor && consecutiveSpaceCount === 1 ? 1 : 0;
}

/** Compatibility projection governed by
 * {@link WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA}. Script-slot acquisition
 * has already separated ordinary East-Asian and ASCII SBCS text; the explicit
 * space branch retains Word's observed U+3000 exception without reclassifying
 * other East-Asian glyphs. Non-ASCII high-ANSI/complex text stays outside the
 * observed projection. */
export function wordBalancedLinesAndCharsGridDeltaFactor(
  text: string,
  script: 'ascii' | 'highAnsi' | 'eastAsia' | 'complexScript',
): 0.5 | 1 | undefined {
  if (script === 'complexScript') return undefined;
  const spaceOnly = text.length > 0 && [...text].every(
    (character) => character === ' ' || character === '\u3000',
  );
  if (spaceOnly) return 0.5;
  if (script === 'eastAsia') return 1;
  return [...text].every((character) => (character.codePointAt(0) ?? 0x80) <= 0x7f)
    ? 0.5
    : undefined;
}

/** Compatibility projection governed by
 * {@link WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL}. */
export function wordJapanesePunctuationRetainedExtentPt(input: Readonly<{
  punctuationAdvancePt: number;
  punctuationInkEndPt: number;
  ideographicCellAdvancePt: number;
}>): number {
  const advancePt = Math.max(0, input.punctuationAdvancePt);
  const cellAdvancePt = Math.max(0, input.ideographicCellAdvancePt);
  if (advancePt < cellAdvancePt) return advancePt;
  return Math.min(
    advancePt,
    Math.max(
      0,
      input.punctuationInkEndPt,
      cellAdvancePt / 2,
    ),
  );
}

/** Compatibility projection governed by
 * {@link WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY}. */
export function wordDocumentCharacterCompressionApplies(
  authoredCharacterSpacingPt: number | undefined,
): boolean {
  return authoredCharacterSpacingPt === undefined || authoredCharacterSpacingPt <= 0;
}

/** Compatibility projection governed by {@link WORD_SOURCE_RUN_SPACE_SEQUENCE}. */
export function wordSourceRunSpaceContinuesSequence(
  previousText: string,
  currentText: string,
): boolean {
  return previousText.endsWith(' ') && currentText.startsWith(' ');
}

const WORD_OVERFLOW_PUNCTUATION = {
  ja: new Set([...',.’”、。」』】），．］｝｡､']),
  zhHans: new Set([...`!%),.:;>?]}¢°·ˇ’”‰′″℃∶、。〃〉》」』】〗〕〞﹚﹜﹞！＂％＇），．：；？］｝￠`]),
  zhHant: new Set([...`!),.:;?]}’”′、。〉》」』】〕〞﹚﹜﹞！），．：；？］｝`]),
  ko: new Set([...`!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠`]),
} as const;
const ALL_WORD_OVERFLOW_PUNCTUATION = new Set([
  ...WORD_OVERFLOW_PUNCTUATION.ja,
  ...WORD_OVERFLOW_PUNCTUATION.zhHans,
  ...WORD_OVERFLOW_PUNCTUATION.zhHant,
  ...WORD_OVERFLOW_PUNCTUATION.ko,
]);
const LATIN_WORD_OVERFLOW_PUNCTUATION = new Set(['.', ',']);

/** {@link WORD_JUSTIFICATION_PUNCTUATION_HANG} policy for a CJK closing mark
 * in its measured scope: at most one mark hangs ('one'), or none does. */
export type CjkClosingMarkHang = 'one' | 'none';

/** The policy for a paragraph in the measured class (see
 * StoryContext.closingMarkHangAdmitted in layout-context.ts) with a measured
 * alignment, 'none' throughout inside a table cell; `undefined` keeps the
 * unmeasured overflowPunct behaviour. `start`
 * is the Strict spelling of `left` and `justify` that of `both` (§17.18.44);
 * the class is left-to-right only, where they coincide. */
export function wordCjkClosingMarkHang(
  alignment: string | null | undefined,
  measuredScope: boolean,
  insideTableCell = false,
  paragraphKinsokuOff = false,
  compatibilityMode?: number,
  documentKinsokuEnabled = true,
): CjkClosingMarkHang | undefined {
  if (!measuredScope) return undefined;
  // §17.3.1.16 defines forbidden line edges, not this Office-version switch.
  // Observed ON controls: unset/11/12/14 hang one mark for nonjustified LTR
  // text; 15 pushes it with its preceding character. Center/right (Strict end)
  // and horizontal table cells were independently measured in those profiles.
  // Explicit punctuation-compression on/off does not select this rule.
  // Other modes, document-off and paragraph-off retain the exact old policy;
  // this does not implement paragraph-off line breaking (#2079).
  const measuredOnProfile = compatibilityMode === undefined
    || compatibilityMode === 11 || compatibilityMode === 12
    || compatibilityMode === 14 || compatibilityMode === 15;
  if (documentKinsokuEnabled && !paragraphKinsokuOff && measuredOnProfile
    && (alignment === 'left' || alignment === 'start' || alignment === 'center'
      || alignment === 'right' || alignment === 'end')) {
    return insideTableCell || compatibilityMode === 15 ? 'none' : 'one';
  }
  switch (alignment) {
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return 'none';
    case 'left':
    case 'start':
    case 'both':
    case 'justify':
    case 'distribute':
      // A table cell never hangs the mark (blk-0898-cjk-kincell).
      return insideTableCell ? 'none' : 'one';
    default:
      return undefined;
  }
}

/** Whether {@link WORD_JUSTIFICATION_PUNCTUATION_HANG} refuses to hang `mark`
 * given the character after it (`next` yields it on demand: undefined when
 * nothing, or no text, follows). Only a CJK mark in horizontal text is
 * judged; everything else keeps its hang, and `next` is read only when the
 * one-mark limit needs it. */
export function wordCjkClosingMarkHangRefused(
  policy: CjkClosingMarkHang | undefined,
  mark: string,
  next: () => string | undefined,
  kinsoku: Readonly<{ enabled: boolean; lineStartForbidden: ReadonlySet<number> }>,
  verticalRun: boolean,
): boolean {
  if (policy === undefined || verticalRun || !isCjkBreakChar(mark.codePointAt(0)!)) return false;
  if (policy === 'none') return true;
  if (!kinsoku.enabled) return false;
  const following = next();
  return following !== undefined && kinsoku.lineStartForbidden.has(following.codePointAt(0)!);
}

export const RTL_PRIMARY_SUBTAGS = new Set([
  'ar', 'fa', 'ur', 'he', 'iw', 'yi', 'ji', 'ps', 'sd', 'ug', 'dv', 'syr', 'ckb',
]);

/** Compatibility projection governed by
 * {@link WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS}. */
export function wordIsOverflowPunctuation(
  character: string,
  language: string | undefined,
  parentRunHasEastAsianText = false,
  parentRunHasLatinText = false,
  parentRunHasComplexScriptText = false,
  bidiLanguage?: string,
): boolean {
  const normalized = language?.toLowerCase();
  if (normalized?.startsWith('ja')) return WORD_OVERFLOW_PUNCTUATION.ja.has(character);
  if (normalized?.startsWith('ko')) return WORD_OVERFLOW_PUNCTUATION.ko.has(character);
  if (normalized?.startsWith('zh')) {
    return (/(?:^|-)(?:tw|hk|mo)(?:-|$)|hant/u.test(normalized)
      ? WORD_OVERFLOW_PUNCTUATION.zhHant
      : WORD_OVERFLOW_PUNCTUATION.zhHans).has(character);
  }
  // ECMA-376 §17.3.1.21 is script-neutral. Although [MS-OE376] §2.1.56
  // describes Word's concrete sets as CJK-language behavior, Office-produced
  // boundary controls hang `.` and `,` in Latin parent runs; 9/14pt Calibri
  // and Arial counterexamples wrap `)` and `}` at the normal measured advance
  // boundary. Complex-script production controls with an absent bidi tag
  // hang `.`, `:`, `)` and `>`; an explicit RTL-primary tag such as
  // `ar-SA` is the counterexample. A CJK run carrying an inherited
  // non-CJK language also uses the union, because its actual script route is
  // more authoritative than that inherited language.
  const bidiPrimary = bidiLanguage?.split('-')[0].toLowerCase();
  const explicitRtlBidi = bidiPrimary != null && RTL_PRIMARY_SUBTAGS.has(bidiPrimary);
  const observedComplexFallback = parentRunHasComplexScriptText && !explicitRtlBidi;
  if (parentRunHasEastAsianText || observedComplexFallback) {
    return ALL_WORD_OVERFLOW_PUNCTUATION.has(character);
  }
  return parentRunHasLatinText && LATIN_WORD_OVERFLOW_PUNCTUATION.has(character);
}

export const WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE = defineCompatibilityRule({
  id: 'word-ruby-paragraph-uniform-line-advance',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/paragraph-measure.test.ts#uses one uniform snapped advance for every line in a ruby paragraph',
  },
  description: 'Every line in a ruby-bearing paragraph uses the paragraph-wide maximum snapped line advance so its baseline rhythm remains uniform. Library policy, no Office measurement: uniform ruby paragraph line advance kept through a refactor (1bfaf434a); nothing measured.',
});

export const WORD_FIT_TEXT_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-fit-text-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/fit-text.test.ts#distributes (val − Σnatural)/(n−1) as the inter-character gap, no trailing gap',
  },
  description: 'Expand a multi-character fitText region to its authored width by distributing the residual evenly across interior character gaps. Evidence (observed Office behaviour): Word PDF of a corpus document in PR #918; [MS-OI29500] §2.1.78(d) supports ignoring authored spacing, not the even-gap formula.',
});

export const WORD_FIT_TEXT_OMITTED_ID_LINKING = defineCompatibilityRule({
  id: 'word-fit-text-omitted-id-linking',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.78(b)',
  },
  description: 'A fitText run without w:id takes id 0 and shares one fit region with contiguous fitText runs whose id is omitted or 0, where ECMA-376 §17.3.2.14 leaves an id-less run unlinked.',
});

export const WORD_CJK_BOTH_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-cjk-both-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#§17.18.44: fills a wrapped pure-CJK line via inter-CJK pitch (expansion default)',
  },
  description: 'Treat inter-CJK boundaries as eligible inter-word gaps when expanding a non-final both-justified line that contains no spaces. Evidence (observed Office behaviour): Word PDF of a corpus document in PR #483; this departs from the literal ECMA-376 §17.18.44 both.',
});

export const WORD_THAI_DISTRIBUTE_CLUSTER_POLICY = defineCompatibilityRule({
  id: 'word-thai-distribute-cluster-policy',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/thai-distribute.test.ts#fills non-final lines to the right margin under thaiDistribute',
  },
  description: 'Expand non-final thaiDistribute lines at Thai grapheme-cluster boundaries while retaining a natural-width final line. Classification: normative, ECMA-376 §17.18.44 thaiDistribute (inter-character distribution); the natural-width final line is library policy (PR #978).',
});

export const WORD_THAI_DISTRIBUTE_PROPORTIONAL_SPREAD = defineCompatibilityRule({
  id: 'word-thai-distribute-proportional-spread',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-cjk-thjust',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Before compatibility mode 15 a thaiDistribute line that is justified (soft-wrapped, or ended by a manual break; a piece ended by a page or column break is unmeasured and keeps the equal spread) and holds Thai text but no CJK spreads its slack in proportion to advance: every Thai grapheme cluster and every space widens by the same fraction of its own advance, the last cluster of the line excepted, while Latin letters keep their advance (a Latin word widens only at its spaces). Final lines stay natural (word-thai-distribute-cluster-policy). Evidence: synthetic Windows Word 16.0.20527 controls blk-0898-cjk-thjust-modenone and blk-0898-cjk-thdfin-modenone, Tahoma 12pt Thai at the full 481.9pt width, with and without th-TH: spaced Thai phrases (J2), spaced Thai without above or below marks (J3), a Times New Roman label before spaced Thai (J1) and Latin words between Thai phrases (J5), and a space-free Thai line ended by a manual break (D10). The share of each cluster and space is its advance times one line factor (residual within 0.3pt per character, increments quantised to about 0.12pt); an equal share per cluster misses by up to 1.7pt per gap on the same lines, and the label and Latin letters stay within 0.05pt of their natural advance. Re-exported after Thai (th-TH) was added as an Office editing language (Thai proofing tools not installed), the same controls place every glyph at the same position (0.01pt), so the rule does not depend on that setting. Rendered against those PDFs, characters land within 1.1pt (before: 2.2-3.9pt, 19.2pt on the break-ended line; worst character per row J1/J2/J3/J5 1.7/1.8/1.0/1.7pt against 2.6/3.6/2.2/3.9pt, D5 1.1 against 2.4pt, D10 0.6 against 19.2pt, every excess over 1.1pt at a SARA AM). Limitation: Word also widens the seam before THAI CHARACTER SARA AM, which shapes as one cluster with its base here, so ำ and the characters after it can sit up to 1.8pt off. Weights are the retained cluster advances (a cluster continued across a formatting seam weighs its whole advance), so nothing is shaped again. Scope: the measured line class only, horizontal left-to-right lines of text runs holding Thai and printable ASCII (U+0020-U+007E), where Thai meets other text only at a U+0020 space, without character spacing, scaling, a character grid, fitText or ruby, and in a paragraph without bidi; every other line keeps the equal per-gap spread of word-thai-distribute-cluster-policy unchanged, because Word was not measured there: Thai directly against Latin letters, digits or punctuation, CJK (supplementary Han and fixed-pitch fitText regions included), Lao, Khmer, Myanmar and other scripts, tabs, inline objects, and lines whose gap cuts do not start a retained shaped cluster. Unchanged also: mode 15, where Latin letters widen too and the share is neither equal nor proportional (unpinned).',
});

export const WORD_NUMERIC_DECIMAL_TAB_INFERENCE = defineCompatibilityRule({
  id: 'word-numeric-decimal-tab-inference',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/decimal-tab-autoalign.test.ts#right-aligns numbers of different digit counts at the decimal tab',
  },
  description: 'Right-align an otherwise tab-less numeric paragraph at its leading decimal tab while leaving non-numeric and no-decimal-tab paragraphs unchanged. Classification: Microsoft-note derived, [MS-OI29500] §2.1.556(d): Word inserts an implicit tab in a table cell that has a single decimal tab stop (origin PR #555). Paragraphs outside table cells are not covered by that note; the evidence stays a regression test until the section is catalogued and the scope narrowed.',
});

export const WORD_NUMBERING_MARKER_OVERFLOW_TAB_ADVANCE = defineCompatibilityRule({
  id: 'word-numbering-marker-overflow-tab-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/numbered-marker-tab-advance.test.ts#advances the body past the marker to the next tab stop, not onto indentLeft',
  },
  description: 'When a numbering marker overruns its hanging-indent budget, advance the body to the next reachable tab stop beyond the marker edge. Classification: normative, ECMA-376 §17.9.28 suff tab with §17.3.1.37 tab stops; Word PDF of a corpus document in PR #555.',
});

export const WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB = defineCompatibilityRule({
  id: 'word-numbering-suffix-coincident-list-tab',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/numbering-marker.test.ts#keeps a suffix tab on the list stop coincident with the marker end',
  },
  description: 'For the tab synthesized by a numbering suffix, accept an authored numeric list tab coincident with the shaped marker end instead of advancing to the next automatic tab stop. Library policy, no Office measurement: coincident list-tab handling from PR #1080; no control.',
});

export const WORD_NUMBERING_SPACE_SUFFIX_ARIAL = defineCompatibilityRule({
  id: 'word-numbering-space-suffix-arial',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-numsp',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A numbering level with w:suff="space" separates marker and body by one U+0020 space in Arial at the marker font size, independently of marker, paragraph-mark, run, docDefaults and Normal-style fonts. All 48 rows of blk-0898-numsp-{nostyles,styles}-mode{none,14,15} identify the suffix PDF span as ArialMT at the marker size. Measured classes (PDF space advance in pt; nominal Arial comparison): level Courier New 12, mark absent or Verdana 16, run Courier New 12 or Georgia 18: 3.38 (Arial 12 = 3.33; Courier New 12 = 7.20, Verdana 16 = 5.63, Georgia 18 = 4.34); level Courier New without size, mark Verdana 16: marker Courier New 16, space 4.48 (Arial 16 = 4.44); level size 12 only, mark Verdana 16: marker Verdana 12, space 3.33 (Verdana 12 = 4.22); no level rPr, mark Verdana 16: marker Verdana 16, space 4.41 (Arial 16 = 4.44); level Verdana 16, mark Courier New 12: space 4.41 (Arial 16 = 4.44); docDefaults Tahoma 8 and Normal Trebuchet MS 10, no level/mark rPr: marker Trebuchet MS 10, space 2.72 (Arial 10 = 2.78; Trebuchet 10 = 3.01); no styles part or level/mark rPr: built-in host default YuMincho 10.5, space 2.90 (Arial 10.5 = 2.92). Identical in modes unset, 14 and 15, the only measured modes; unmeasured modes 11/12/13 and all other explicit modes keep the previous marker-font space. Scope: the measured classes only — left-to-right paragraphs, regular (not bold or italic, not complex-script) markers in ordinary Latin and East Asian faces, picture bullets excluded. Right-to-left paragraphs, bold or italic markers, complex-script markers, symbol-encoded markers (Private Use Area text or text the shared symbol-font table remaps) and modes 11/12/13 are unmeasured and keep the previous marker-face space. Advances come from TextLayoutService font data and its metric-compatible fallback, never these comparison numbers or a fitted em fraction. ECMA-376 §17.9.28 and §17.18.46 specify the space character, not this font choice.',
});

/** WORD_NUMBERING_SPACE_SUFFIX_ARIAL admission: the measured classes only;
 * everything else keeps the space in the marker's own face. */
export function wordNumberingSpaceSuffixUsesArial(input: Readonly<{
  compatibilityMode: number | undefined;
  leftToRight: boolean;
  markerWeight: number;
  markerStyle: 'normal' | 'italic';
  complexScriptMarker: boolean;
  symbolMarker: boolean;
}>): boolean {
  void WORD_NUMBERING_SPACE_SUFFIX_ARIAL;
  return (input.compatibilityMode === undefined || input.compatibilityMode === 14 || input.compatibilityMode === 15)
    && input.leftToRight
    && input.markerWeight <= 400
    && input.markerStyle === 'normal'
    && !input.complexScriptMarker
    && !input.symbolMarker;
}

export const WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK = defineCompatibilityRule({
  id: 'word-numbering-marker-paragraph-mark-fallback',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-paragraph-mark-formatting',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When numbering-level rPr omits a marker formatting axis, Word takes that axis from the effective paragraph-mark rPr rather than a content run. A numbering-level concrete value or explicit auto remains authoritative, and body and text-box stories use the same cascade.',
});

/** Compatibility projection governed by {@link WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB}. */
export function wordNumberingSuffixAcceptsCoincidentListTab(
  markerEndPt: number,
  stop: Readonly<{ pos: number; alignment: string }>,
): boolean {
  return stop.alignment === 'num' && Math.abs(stop.pos - markerEndPt) <= 1e-6;
}

export const WORD_TAB_STOP_PAGE_EDGE_CLAMP = defineCompatibilityRule({
  id: 'word-tab-stop-page-edge-clamp',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/rtl-tab-stops.test.ts#pins a page number to the left text margin when the stop is past it',
  },
  description: 'Clamp content assigned to a tab stop beyond the trailing text edge back onto that edge instead of placing ink outside the page content band. Evidence (observed Office behaviour): Word comparison of a corpus document (issues #820 and #830) in PR #835.',
});

export const WORD_DICTIONARY_SEA_ATOMIC_CHUNK = defineCompatibilityRule({
  id: 'word-dictionary-sea-atomic-chunk',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/sea-justified-fit.test.ts#Rule 2: a no-space chunk that fits a full line moves whole instead of splitting',
  },
  description: 'Move a glued dictionary Southeast-Asian chunk to a fresh line whole when it fits that full line, using dictionary breaks only when the chunk itself is overlong. Evidence (observed Office behaviour): Word PDF calibration fixture in issue #991 / PR #1003.',
});

export const WORD_NUMBERED_WINDOW_FRAME = defineCompatibilityRule({
  id: 'word-numbered-window-frame',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/legacy-numbered-window-frame.test.ts#rebases automatic and authored stops without rebasing the implicit hanging stop',
  },
  description: 'Observed Office legacy false-flag frame: the original 40 numbered float-window rows and 48 controlled compatibility variants distinguish implicit hanging stops, custom stops, automatic grids, equality/pre-start exclusions, narrow heads and tight windows. The supported class is compatibility mode absent, 12 or 14; doNotSuppressIndentation, ww11IndentRules, doNotUseIndentAsNumberingTabStop and noTabHangInd must not be true. It is restricted to ordinary root-body horizontal LTR left/start paragraphs with self-owned initially unresolved floats, no external exclusions, no East Asian text or active line/character grid, a nonempty left text marker with TAB suffix, nonpositive first indent and supported left stops. For actual window origin g, left indent L and hanging H, marker origin is g+L-H; the implicit stop is max(L,g), custom stop c is g+c and the automatic grid is based at g. Actual marker advance q and the first atom are admitted in that same per-window frame through initial breaking, retry and retained placement; no font-derived offset or width correction is added. This is a bounded observation, not the true-flag rule of Part 4 sections 14.8.3.12, 14.8.3.16 or 14.8.3.31. Modern modes, RTL, external/mixed or pre-resolved exclusions, true flags and unsupported geometry retain the established path. Eight original below-float Y residuals remain; source partitions and horizontal improvement do not establish complete Word equivalence.',
});

/** Issue #2055 controls (blk-0898-floats-num-mode15, Windows Word 16.0.20527,
 * Courier New 12): a "1." marker with tab, space and nothing suffixes, left
 * 50 pt hanging 50 pt, square floats leaving the gaps [0,100], [150,250] and
 * [300,450] pt, 15-19 character head words, and a custom stop at 340 pt; a
 * tight triangle on the left whose edge moves the marker between line-grid
 * positions. */
export const WORD_NUMBERED_FIRST_LINE_GAP_FIT = defineCompatibilityRule({
  id: 'word-numbered-first-line-gap-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-floats-num',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A numbered first line beside floats takes the first window, in reading order and then down the page, that holds its first word after the marker body that window resolves (the body offset the indents give is no width floor for a window, before or after a rejection, so admission stays with break-time placement, hanging punctuation included; this follows from the per-window body: no control has that offset wider than a later window) (the WORD_NUMBERING_MARKER marker edge and suffix tab of issue #1883, per window): in mode 15 a 15 or 17 character word goes to the [300,450] gap with the marker at 300 and the body at the 324 pt default stop (the 340 pt custom stop for the custom row), although the first gap leaves 50 pt after its 50 pt body; an 18 or 19 character word that no gap holds goes below the floats (60 pt) with the body at 50 pt; beside the tight triangle the line steps to the grid line whose edge resolves a 144 pt body. Space and nothing suffixes follow the same order. Modes unset and 14 differ for a paragraph that anchors its own floats (the [300,450] body at the 336 pt stop, the 17 character and custom-stop rows below the floats, and the triangle rows at other positions), a separate pre-15 anchor-paragraph behaviour that is not reproduced. The suffix-space font discrepancy in these controls (marker-face 7.2 pt versus Word 3.34 pt) is handled separately by WORD_NUMBERING_SPACE_SUFFIX_ARIAL.',
});

/** Issue #2057 controls (blk-0898-floats-head-mode15, Windows Word 16.0.20527,
 * Courier New 12; both, left, distribute and the three kashida types). */
export const WORD_OVERLONG_HEAD_AFTER_ANCHOR = defineCompatibilityRule({
  id: 'word-overlong-head-after-anchor',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-floats-head',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A first word wider than the whole line that follows the anchor of a floating object starts on the first line that admits it: in mode 15 an 80 character word (576 pt) beside a 100 pt window, under a 10 or 40 pt square float, is broken from the float bottom (66 + 14 characters at 10 and 40 pt), as a word that fits the band moves whole to the float bottom. The anchor character has no advance, so the line holding only it is the word\'s empty line; ending that line first would set the word one line lower. The same holds in every mode for a float that does not narrow the line (blk-0898-floats-negind, a 720 pt word after the anchor of a float 1 pt past the band edge: its first piece starts the paragraph in modes unset, 11, 12, 14 and 15). Before mode 15 (modes unset and 14) Word instead breaks a narrowed anchor paragraph\'s first word inside the window (40 pt: 5 characters, 100 pt: 13, fewer under medium and high kashida) on every line beside the float when the paragraph is the first on its page; that behaviour is not reproduced (see WORD_FLOAT_GAP_FLOW, which keeps moving such a word below the object). Library scope policy preserves the earlier anchor-only line allocation whenever a pre-15 window narrows or displaces the line; applying the mode-15 suppression there can move an unchanged glyph line farther from Word. The controls hold Latin words only (one run; a word joined across differently formatted runs takes the same line); a Thai, Lao or Khmer dictionary word wider than the line after an anchor (alone or joined with other runs into one unit) is not measured and keeps the previous behaviour (the anchor-only line is ended before the word), pending Word measurements.',
});

/** WORD_OVERLONG_HEAD_AFTER_ANCHOR only establishes the narrowed/displaced
 * rule in mode 15. Preserve prior allocation in older or unauthored modes;
 * their anchored-head behavior differs, and modes 11–13 are unmeasured.
 * An unchanged full-width line keeps the separately measured all-mode rule. */
export function wordPreservesPre15AnchorLine(compatibilityMode?: number): boolean {
  return compatibilityMode !== 15;
}

export const WORD_OVERLONG_TOKEN_EMERGENCY_BREAK = defineCompatibilityRule({
  id: 'word-overlong-token-emergency-break',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/run-inline-formatting.test.ts#breaks a no-space token wider than the line at the character level',
  },
  description: 'Emergency-break an overlong token at grapheme-safe character boundaries on an empty line so the complete token remains inside the content band. Library policy, no Office measurement: emergency break so an overlong token stays inside the column (PR #564); Word behaviour asserted, not measured.',
});

export const WORD_EXTERNAL_LINK_SYNTAX_BREAKS = defineCompatibilityRule({
  id: 'word-external-link-syntax-breaks',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'external-link-syntax-formatting-seam-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Treat readable separators in the path and query of displayed external URLs as line-break opportunities, while keeping the scheme and authority intact and preserving authored no-break hyphens and grapheme clusters.',
});

/** Compatibility projection governed by {@link WORD_EXTERNAL_LINK_SYNTAX_BREAKS}.
 * `graphemeBoundaries` and `authoredNoBreakOffsets` are UTF-16 offsets in the
 * complete displayed link token, not in an individual formatting run. */
export function wordExternalLinkSyntaxBreakOffsets(
  text: string,
  graphemeBoundaries: ReadonlySet<number>,
  authoredNoBreakOffsets: ReadonlySet<number>,
): readonly number[] {
  const scheme = /^[A-Za-z][A-Za-z0-9+.-]*:\/\//u.exec(text);
  if (!scheme) return [];
  const authorityStart = scheme[0].length;
  const authorityEndCandidate = text.slice(authorityStart).search(/[/?#]/u);
  const authorityEnd = authorityEndCandidate < 0
    ? text.length
    : authorityStart + authorityEndCandidate;
  const offsets: number[] = [];
  for (let index = authorityEnd; index < text.length; index += 1) {
    const character = text[index]!;
    const offset = index + 1;
    const readable =
      (character === '/' && index > authorityEnd)
      || character === '-'
      || character === '?'
      || character === '&';
    if (
      readable
      && graphemeBoundaries.has(offset)
      && !authoredNoBreakOffsets.has(offset)
    ) offsets.push(offset);
  }
  return offsets;
}

export const WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT = defineCompatibilityRule({
  id: 'word-run-vertical-align-baseline-shift',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/run-char-metrics-render.test.ts#w:vertAlign raises superscript, lowers subscript, and leaves ordinary baselines unchanged',
  },
  description: 'Retain the established run-level baseline displacement for vertically aligned text: superscript rises by 0.35 of its authored font size and subscript falls by 0.15, while the separately authored w:position remains additive. Library policy, no Office measurement: super/subscript shift of 0.35/0.15 em with the 0.65 size (PR #8); ECMA-376 §17.3.2.42 gives no amount, and the font rule (OS/2 ySuperscriptYOffset / ySubscriptYOffset) is not in the reference metrics table.',
});

export const WORD_UNIFORM_RUN_POSITION_LEADING = defineCompatibilityRule({
  id: 'word-uniform-run-position-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'uniform-run-position-leading',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When every metric-bearing item on a line has the same non-zero w:position, Word places the allocated line box around the positioned glyphs. A line containing a differently-positioned item retains the full relative displacement; automatic line allocation is handled separately.',
});

/** Paint-relative baseline position governed by
 * {@link WORD_UNIFORM_RUN_POSITION_LEADING}. */
export function wordUniformRunPositionPaintPt(
  authoredPositionPt: number,
  commonLinePositionPt: number,
): number {
  return commonLinePositionPt === 0
    ? authoredPositionPt
    : authoredPositionPt - commonLinePositionPt / 2;
}

/** Compatibility projection governed by
 * {@link WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT}. */
export function wordRunVerticalAlignRaisePt(
  verticalAlign: string | null | undefined,
  authoredFontSizePt: number,
): number {
  if (verticalAlign === 'super') return authoredFontSizePt * 0.35;
  if (verticalAlign === 'sub') return -authoredFontSizePt * 0.15;
  return 0;
}

export const WORD_FAR_EAST_SINGLE_LINE_FACTOR = OFFICE_FAR_EAST_SINGLE_LINE_FACTOR;

export const WORD_INLINE_PICTURE_AUTO_LEADING = defineCompatibilityRule({
  id: 'word-inline-picture-auto-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'inline-picture-auto-leading-height-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For an inline picture or drawing shape (wp:inline WPS or in-flow VML text box) with automatic spacing at or above one line, add the authored leading of the selected text/paragraph-mark face to the natural drawing/text baseline union, rather than multiplying the drawing height. Image-only, text-only and alternating 28.35pt-picture controls at line=240/259 and exact=259 distinguish the two advances; earlier Word controls covered picture heights 5–255pt at auto multiples 1, 1.079 and 1.15. Windows Word 16.0.20527 issue #2063 PDF controls (72 inline cases across no-styles/docDefaults documents, compatibility mode omitted/15) also establish box-only single spacing equal to box height and mixed-text box bottoms on the baseline. At 1.5 spacing a 40pt box with a 10pt Times New Roman mark adds about 5.74pt, not 20pt. Additional Windows mode-15 controls with Times New Roman, Arial and Yu Mincho distinguish document defaults, resolved Normal/default styles, an applied paragraph style, direct drawing-run font/size and direct paragraph-mark formatting: automatic leading follows the effective outer drawing run, with paragraph style below direct run rPr; a direct paragraph mark does not select it. Literal pictures and WPS shapes agree. Grid-snapped allocation remains independent, and the existing document-script font-slot policy is preserved. When an authored face is unavailable, the selected fallback supplies its own metrics; direct paragraph-mark size must not compensate for that resource difference. ECMA-376 §17.7.5.1 leaves omitted document defaults application-defined. Without document/style font declarations, preserve unspecified font slots and selected generic metrics rather than guessing an Office host default or borrowing the direct paragraph mark. In controlled inline-object clones, explicit Times New Roman 10pt default-font declarations make the renderer’s source-case page boundaries agree with Word; their fontless originals can differ. Fully authored Windows mode-15 auto-single controls use outer run8/20pt, object8/40pt and mark8pt: 16 WPS cases cover Times New Roman/Arial, four literal-picture cases cover Times New Roman, and four static VML textbox cases cover Times New Roman. DrawingML WPS and literal pictures reserve their outer-run natural minimum as extra ascent while preserving actual suffix descent; large objects retain their actual geometry. The static in-flow VML textbox controls have the same minimum advance but keep their object top, reserving extra below the visible box. The minimum is limited to horizontal off-grid main-body flow, with mixed ordinary text admitted only for DrawingML; mixed VML, relative VML, active WordArt including retained textbox content, unknown/public origins, ruby, numbering, frames and other stories remain outside this measured class. The parser retains actual source provenance; this minimum is not an invisible glyph or additional descent. The earlier box-height-only observation applies when the resolved outer-run minimum does not exceed the actual union. Minimum-height (atLeast) spacing remains outside this outer-run minimum rule. Controlled cases retain a preexisting object-top mismatch with Word even when line advance agrees; synthetic host metrics must not compensate for that separate limitation. Exact spacing and lines without inline pictures/shapes retain their own allocation.',
});

/** Observed outer-run leading selection is bounded to mode-15 box-only body
 * paragraphs. The caller additionally admits only horizontal off-grid root
 * body flow; other stories, mixed text and legacy modes keep their old route. */
export function wordInlineDrawingRunLeadingApplies(
  compatibilityMode: number | undefined,
  paragraph: Readonly<{
    lineSpacing?: Readonly<LineSpacing> | null;
    numbering?: unknown;
    framePr?: unknown;
    runs: readonly Readonly<{ type: string; anchor?: boolean; inline?: boolean }>[];
  }>,
): boolean {
  return compatibilityMode === 15 && !paragraph.numbering && !paragraph.framePr
    && paragraph.lineSpacing?.rule === 'auto' && paragraph.lineSpacing.value > 1
    && paragraph.runs.length > 0
    && paragraph.runs.every(run => (run.type === 'image' && run.anchor !== true)
      || (run.type === 'shape' && run.inline === true));
}

/** Word-observed inline-picture projection of ECMA-376 §17.3.1.33. */
export function wordInlinePictureAutoLineHeightPx(
  naturalUnionPx: number,
  textSinglePx: number,
  multiple: number,
): number {
  void WORD_INLINE_PICTURE_AUTO_LEADING;
  return naturalUnionPx + textSinglePx * (multiple - 1);
}

/** Windows-adjudicated Word automatic line allocation. OS/2 code-page bits
 * 17–20 select 1.3 × usWin, with added leading split equally above/below.
 * Historical Mac 16.112.4 synthetic controls established the code-page class
 * (independent of Latin cmap) but used hhea; Windows 16.0.20430 Yu/equal-box
 * controls select usWin instead (see the shared projection and the grid/specified-line rules).
 * Clear bits retain signed hhea leading. ECMA-376 §17.3.1.33 defines the
 * multiplier, not this table choice. Missing usWin declines the FE projection. */
export function wordOpenTypeAutoLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
  winAscent?: number;
  winDescent?: number;
  farEastCodePage: boolean;
}>): Readonly<{
  lineHeightRatio: number;
  designAscentRatio: number;
  designDescentRatio: number;
}> | null {
  return officeOpenTypeAutoLineRatios(metrics);
}

export function wordEastAsianGridLineCells(
  naturalHeightPx: number,
  pitchPx: number,
): number {
  return pitchPx > 0 ? Math.max(1, Math.ceil(naturalHeightPx / pitchPx)) : 1;
}

export function wordFarEastSingleLinePx(
  intendedSinglePx: number,
  emPx: number,
): number {
  return intendedSinglePx > 0
    ? intendedSinglePx
    : emPx * WORD_FAR_EAST_SINGLE_LINE_FACTOR;
}

/** Compatibility projection governed by
 * {@link WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM}. */
export function wordUseFeLayoutInheritedGridHeightPx(
  allocatedCellHeightPx: number,
  pitchPx: number,
  inheritedMultiple: number,
): number {
  return Math.max(allocatedCellHeightPx, pitchPx * inheritedMultiple);
}

export function wordGridAtLeastLineHeightPx(
  naturalPx: number,
  authoredMinimumPx: number,
  gridMinimumPx: number,
): number {
  return Math.max(naturalPx, authoredMinimumPx, gridMinimumPx);
}

export function wordDegenerateLineSpacingIsSingle(
  rule: string,
  value: number,
): boolean {
  return (rule === 'exact' || rule === 'auto') && value <= 0;
}

export const WORD_LATIN_DESIGN_GRID_CELLS = defineCompatibilityRule({
  id: 'word-latin-design-grid-cells',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'float-grid-picture-origin',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'Issue #1674 modes 14/15 controls reserve 20/40/40pt for 10/20/30pt Arial single lines on a 20pt grid, including a preceding 20pt line. No-grid and snap-off controls retain natural advances. Apply whole-cell counting only to visible text with an admitted non-Far-East reference design profile. Far-East reference faces, native fallback boxes, empty marks, ruby, explicit multiples and exact/atLeast spacing retain their established paths. Later independent-font and anchor-free controls established WORD_SPECIFIED_TEXT_LINE_BOX for its separately gated ordinary Latin class; every other exact/atLeast class retains the established placement. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export function wordLatinDesignGridSingleHeight(natural: number, pitch: number, admittedDesign: number): number {
  // The admitted face controls only its own whole-cell reserve. Other fonts,
  // inline objects and baseline displacements still own their natural extent
  // (§17.3.1.33); admission of one run must not shrink any peer's line box.
  return Math.max(natural, admittedDesign > 0
    ? Math.max(1, Math.ceil(admittedDesign / pitch)) * pitch : pitch);
}

export const WORD_SPECIFIED_TEXT_LINE_BOX = defineCompatibilityRule({
  id: 'word-specified-text-line-box',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'specified-spacing-anchor-free-host-font-grid-matrix',
    application: 'Microsoft Word',
    version: 'export build unrecorded (16.113.3 installed at measurement)',
    platform: 'macOS 27.0',
  },
  description: 'Issue #1674: 636 printing-PDF pages, modes 14/15, no grid and 20pt lines/linesAndChars grids, Arial/Times New Roman/MS Mincho/Yu Mincho/Verdana/Calibri Latin glyphs. Exact 12/20/24/30pt lines put the baseline four fifths down the authored line, independent of the 10pt visible face or zero/one/two hosts; 30pt visible-text and 30pt-host diagnostics distinguish font ascent, centering and host-height hypotheses. Off-grid atLeast 24/40pt retains the normal design descent and places added leading before the normal line box. Grid-atLeast uses the same centered normal-cell box as explicit auto=1, with minimum leading before that box: 24pt minima remain 24pt for 10pt visible text, while oversized 30pt normal boxes reserve 40pt/60pt. A same-face floating host does not enlarge the admitted visible-text box, including the measured 30pt hosts. Independent readback of the complete source-bound matrix establishes these allocations and baselines. Fresh Word 16.113.3 mode-15 Arial-10pt controls with 13 lines, a 20pt grid, 40pt minima, three preceding-line phases, grid/snap-off/exact controls and line anchors distinguish device projection from layout: twelve text steps total exactly 480pt, with alternating 40.08/40.08/39.84pt or rotations, and anchors retain exact point advances. Do not accumulate the printing cadence or round the authored minimum to a grid-cell multiple. Explicit auto=1 diagnostics on an active grid use the same normal grid box; off-grid diagnostics already agree with the established path, which is preserved. PDF coordinates/font sizes are quantized to 0.24pt; this is export precision, not a layout correction. The exact-height evidence interval is 12 through 30pt; smaller captions do not isolate baseline placement from accumulated printing cadence, and heights outside this interval retain the established path. Scope is homogeneous regular unpositioned undecorated Latin text and matching paragraph-mark face/size with an admitted authored reference profile and a native or positively loaded installed Office local() face. Registered local aliases retain the authored family from the resolver; Application-provided SFNT, CSS, embedded, Google and substitute resources remain outside the evidence. Float-exclusion-conditioned origins are also excluded: a narrow side-gap Office counterexample has a different physical partition, so this round does not isolate its baseline from the wrap origin. Admission is paragraph-wide: a mixed-script or decorated continuation paragraph cannot switch contracts on its ASCII-only undecorated physical lines. Mixed scripts/font axes/faces/sizes, differing paragraph marks, decorated/hyperlink text, transformed text, ruby, inline objects, numbering and resource/fallback faces are unmeasured and retain the established path. Windows Word 16.0 build 16.0.20430 retest: Exact-height baseline offsets retain four fifths of the line. Yu Mincho no-grid atLeast 24pt with 10pt text has a 19.20pt baseline offset from paragraph shading top versus 20.16pt in the retained Mac PDF. Oversized 30pt text atLeast 24pt advances 50.16pt versus 42.96pt; 20pt auto1 diagnostics advance about 33.4pt versus 28.7pt. Adjudication (2026-10-16): read the selected host face tables and use the Windows result when identical vertical data yields different Office output. The tested Mac Office and Windows PDF Yu Mincho faces have hhea=(1802, -455, 1024), usWin=(2038, 598), UPM=2048. Windows matches 1.3 times usWin: 30pt atLeast 24pt advances 50.20pt with baseline offset 35.65pt; 10pt atLeast 24pt has offset 19.15pt, within 0.24pt PDF precision. The Mac hhea result remains historical evidence, not the selected policy. No viewer-OS or font-name selector is used. This establishes the tested regular static faces, not arbitrary Office versions, variable fonts or fallback resource identity. Windows Word 16.0 build 16.0.20430 #1884 controls (Arial 10pt and Times New Roman 12pt text; exact 12/18/30pt, atLeast 18/30/60pt; 6-20pt Symbol bullets and Arial numbers) extend the class to numbered lines and to paragraph marks of another face at the same size, including an East Asian mark slot: plain, numbered and empty items all keep four fifths of an exact line and the text descent above the bottom of an atLeast line. A taller first-line marker joins the atLeast normal box by its ascent over the text descent (paragraph-measure.ts).',
});

/** Word observation, not the normative centered-text rule: ECMA-376
 * §17.3.1.33 defines exact/atLeast units and describes bottom placement/clipping
 * for too-small lines and centering for too-large lines. [MS-OI29500]
 * §2.1.60 discusses style-hierarchy spacing but supplies no baseline formula.
 * The measured Word printing output instead preserves a fixed 4:1 exact baseline partition,
 * including oversized text that visibly overhangs without top clipping.
 * This partition is supported by the complete independent-font/host matrix,
 * not a fitted ascent or a sample-specific offset. §17.6.5 makes exact spacing
 * override the grid; non-exact normal boxes retain whole-cell grid leading.
 * Unsupported classes are declined by specifiedTextLineMetrics beside its gate. */
export function wordSpecifiedTextLineMetrics(input: Readonly<{
  rule: string; value: number; descentPt: number;
  singlePt: number; pitchPt: number | null;
}>): Readonly<{ advancePt: number; baselineOffsetPt: number }> {
  void WORD_SPECIFIED_TEXT_LINE_BOX;
  if (input.rule === 'exact') {
    return { advancePt: input.value, baselineOffsetPt: input.value * 4 / 5 };
  }
  const normalPt = input.pitchPt !== null && input.pitchPt > 0
    ? Math.ceil(input.singlePt / input.pitchPt) * input.pitchPt
    : input.singlePt;
  const advancePt = input.rule === 'atLeast' ? Math.max(input.value, normalPt) : normalPt;
  return {
    advancePt,
    baselineOffsetPt: advancePt - (normalPt - input.singlePt) / 2 - input.descentPt,
  };
}

/** How {@link WORD_LINE_END_JUSTIFICATION} justifies a line of a justified paragraph:
 * at its gaps ('word'), by an equal advance after every character
 * ('character'), at its gaps or else centred ('word-or-centre'), or not at all.
 * Soft-wrapped lines keep the paragraph's ordinary justification. */
export function wordLineEndJustification(
  input: Readonly<{ alignment?: string; isLastLine: boolean; stretchLastLine?: boolean;
    doNotExpandShiftReturn?: boolean; lineEndJustificationBeforeMode15?: boolean; endsAtHardBreak?: boolean }>,
  endsWithBreak: boolean,
): 'word' | 'character' | 'word-or-centre' | null {
  if (!endsWithBreak && !input.isLastLine) return 'word';
  // A line ended by a page or column break is justified like a soft-wrapped
  // line before compatibility mode 15 and like a final line from mode 15;
  // doNotExpandShiftReturn does not apply to it.
  const hardBreak = !endsWithBreak && input.endsAtHardBreak === true;
  if (hardBreak && input.alignment !== 'distribute') {
    return input.lineEndJustificationBeforeMode15 === true ? 'word' : null;
  }
  if (input.alignment === 'distribute' || (input.alignment === undefined && input.stretchLastLine)) {
    return input.lineEndJustificationBeforeMode15 === true ? 'character' : 'word-or-centre';
  }
  if (!endsWithBreak || input.doNotExpandShiftReturn === true) return null;
  return input.alignment === 'thaiDistribute' && input.lineEndJustificationBeforeMode15 !== true
    ? 'character'
    : 'word';
}

/** Unlike soft candidates, only the true authored EOF retains this separator
 * in admission. Source/context ownership, not paragraph-fragment position,
 * establishes EOF; measured glyph geometry and intrinsic/paint stay charged. */
export const WORD_LTR_RTL_ARABIC_EOF_SPACE = defineCompatibilityRule({
  id: 'word-ltr-rtl-arabic-eof-space',
  evidence: { kind: 'regression-test', reference: 'packages/docx/src/kashida-line-fit.test.ts#charges only the original EOF SPACE of a mode15 LTR paragraph with one RTL Arabic source' },
  description: 'Observed mode15 literal left/LTR paragraph with one authored RTL Arabic-letter/U+0020 source run ending in SPACE: soft prefixes fit their visible advance, but a prospective candidate reaching the original complete source end charges the existing acquired separator. Two installed-font controls independently bracket visible versus charged widths; matched mode14 preserves visible final fitting. This is not prescribed by rFonts/rtl/cs (§§17.3.2.26/.30/.7), and no parser omission or font selector is inferred. Initial capability requires complete authoritative ordinary horizontal main-body source, one fixed column, no document floating source, frame, table/story, grid, tab, numbering, ruby, transformed/protected/nontext or mixed-script input. Absent/false rtl, cs-only, other scripts, modes and alignments remain unchanged. A queue/fragment end is never original EOF; source sequence and UTF-16 offsets survive slicing and continuation. Only fitting changes: acquired widths, intrinsic geometry, source SPACE and natural paint remain intact. The preexisting physical soft-tail origin residual is not repaired or certified as contained by this comparison. No new reserve or tolerance applies.',
});
