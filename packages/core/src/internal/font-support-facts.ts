/** Sufficient OpenType definedness certificates, not a glyph interpreter.
 * GSUB replaces cmap's default glyphs (OpenType GSUB §§Overview, Lookup types
 * 1–8); GDEF classifications and lookup flags can skip missing glyphs. Inspect
 * every lookup, independent of guessed Canvas feature activation. A safe output
 * certificate is useful even when the stronger missing-isolation proof fails.
 * Unsupported mechanisms and budget exits remain unknown. No font bytes,
 * lookup graph or glyph arrays survive this bounded parse.
 */
import type { FontSupportFacts, FontTable } from './font-support-registry.js';
export type { FontSupportFacts, FontTable } from './font-support-registry.js';
export { fontSupportFacts, retainFontSupportFacts } from './font-support-registry.js';
import { readFontGlyphCount } from './font-glyph-domain.js';
const erasureGlyphs = new WeakMap<FontSupportFacts, ReadonlySet<number>>();
/** Consume transient glyph-domain analysis before retaining parser facts. */
export function takeFontErasureGlyphs(value: FontSupportFacts): ReadonlySet<number> | undefined {
  const glyphs = erasureGlyphs.get(value); erasureGlyphs.delete(value); return glyphs;
}
const tag = (value: string) => [...value].reduce((n, c) => (n * 256 + c.charCodeAt(0)) >>> 0, 0);
// The admitted domain D=[0,numGlyphs) comes from validated maxp/cmap.
// Out-of-D GSUB input predicates are unreachable. Preserve original coverage
// indices, parse every referenced byte/count, and audit outputs only for effects
// with in-D consuming inputs. This is induction over the domain, never generic
// malformed-table forgiveness. An invalid reachable output remains unknown.
// Resource governance: bounded table/glyph visits, independent of font names,
// text length or measured outcomes. Table reads and graph edges/closure have
// independent cumulative budgets. Graph memoization prevents reference fanout.
const WORK_LIMIT = 262_144;
const LOOKUP_LIMIT = 4096;
const GRAPH_DEPTH_LIMIT = 64;

export function parseFontSupportFacts(view: DataView, tables: ReadonlyMap<number, FontTable>): FontSupportFacts {
  let glyphCount: number | undefined;
  let nonzero = true;
  let isolated = true;
  let anyIndic3ScriptPresent = false;
  let gsubLookupCount = 0;
  let gsubDisposition: FontSupportFacts['gsubDisposition'] = 'identity';
  let remaining = WORK_LIMIT;
  let graphRemaining = WORK_LIMIT;
  const deleting = new Set<number>();
  const reverse = new Map<number, Set<number>>();
  const stop = (reason: FontSupportFacts['reason']): never => { throw reason; };
  const work = () => { if (--remaining < 0) stop('budget'); };
  const graphWork = () => { if (--graphRemaining < 0) stop('budget'); };
  const u16 = (t: FontTable, at: number): number => {
    work();
    if (!Number.isSafeInteger(at) || at < t.offset || at > t.offset + t.length - 2) stop('malformed');
    return view.getUint16(at);
  };
  const u32 = (t: FontTable, at: number): number => (u16(t, at) * 65536 + u16(t, at + 2));
  const offset = (t: FontTable, base: number, at: number, wide = false): number => {
    const delta = wide ? u32(t, at) : u16(t, at);
    if (!delta) stop('malformed');
    const result = base + delta;
    u16(t, result);
    return result;
  };
  const glyph = (n: number, output = false, readOnly = false) => {
    if (n >= (glyphCount as number)) { if (output) stop('malformed'); return; }
    if (!n) { if (output) nonzero = false; else if (!readOnly) isolated = false; }
  };
  const edge = (input: number, output: number) => {
    graphWork();
    let parents = reverse.get(output);
    if (!parents) { parents = new Set(); reverse.set(output, parents); }
    parents.add(input);
  };
  const effect = (input: number, output: number) => { glyph(input); glyph(output, true); edge(input, output); };
  const coverage = (t: FontTable, at: number, readOnly = false): number[] => {
    const format = u16(t, at), count = u16(t, at + 2);
    const result: number[] = [];
    let previous = -1;
    if (format === 1) {
      for (let i = 0; i < count; i++) {
        const g = u16(t, at + 4 + i * 2);
        if (g <= previous) stop('malformed');
        glyph(g, false, readOnly); result.push(g); previous = g;
      }
    } else if (format === 2) {
      for (let i = 0; i < count; i++) {
        const p = at + 4 + i * 6;
        const from = u16(t, p), to = u16(t, p + 2), index = u16(t, p + 4);
        if (from > to || from <= previous || index !== result.length) stop('malformed');
        glyph(to, false, readOnly);
        for (let g = from; g <= to; g++) { work(); glyph(g, false, readOnly); result.push(g); }
        previous = to;
      }
    } else stop('unsupported');
    return result;
  };
  // Class-zero predicates can match .notdef, but matching is read-only. Every
  // primitive action is audited independently: if none targets/consumes zero,
  // a contextual lookup cannot repair or delete it. This is deliberately
  // stronger than following only the features expected for one Canvas string.
  const classDef = (t: FontTable, at: number): Map<number, number> => {
    const format = u16(t, at);
    const result = new Map<number, number>();
    if (format === 1) {
      const start = u16(t, at + 2), count = u16(t, at + 4);
      if (start + count > 65536) stop('malformed');
      for (let i = 0; i < count; i++) result.set(start + i, u16(t, at + 6 + i * 2));
    } else if (format === 2) {
      const count = u16(t, at + 2); let previous = -1;
      for (let i = 0; i < count; i++) {
        const p = at + 4 + i * 6;
        const from = u16(t, p), to = u16(t, p + 2), cls = u16(t, p + 4);
        if (from > to || from <= previous) stop('malformed');
        for (let g = from; g <= to; g++) { work(); result.set(g, cls); }
        previous = to;
      }
    } else stop('unsupported');
    return result;
  };
  try {
    glyphCount = readFontGlyphCount(view, tables);
    if (!glyphCount) stop('glyph-domain');
    if (tables.has(tag('morx')) || tables.has(tag('mort')) || tables.has(tag('fvar'))) stop('unsupported');
    let zeroClass = 0;
    let hasGlyphClassDef = false;
    let markSetCount = 0;
    const gdef = tables.get(tag('GDEF'));
    if (gdef) {
      const version = u32(gdef, gdef.offset);
      if (![0x00010000, 0x00010002, 0x00010003].includes(version)) stop('unsupported');
      // Validate the header before trusting an omitted glyph classification.
      u16(gdef, gdef.offset + (version === 0x00010000 ? 10 : version === 0x00010002 ? 12 : 16));
      const relative = u16(gdef, gdef.offset + 4);
      if (relative) {
        const classes = classDef(gdef, gdef.offset + relative);
        if ([...classes.values()].some((value) => value > 4)) stop('malformed');
        hasGlyphClassDef = true;
        zeroClass = classes.get(0) ?? 0;
      }
      const attachment = u16(gdef, gdef.offset + 10);
      if (attachment && [...classDef(gdef, gdef.offset + attachment).values()].some((value) => value > 255)) stop('malformed');
      if (version !== 0x00010000) {
        const relativeSets = u16(gdef, gdef.offset + 12);
        if (relativeSets) {
          const sets = gdef.offset + relativeSets;
          if (u16(gdef, sets) !== 1) stop('unsupported');
          markSetCount = u16(gdef, sets + 2);
          for (let i = 0; i < markSetCount; i++) coverage(gdef, offset(gdef, sets, sets + 4 + i * 4, true), true);
        }
      }
      // Attachment points/ligature carets and variation coordinates affect
      // placement, not glyph definedness; this certificate does not parse GPOS.

    }
    const t = tables.get(tag('GSUB'));
    const gsubVersion = t ? u32(t, t.offset) : undefined;
    const gsubMajor = gsubVersion === undefined ? undefined : gsubVersion >>> 16;
    // HarfBuzz 11.0.0 GSUBGPOS dispatch (hb-ot-layout-gsubgpos.hh
    // §§get_script_list/get_feature_list/get_lookup_count/get_lookup) executes
    // majors 1/2; other readable majors return null lists and zero lookups.
    // Source SHA-256 31979c6a0a62f43183c3fc44baa82ff4beb92c8b1acb1100ffa8a04ba6315ecc.
    // This is an explicit canonical-static-v1 inactive-mechanism fact, not
    // unsupported/malformed => safe. Never reinterpret offsets/version bytes.
    if (t && gsubMajor !== 1 && gsubMajor !== 2) gsubDisposition = 'profile-inactive-major';
    if (t && gsubDisposition !== 'profile-inactive-major') {
      gsubDisposition = 'active-open-type';
      const version = gsubVersion;
      if (version !== 0x00010000 && version !== 0x00010001) stop('unsupported');
      if (version === 0x00010001 && u32(t, t.offset + 10)) stop('unsupported');
      // Pinned HarfBuzz 11.0.0 GSUBGPOS uses nullable top-level
      // OffsetTo lists. NULL LookupList contributes no actions; ScriptList
      // tags still select script preprocessing and must remain represented.
      // Every non-NULL offset/selector is validated normally. This is profile
      // semantics, never an exception for arbitrary malformed active tables.
      const nullableList = (at: number) => u16(t, at) ? offset(t, t.offset, at) : 0;
      const list = nullableList(t.offset + 8);
      const count = list ? u16(t, list) : 0;
      gsubLookupCount = count;
      if (count > LOOKUP_LIMIT) stop('budget');
      // The profile inspects all potentially enabled lookups, yet malformed
      // script/feature selectors cannot be silently granted an active proof.
      const features = nullableList(t.offset + 6), featureCount = features ? u16(t, features) : 0;
      for (let i = 0; i < featureCount; i++) {
        const feature = offset(t, features, features + 2 + i * 6 + 4);
        const n = u16(t, feature + 2);
        for (let j = 0; j < n; j++) if (u16(t, feature + 4 + j * 2) >= count) stop('malformed');
      }
      const langSys = (at: number) => {
        const required = u16(t, at + 2), n = u16(t, at + 4);
        if (u16(t, at) !== 0 || required !== 0xffff && required >= featureCount) stop('malformed');
        for (let i = 0; i < n; i++) if (u16(t, at + 6 + i * 2) >= featureCount) stop('malformed');
      };
      const scripts = nullableList(t.offset + 4), scriptCount = scripts ? u16(t, scripts) : 0;
      for (let i = 0; i < scriptCount; i++) {
        anyIndic3ScriptPresent ||= (u32(t, scripts + 2 + i * 6) & 255) === 0x33;
        const script = offset(t, scripts, scripts + 2 + i * 6 + 4);
        const defaultOffset = u16(t, script), n = u16(t, script + 2);
        if (defaultOffset) langSys(script + defaultOffset);
        for (let j = 0; j < n; j++) langSys(offset(t, script, script + 4 + j * 6 + 4));
      }
      const lookupOffsets: number[] = [];
      for (let i = 0; i < count; i++) lookupOffsets.push(offset(t, list, list + 2 + i * 2));
      const visiting = new Set<number>(), done = new Set<number>();
      const subtables = new Set<string>();
      const records = (at: number, n: number, inputs: number) => {
        for (let i = 0; i < n; i++) {
          const sequence = u16(t, at + i * 4), lookup = u16(t, at + i * 4 + 2);
          if (sequence >= inputs || lookup >= count) stop('malformed');
          inspectLookup(lookup);
        }
      };
      const inspectSubtable = (type: number, at: number, depth: number) => {
        if (depth > GRAPH_DEPTH_LIMIT) stop('budget');
        const key = `${type}:${at}`;
        if (subtables.has(key)) return;
        subtables.add(key);
        const format = u16(t, at);
        if (type === 7) {
          if (format !== 1) stop('unsupported');
          const extensionType = u16(t, at + 2);
          if (extensionType === 7 || extensionType < 1 || extensionType > 8) stop('malformed');
          inspectSubtable(extensionType, offset(t, at, at + 4, true), depth + 1); return;
        }
        if (type < 1 || type > 8) stop('unsupported');
        if (type === 1) {
          const inputs = coverage(t, offset(t, at, at + 2), true);
          if (format === 1) {
            const delta = u16(t, at + 4);
            for (const g of inputs) if (g < (glyphCount as number)) { effect(g, (g + delta) & 0xffff); }
          } else if (format === 2) {
            if (u16(t, at + 4) !== inputs.length) stop('malformed');
            for (let i = 0; i < inputs.length; i++) { const out = u16(t, at + 6 + i * 2); if (inputs[i] < (glyphCount as number)) { effect(inputs[i], out); } }
          } else stop('unsupported');
          return;
        }
        if (type === 2 || type === 3 || type === 4) {
          if (format !== 1) stop('unsupported');
          const inputs = coverage(t, offset(t, at, at + 2), true);
          const n = u16(t, at + 4);
          if (n !== inputs.length) stop('malformed');
          for (let i = 0; i < n; i++) {
            const set = offset(t, at, at + 6 + i * 2), size = u16(t, set);
            const reachable = inputs[i] < (glyphCount as number);
            // MultipleSubst empty sequence deletes its input. An empty
            // alternate/ligature set has no applicable substitution action.
            if (!size && reachable && type === 2) deleting.add(inputs[i]);
            if (type !== 4) {
              if (reachable) glyph(inputs[i]);
              for (let j = 0; j < size; j++) { const out = u16(t, set + 2 + j * 2); if (reachable) { glyph(out, true); edge(inputs[i], out); } }
            } else {
              for (let j = 0; j < size; j++) {
                const lig = offset(t, set, set + 2 + j * 2);
                const out = u16(t, lig), components = u16(t, lig + 2);
                if (components < 1) stop('malformed');
                let applicable = reachable, consumesZero = inputs[i] === 0;
                const consumed = [inputs[i]];
                for (let k = 1; k < components; k++) { const component = u16(t, lig + 2 + k * 2); applicable &&= component < (glyphCount as number); consumesZero ||= component === 0; consumed.push(component); }
                if (applicable) { if (consumesZero) isolated = false; glyph(out, true); for (const input of consumed) edge(input, out); }
              }
            }
          }
          return;
        }
        if (type === 8) {
          if (format !== 1) stop('unsupported');
          const inputs = coverage(t, offset(t, at, at + 2), true);
          let p = at + 4;
          for (let side = 0; side < 2; side++) {
            const n = u16(t, p); p += 2;
            for (let i = 0; i < n; i++) coverage(t, offset(t, at, p + i * 2), true);
            p += n * 2;
          }
          const n = u16(t, p); p += 2;
          if (n !== inputs.length) stop('malformed');
          for (let i = 0; i < n; i++) { const out = u16(t, p + i * 2); if (inputs[i] < (glyphCount as number)) { effect(inputs[i], out); } }
          return;
        }
        // Context/chained-context formats reference lookups; never recursively
        // enumerate matching glyph sequences. Read-only input/back/lookahead
        // predicates may include zero; referenced primitive effects, not the
        // predicate alone, determine whether it can be consumed or repaired.
        if (format === 3) {
          let p = at + 2, inputCount = 0, recordCount = 0;
          if (type === 5) {
            inputCount = u16(t, p); recordCount = u16(t, p + 2); p += 4;
            for (let i = 0; i < inputCount; i++) coverage(t, offset(t, at, p + i * 2), true);
            p += inputCount * 2;
          } else {
            for (let side = 0; side < 3; side++) {
              const n = u16(t, p); p += 2;
              if (side === 1) inputCount = n;
              for (let i = 0; i < n; i++) coverage(t, offset(t, at, p + i * 2), true);
              p += n * 2;
            }
            recordCount = u16(t, p); p += 2;
          }
          if (!inputCount) stop('malformed');
          records(p, recordCount, inputCount); return;
        }
        if (format !== 1 && format !== 2) stop('unsupported');
        const inputs = coverage(t, offset(t, at, at + 2), true);
        let p = at + 4;
        if (format === 2) {
          const definitions = type === 5 ? 1 : 3;
          // HarfBuzz 11.0.0 ChainContextFormat2 uses nullable
          // OffsetTo<ClassDef>; NULL is the empty map (all class zero).
          // This is adopted profile semantics, not a normative claim that
          // every shaper admits an omitted context-class table.
          for (let i = 0; i < definitions; i++) { const relative = u16(t, p + i * 2); if (relative) classDef(t, at + relative); }
          p += definitions * 2;
        }
        const sets = u16(t, p); p += 2;
        if (format === 1 && sets !== inputs.length) stop('malformed');
        for (let i = 0; i < sets; i++) {
          const delta = u16(t, p + i * 2); if (!delta) continue;
          const set = at + delta, n = u16(t, set);
          for (let j = 0; j < n; j++) {
            const rule = offset(t, set, set + 2 + j * 2);
            let q = rule, inputCount = 0, recordCount = 0;
            const predicates = (n: number) => {
              for (let k = 0; k < n; k++) {
                const value = u16(t, q + k * 2);
                if (format === 1) glyph(value, false, true);
              }
              q += n * 2;
            };
            if (type === 5) {
              inputCount = u16(t, q); recordCount = u16(t, q + 2); q += 4;
              if (!inputCount) stop('malformed');
              predicates(inputCount - 1);
            } else {
              const back = u16(t, q); q += 2; predicates(back);
              inputCount = u16(t, q); q += 2;
              if (!inputCount) stop('malformed');
              predicates(inputCount - 1);
              const ahead = u16(t, q); q += 2; predicates(ahead);
              recordCount = u16(t, q); q += 2;
            }
            records(q, recordCount, inputCount);
          }
        }
      };
      const inspectLookup = (index: number) => {
        if (done.has(index)) return;
        if (visiting.has(index)) stop('cycle');
        if (visiting.size >= GRAPH_DEPTH_LIMIT) stop('budget');
        visiting.add(index);
        const at = lookupOffsets[index];
        const type = u16(t, at), flags = u16(t, at + 2), n = u16(t, at + 4);
        if (flags & 0x00e0) stop('malformed');
        // LookupFlag acts on the corresponding glyph class. Under pinned HB
        // 11.0.0, GDEF::has_glyph_classes tests ClassDef offset presence (even
        // an empty ClassDef); absent/NULL selects hb_synthesize_glyph_classes
        // in hb-ot-shape.cc. That preserves the original Unicode category on
        // glyph zero: non-ignorable Mn becomes Mark, everything else Base.
        // Thus no ClassDef is not an explicit unclassified zero. IgnoreMarks
        // cannot skip an explicitly classified Base zero, but can skip a
        // synthesized missing mark; IgnoreLigatures alone cannot skip either.
        if (hasGlyphClassDef
          ? (zeroClass === 1 && (flags & 2)) || (zeroClass === 2 && (flags & 4)) || (zeroClass === 3 && (flags & 0xff18))
          : flags & 0xff1a) isolated = false;
        if (flags & 0x0010 && u16(t, at + 6 + n * 2) >= markSetCount) stop('malformed');
        for (let i = 0; i < n; i++) inspectSubtable(type, offset(t, at, at + 6 + i * 2), 0);
        visiting.delete(index); done.add(index);
      };
      for (let i = 0; i < count; i++) inspectLookup(i);
    }
    // Reverse closure quantifies over every supported feature/context; a glyph
    // outside Bad cannot reach a deletion. Blink HarfBuzzShaper's empty-buffer
    // extraction commits no font run, so no-erasure is separate from .notdef
    // preservation. Project this bounded transient graph through cmap next.
    const queue = [...deleting];
    for (let at = 0; at < queue.length; at++) for (const input of reverse.get(queue[at]) ?? []) {
      graphWork(); if (!deleting.has(input)) { deleting.add(input); queue.push(input); }
    }
    const result: FontSupportFacts = Object.freeze({ schema: 'ot-definedness-1', glyphCount,
      nonzeroPreserved: nonzero, missingIsolated: isolated, noErasure: deleting.size === 0,
      gsubDisposition, gsubLookupCount, anyIndic3ScriptPresent, profile: 'canonical-static-v1' });
    if (deleting.size) erasureGlyphs.set(result, deleting);
    return result;
  } catch (reason) {
    return Object.freeze({ schema: 'ot-definedness-1', glyphCount, nonzeroPreserved: undefined,
      missingIsolated: undefined, reason: typeof reason === 'string' ? reason as FontSupportFacts['reason'] : 'malformed' });
  }
}
