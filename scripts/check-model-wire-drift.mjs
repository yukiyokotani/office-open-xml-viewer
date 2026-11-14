#!/usr/bin/env node
// Detects drift between a Rust parser model (the serde JSON wire format) and
// the hand-written TypeScript model types that read it.
//
// The Rust side is described by the JSON Schema that schemars derives from the
// model types behind the dev-only `schema` cargo feature (it follows the same
// serde attributes as the wire: rename/rename_all, skip_serializing_if, tags,
// flatten). The feature is off in every WASM build, so the shipped parsers are
// unchanged. The TS side is read with the TypeScript compiler API. Both graphs
// are walked structurally from their root types, and every difference that is
// not in the commented allowlist fails the check. Stale allowlist entries fail
// too, so the allowlist can only shrink as drift is fixed.
//
// Usage:
//   node scripts/check-model-wire-drift.mjs [format...]      (default: all)
//   node scripts/check-model-wire-drift.mjs --schema <file> docx
//   node scripts/check-model-wire-drift.mjs --list docx       (print all, no allowlist)

import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { MODEL_WIRE_DRIFT_ALLOWLIST, MODEL_WIRE_TS_WIDENING_ALLOWLIST } from './model-wire-drift-allowlist.mjs';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** Per-format wiring: the Rust crate/example that prints the root schema and
 * the TS file + exported root type that models the same JSON. */
export const FORMATS = {
  docx: {
    crate: 'docx-model',
    tsconfig: 'packages/docx/tsconfig.json',
    tsFile: 'packages/docx/src/types.ts',
    tsRoot: 'DocxDocumentModel',
    // Sources of every serialized type reachable from the root, scanned for
    // the null-provenance convention (see scanNullProvenance).
    rustSources: ['packages/docx/model/src', 'packages/ooxml-common/src'],
  },
};


// ---------------------------------------------------------------------------
// Normalized shapes shared by both sides:
//   { kind: 'any' } | { kind: 'prim', name: 'string'|'number'|'boolean'|'null' }
//   { kind: 'literal', values: Set } | { kind: 'array', item } | { kind: 'tuple', items }
//   { kind: 'map', value } | { kind: 'object', name, props: () => Map<name, {shape, optional}> }
//   { kind: 'union', members }
// Object props are lazy so recursive graphs stay finite.

const ANY = { kind: 'any' };
const prim = (name) => ({ kind: 'prim', name });

// ----------------------------- Rust (JSON Schema) --------------------------

function rustShapes(schema) {
  const defs = schema.$defs ?? {};
  const cache = new Map();
  const ref = (pointer) => {
    const name = decodeURIComponent(pointer.replace(/^#\/\$defs\//, ''));
    if (!(name in defs)) throw new Error(`schema: unresolved ${pointer}`);
    if (!cache.has(name)) cache.set(name, lazy(() => convert(defs[name], name)));
    return cache.get(name);
  };

  function convert(node, name) {
    if (node === true || node == null) return ANY;
    if (node === false) return { kind: 'never' };
    if (node.$ref) {
      // schemars puts an internally-tagged enum's tag next to the variant's
      // `$ref` (`{ $ref, properties: { type: { const } } }`).
      const target = ref(node.$ref);
      if (!node.properties) return target;
      return mergeObjects([target, convertObject({ ...node, $ref: undefined }, name)], name);
    }
    const variants = node.oneOf ?? node.anyOf;
    if (variants) {
      // Sibling properties next to oneOf (serde flatten of an enum) apply to
      // every variant.
      const base = node.properties ? convertObject({ ...node, oneOf: undefined, anyOf: undefined }, name) : null;
      const members = variants.map((v, i) => {
        const m = convert(v, name ? `${name}#${i}` : undefined);
        return base ? mergeObjects([base, m], name) : m;
      });
      return union(members, name);
    }
    if (node.allOf) {
      const parts = node.allOf.map((p) => convert(p, name));
      if (node.properties) parts.push(convertObject({ ...node, allOf: undefined }, name));
      return mergeObjects(parts, name);
    }
    if ('const' in node) return literalOrPrim([node.const]);
    if (Array.isArray(node.enum)) return literalOrPrim(node.enum);
    if (Array.isArray(node.type)) {
      return union(node.type.map((t) => convert({ ...node, type: t }, name)));
    }
    switch (node.type) {
      case 'string': return prim('string');
      case 'number':
      case 'integer': return prim('number');
      case 'boolean': return prim('boolean');
      case 'null': return prim('null');
      case 'array':
        if (Array.isArray(node.prefixItems)) {
          return { kind: 'tuple', items: node.prefixItems.map((p) => convert(p)), flags: node.prefixItems.map(() => 1) };
        }
        return { kind: 'array', item: convert(node.items ?? true), minItems: node.minItems, maxItems: node.maxItems };
      case 'object':
        if (!node.properties && node.additionalProperties !== undefined && node.additionalProperties !== false) {
          return { kind: 'map', value: convert(node.additionalProperties) };
        }
        return convertObject(node, name);
      default:
        return node.properties ? convertObject(node, name) : ANY;
    }
  }

  function convertObject(node, name) {
    const required = new Set(node.required ?? []);
    const props = new Map();
    for (const [key, value] of Object.entries(node.properties ?? {})) {
      const optional = !required.has(key);
      let shape = convert(value, name);
      // Serialize-mode schemas leave a field out of `required` only when serde
      // skips it (`skip_serializing_if`). An optional `Option<T>` field that
      // skips exactly `None` never carries null, so the null member is
      // dropped and optionality expresses absence. scanNullProvenance fails
      // the check for any field where that does not hold (another predicate
      // on an Option, or `Option<Option<T>>`).
      if (optional) {
        const inner = shape;
        shape = lazy(() => stripNull(deref(inner)));
      }
      props.set(key, { shape, optional });
    }
    const extra = node.additionalProperties;
    const additional = extra !== undefined && extra !== false ? convert(extra) : undefined;
    return { kind: 'object', name: name ?? '(inline)', props: () => props, additional };
  }

  function mergeObjects(parts, name) {
    return lazy(() => mergeResolved(parts, name));
  }

  function mergeResolved(parts, name) {
    const resolved = parts.map(deref);
    const unionPart = resolved.find((p) => p.kind === 'union');
    if (unionPart) {
      const rest = resolved.filter((p) => p !== unionPart);
      return deref(union(unionPart.members.map((m) => mergeObjects([...rest, m], name))));
    }
    const objects = resolved.filter((p) => p.kind === 'object');
    if (objects.length !== resolved.length) return resolved.length === 1 ? resolved[0] : ANY;
    const label = objects.find((o) => o.name !== '(inline)' && !o.name.includes('#'))?.name ?? name ?? '(inline)';
    return {
      kind: 'object',
      name: label,
      additional: objects.find((o) => o.additional)?.additional,
      props: () => {
        const merged = new Map();
        for (const o of objects) for (const [k, v] of o.props()) merged.set(k, v);
        return merged;
      },
    };
  }

  return convert(schema, schema.title);
}

function literalOrPrim(values) {
  const nonNull = values.filter((v) => v !== null);
  const members = [];
  if (nonNull.length > 0) members.push({ kind: 'literal', values: new Set(nonNull) });
  if (nonNull.length !== values.length) members.push(prim('null'));
  return union(members);
}

function union(members, name) {
  return lazy(() => unionResolved(members, name));
}

function unionResolved(members, name) {
  const flat = [];
  for (const m of members) {
    const d = deref(m);
    if (d.kind === 'union') flat.push(...d.members);
    else if (d.kind !== 'never') flat.push(m);
  }
  // Collapse literal members into one literal set and repeated primitives
  // (`2 | 3` is two number members) into one.
  const literals = flat.filter((m) => deref(m).kind === 'literal');
  const others = flat.filter((m, i) => {
    const d = deref(m);
    if (d.kind === 'literal') return false;
    return d.kind !== 'prim' || flat.findIndex((x) => deref(x).kind === 'prim' && deref(x).name === d.name) === i;
  });
  const out = [...others];
  if (literals.length > 0) {
    out.push({ kind: 'literal', values: new Set(literals.flatMap((l) => [...deref(l).values])) });
  }
  if (out.some((m) => deref(m).kind === 'any')) return ANY;
  return out.length === 1 ? deref(out[0]) : { kind: 'union', members: out, name };
}

/** A shape computed on first use; recursive model graphs are only expanded
 * while they are compared, never while they are built. */
function lazy(fn) {
  let state = 0;
  let value;
  return {
    kind: 'lazy',
    get() {
      if (state === 2) return value;
      if (state === 1) throw new Error('model-wire-drift: a type resolves to itself (cyclic union/flatten)');
      state = 1;
      value = deref(fn());
      state = 2;
      return value;
    },
  };
}

function deref(shape) {
  let s = shape;
  while (s && s.kind === 'lazy') s = s.get();
  return s ?? ANY;
}

// ------------------------------- TypeScript --------------------------------

function tsShapes(format) {
  const configPath = path.resolve(repoRoot, format.tsconfig);
  const config = ts.readConfigFile(configPath, ts.sys.readFile);
  const parsed = ts.parseJsonConfigFileContent(config.config, ts.sys, path.dirname(configPath));
  const file = path.resolve(repoRoot, format.tsFile);
  const { rootDir, declarationDir, emitDeclarationOnly, ...options } = parsed.options;
  const program = ts.createProgram([file], { ...options, noEmit: true, composite: false, declaration: false });
  const checker = program.getTypeChecker();
  const source = program.getSourceFile(file);
  const moduleSymbol = source && checker.getSymbolAtLocation(source);
  const rootSymbol = moduleSymbol && checker.getExportsOfModule(moduleSymbol).find((s) => s.name === format.tsRoot);
  if (!rootSymbol) throw new Error(`${format.tsFile}: cannot find exported ${format.tsRoot}`);
  const rootType = checker.getDeclaredTypeOfSymbol(rootSymbol);
  const cache = new Map();

  function typeName(type) {
    const symbol = type.aliasSymbol ?? type.getSymbol();
    const name = symbol?.getName();
    if (name && !name.startsWith('__')) return name;
    const text = checker.typeToString(type);
    return text.length > 80 ? `${text.slice(0, 77)}...` : text;
  }

  function convert(type) {
    if (!cache.has(type)) cache.set(type, lazy(() => convertUncached(type)));
    return cache.get(type);
  }

  function convertUncached(type) {
    const f = type.flags;
    if (f & (ts.TypeFlags.Any | ts.TypeFlags.Unknown)) return ANY;
    if (f & ts.TypeFlags.Never) return { kind: 'never' };
    if (f & ts.TypeFlags.Union) {
      // `undefined` is kept here; optional properties drop it at the property
      // boundary (absence), anywhere else it is a value JSON cannot carry.
      const members = type.types;
      const bools = members.filter((t) => t.flags & ts.TypeFlags.BooleanLiteral);
      const rest = members.filter((t) => !(t.flags & ts.TypeFlags.BooleanLiteral));
      const shapes = rest.map(convert);
      // `boolean` is `true | false`; a single boolean literal stays a literal.
      if (bools.length === 2) shapes.push(prim('boolean'));
      else shapes.push(...bools.map(convert));
      // `'a' | 'b' | string` widens to string in TS already.
      return union(shapes, type.aliasSymbol?.getName());
    }
    if (f & (ts.TypeFlags.StringLiteral | ts.TypeFlags.NumberLiteral)) return { kind: 'literal', values: new Set([type.value]) };
    if (f & ts.TypeFlags.BooleanLiteral) return { kind: 'literal', values: new Set([type.intrinsicName === 'true']) };
    if (f & (ts.TypeFlags.TemplateLiteral | ts.TypeFlags.StringMapping)) return { kind: 'pattern', text: checker.typeToString(type) };
    if (f & ts.TypeFlags.String) return prim('string');
    if (f & (ts.TypeFlags.Number | ts.TypeFlags.Enum | ts.TypeFlags.EnumLiteral)) return prim('number');
    if (f & ts.TypeFlags.Boolean) return prim('boolean');
    if (f & ts.TypeFlags.Null) return prim('null');
    if (f & (ts.TypeFlags.Undefined | ts.TypeFlags.Void)) return prim('undefined');
    if (checker.isTupleType(type)) {
      const flags = type.target.elementFlags ?? [];
      return {
        kind: 'tuple',
        // An optional element's `undefined` means a shorter tuple.
        items: checker.getTypeArguments(type).map((t, i) => ((flags[i] ?? 1) & 2 ? withoutUndefined(convert(t)) : convert(t))),
        // ts.ElementFlags: Required 1, Optional 2, Rest 4, Variadic 8.
        flags: checker.getTypeArguments(type).map((_, i) => flags[i] ?? 1),
      };
    }
    if (checker.isArrayType(type)) {
      return { kind: 'array', item: convert(checker.getTypeArguments(type)[0]) };
    }
    if (f & ts.TypeFlags.Object || f & ts.TypeFlags.Intersection) {
      const properties = checker.getPropertiesOfType(type);
      const index = checker.getIndexInfoOfType(type, ts.IndexKind.String);
      if (properties.length === 0 && index) return { kind: 'map', value: convert(index.type) };
      const name = typeName(type);
      return {
        kind: 'object',
        name,
        index: index ? convert(index.type) : undefined,
        props: memo(() => {
          const props = new Map();
          for (const symbol of properties) {
            const declaration = symbol.valueDeclaration ?? symbol.declarations?.[0];
            const propType = declaration
              ? checker.getTypeOfSymbolAtLocation(symbol, declaration)
              : checker.getTypeOfSymbol(symbol);
            const optional = (symbol.flags & ts.SymbolFlags.Optional) !== 0;
            const shape = convert(propType);
            props.set(symbol.getName(), { shape: optional ? withoutUndefined(shape) : shape, optional });
          }
          return props;
        }),
      };
    }
    return ANY;
  }

  return convert(rootType);
}

/** An optional property's `undefined` member means absence, not a value. */
function withoutUndefined(shape) {
  return lazy(() => {
    const s = deref(shape);
    if (s.kind !== 'union') return s.kind === 'prim' && s.name === 'undefined' ? { kind: 'never' } : s;
    const members = s.members.filter((m) => !(deref(m).kind === 'prim' && deref(m).name === 'undefined'));
    return members.length === 1 ? deref(members[0]) : { ...s, members };
  });
}

function memo(fn) {
  let value;
  return () => (value ??= fn());
}

// ------------------------------- Comparison --------------------------------

/** Returns drift findings `{ key, category, detail, path }`, one per key.
 * A key is `<owner>.<field>: <category>[ <payload>]`; the payload carries the
 * concrete values (literals, mismatched kinds) so an allowlisted difference
 * cannot silently grow. */
export function compareShapes(rustRoot, tsRoot) {
  const findings = new Map();
  comparer((f) => {
    if (!findings.has(f.key)) findings.set(f.key, f);
  }).walk(rustRoot, tsRoot, { owner: '(root)', field: '(root)', path: '$' });
  return [...findings.values()].sort((a, b) => a.key.localeCompare(b.key));
}

const PROBE_LIMIT = Symbol('probe-limit');

/** Object pairs currently being compared, across nested probes. A pair met
 * again while it is still open is assumed compatible (coinduction), which
 * keeps probes of recursive unions finite. */
const OPEN_PAIRS = new Set();

/** Number of findings comparing `r` with `t` in isolation, capped at
 * `limit + 1`; used to pick the compatible alternative of a union. */
function probeCost(r, t, limit) {
  let n = 0;
  try {
    comparer(() => {
      n += 1;
      if (n > limit) throw PROBE_LIMIT;
    }).walk(r, t, { owner: '?', field: '?', path: '?' });
  } catch (error) {
    if (error !== PROBE_LIMIT) throw error;
  }
  return n;
}

/** The candidate with the fewest differences, or `null` when there is none
 * or when several candidates are equally (in)compatible: the result never
 * depends on the order of union members. */
function bestMatch(r, candidates) {
  if (candidates.length <= 1) return candidates[0] ?? null;
  const costs = candidates.map((c) => probeCost(r, c, 50));
  const min = Math.min(...costs);
  // Any fully compatible member is as good as another.
  if (min === 0 || costs.filter((c) => c === min).length === 1) return candidates[costs.indexOf(min)];
  return null;
}

const scalarKind = (v) => typeof v; // 'string' | 'number' | 'boolean'

function comparer(sink) {
  const visited = new Set();
  const report = (owner, field, category, detail, where, payload) => {
    const key = `${owner}.${field}: ${category}${payload ? ` ${payload}` : ''}`;
    sink({ key, category, detail, path: where });
  };

  function walk(r, t, ctx) {
    r = deref(r);
    t = deref(t);
    // Recursive non-object shapes (`type A = A[]`) are bounded like objects:
    // a pair met again while still open is assumed compatible.
    const pair = `w${objectId(r)}|${objectId(t)}`;
    if (OPEN_PAIRS.has(pair)) return;
    OPEN_PAIRS.add(pair);
    try {
      walkOpen(r, t, ctx);
    } finally {
      OPEN_PAIRS.delete(pair);
    }
  }

  function walkOpen(r, t, ctx) {
    if (r.kind === 'never') return;
    if (r.kind === 'any' || t.kind === 'any') {
      if (r.kind === 'any' && t.kind !== 'any') {
        report(ctx.owner, ctx.field, 'rust-untyped', 'the Rust schema does not constrain this value; it cannot be checked', ctx.path);
      } else if (t.kind === 'any' && r.kind !== 'any') {
        report(ctx.owner, ctx.field, 'ts-untyped', `TS types it as any/unknown; Rust emits ${describe(r)}`, ctx.path);
      }
      return;
    }
    if (t.kind === 'never') {
      report(ctx.owner, ctx.field, 'type-mismatch', 'TS never accepts the value Rust emits', ctx.path, `(Rust ${describe(r)} vs TS never)`);
      return;
    }
    if (hasPrim(t, 'undefined')) {
      report(ctx.owner, ctx.field, 'ts-undefined', 'TS accepts undefined as a value; JSON never carries it', ctx.path);
      t = withoutPrim(t, 'undefined');
      if (t.kind === 'never') return;
    }
    if (hasNull(r) && !hasNull(t)) {
      report(ctx.owner, ctx.field, 'rust-null', 'Rust can emit null; the TS type does not allow null', ctx.path);
    } else if (hasNull(t) && !hasNull(r)) {
      report(ctx.owner, ctx.field, 'ts-null', 'TS allows null; Rust never emits null here', ctx.path);
    }
    r = stripNull(r);
    t = stripNull(t);
    if (r.kind === 'prim' && r.name === 'null') {
      if (!(t.kind === 'prim' && t.name === 'null')) {
        report(ctx.owner, ctx.field, 'variant-never-emitted', 'Rust only emits null', ctx.path, `(TS ${describeDeep(t)})`);
      }
      return;
    }
    if (t.kind === 'prim' && t.name === 'null') {
      report(ctx.owner, ctx.field, 'type-mismatch', 'TS accepts only null', ctx.path, `(Rust ${describe(r)} vs TS null)`);
      return;
    }

    if ([r.kind, t.kind].some((k) => k === 'union' || k === 'literal' || k === 'pattern')) return walkUnion(r, t, ctx);
    if ((r.kind === 'array' || r.kind === 'tuple') && t.kind === 'tuple') return walkTuple(r, t, ctx);
    if (r.kind === 'tuple' && t.kind === 'array') {
      r.items.forEach((item, i) => walk(item, t.item, { ...ctx, field: `${ctx.field}[${i}]`, path: `${ctx.path}[${i}]` }));
      return;
    }
    if (r.kind === 'map' && t.kind === 'object' && !t.index && t.props().size > 0 && [...t.props().values()].every((p) => p.optional)) {
      // A Rust map keyed by a closed vocabulary is typed as
      // `Partial<Record<Key, V>>` in TS: every value must match.
      for (const [key, p] of t.props()) walk(r.value, p.shape, { ...ctx, field: `${ctx.field}{${key}}`, path: `${ctx.path}.${key}` });
      return;
    }
    if (r.kind !== t.kind) {
      report(ctx.owner, ctx.field, 'type-mismatch', 'incompatible shapes', ctx.path, `(Rust ${describeDeep(r)} vs TS ${describeDeep(t)})`);
      return;
    }
    switch (r.kind) {
      case 'prim':
        if (r.name !== t.name) report(ctx.owner, ctx.field, 'type-mismatch', 'incompatible primitives', ctx.path, `(Rust ${r.name} vs TS ${t.name})`);
        return;
      case 'array':
        return walk(r.item, t.item, { ...ctx, field: `${ctx.field}[]`, path: `${ctx.path}[]` });
      case 'map':
        return walk(r.value, t.value, { ...ctx, field: `${ctx.field}{}`, path: `${ctx.path}{}` });
      case 'object':
        return walkObject(r, t, ctx);
      default:
        return;
    }
  }

  /** Rust fixed-size array or tuple against a TS tuple (with optional and
   * rest elements): accepted lengths and each position must agree. */
  function walkTuple(r, t, ctx) {
    const rest = t.flags.findIndex((f) => f & 12);
    const fixed = rest < 0 ? t.items.length : rest;
    const tMin = t.flags.filter((f) => f & 1).length;
    const tMax = rest < 0 ? t.items.length : Infinity;
    const rMin = r.kind === 'tuple' ? r.items.length : r.minItems ?? 0;
    const rMax = r.kind === 'tuple' ? r.items.length : r.maxItems ?? Infinity;
    if (rMin < tMin || rMax > tMax) {
      report(ctx.owner, ctx.field, 'type-mismatch', 'array lengths differ', ctx.path, `(Rust length ${rMin}..${rMax} vs TS length ${tMin}..${tMax})`);
    }
    if (r.kind === 'array') {
      // Every element of a Rust array has the item type: it must fit every
      // TS position it can occupy.
      t.items.forEach((item, i) => walk(r.item, item, { ...ctx, field: `${ctx.field}[${i}]`, path: `${ctx.path}[${i}]` }));
      return;
    }
    // Known Rust length: fixed TS head, rest middle, fixed TS tail.
    const n = r.items.length;
    const trailing = rest < 0 ? 0 : t.items.length - rest - 1;
    const tAt = (i) => {
      if (i < fixed) return t.items[i];
      if (rest < 0) return null;
      if (i >= n - trailing) return t.items[rest + 1 + (i - (n - trailing))];
      return t.items[rest];
    };
    for (let i = 0; i < n; i++) {
      const ti = tAt(i);
      if (ti) walk(r.items[i], ti, { ...ctx, field: `${ctx.field}[${i}]`, path: `${ctx.path}[${i}]` });
    }
  }

  function walkObject(r, t, ctx, variant) {
    const visitKey = `${objectId(r)}|${objectId(t)}`;
    if (visited.has(visitKey) || OPEN_PAIRS.has(visitKey)) return;
    visited.add(visitKey);
    OPEN_PAIRS.add(visitKey);
    try {
      walkObjectFields(r, t, ctx, variant);
    } finally {
      OPEN_PAIRS.delete(visitKey);
    }
  }

  function walkObjectFields(r, t, ctx, variant) {
    const owner = ownerName(r, t, variant);
    const rp = r.props();
    const tp = t.props();
    for (const [field, rf] of rp) {
      // The discriminant was matched by value when the variant was paired.
      if (variant && field === variant.tag) continue;
      const tf = tp.get(field);
      const where = `${ctx.path}.${field}`;
      if (!tf) {
        if (t.index) walk(rf.shape, t.index, { owner, field, path: where });
        else report(owner, field, 'missing-in-ts', `TS ${t.name} has no "${field}"`, where);
        continue;
      }
      if (rf.optional && !tf.optional) report(owner, field, 'rust-optional', `Rust may omit it; TS ${t.name} requires it`, where);
      else if (!rf.optional && tf.optional) report(owner, field, 'ts-optional', `Rust always emits it; TS ${t.name} marks it optional`, where);
      walk(rf.shape, tf.shape, { owner, field, path: where });
    }
    for (const [field, tf] of tp) {
      if (rp.has(field) || (variant && field === variant.tag)) continue;
      const where = `${ctx.path}.${field}`;
      if (!r.additional) {
        report(owner, field, 'ts-only', `TS ${t.name} declares "${field}"; Rust never emits it`, where);
        continue;
      }
      // Arbitrary extra keys never guarantee a particular key.
      if (!tf.optional) report(owner, field, 'rust-optional', `only Rust's extra keys can carry it; TS ${t.name} requires it`, where);
      walk(r.additional, tf.shape, { owner, field, path: where });
    }
    // String-keyed extra entries (`additionalProperties` / index signature).
    if (r.additional && t.index) walk(r.additional, t.index, { owner, field: '[key]', path: `${ctx.path}[key]` });
    else if (r.additional) report(owner, '[key]', 'missing-in-ts', `Rust emits extra keys; TS ${t.name} has no index signature`, `${ctx.path}[key]`);
    else if (t.index) report(owner, '[key]', 'ts-only', `TS ${t.name} has an index signature; Rust emits only named fields`, `${ctx.path}[key]`);
  }

  function walkUnion(r, t, ctx) {
    const rMembers = r.kind === 'union' ? r.members.map(deref) : [r];
    const tMembers = t.kind === 'union' ? t.members.map(deref) : [t];
    const literalsOf = (members) => new Set(members.filter((m) => m.kind === 'literal').flatMap((m) => [...m.values]));
    const primsOf = (members) => new Set(members.filter((m) => m.kind === 'prim').map((m) => m.name));
    const rLit = literalsOf(rMembers);
    const tLit = literalsOf(tMembers);
    const rPrim = primsOf(rMembers);
    const tPrim = primsOf(tMembers);
    // TS template-literal / `Uppercase<string>` members: string subsets the
    // Rust schema cannot express. Rust strings checked against them are
    // reported as narrowing; Rust string literals are not matched against
    // the pattern and so must be accepted by a plain string or literal.
    const tPatterns = tMembers.filter((m) => m.kind === 'pattern').map((m) => m.text).sort();
    if (tPatterns.length && rPrim.has('string') && !tPrim.has('string')) {
      report(ctx.owner, ctx.field, 'ts-narrower-pattern', 'Rust emits any string; TS narrows it to a pattern', ctx.path, `(${tPatterns.join(' | ')})`);
    }
    if (tPatterns.length && !rPrim.has('string') && ![...rLit].some((v) => typeof v === 'string')) {
      report(ctx.owner, ctx.field, 'variant-never-emitted', 'TS accepts a string pattern Rust never emits', ctx.path, `(TS ${tPatterns.join(' | ')})`);
    }

    // Scalar literals and primitives.
    const missing = [...rLit].filter((v) => !tLit.has(v) && !tPrim.has(scalarKind(v)));
    if (missing.length) report(ctx.owner, ctx.field, 'literal-missing-in-ts', 'Rust emits values TS does not accept', ctx.path, values(missing));
    const extra = [...tLit].filter((v) => !rLit.has(v) && !rPrim.has(scalarKind(v)));
    if (extra.length) report(ctx.owner, ctx.field, 'literal-never-emitted', 'TS accepts values Rust never emits', ctx.path, values(extra));
    for (const p of rPrim) {
      if (tPrim.has(p) || (p === 'string' && tPatterns.length)) continue;
      const narrowed = [...tLit].filter((v) => scalarKind(v) === p);
      if (narrowed.length) report(ctx.owner, ctx.field, 'ts-narrower-literal', `Rust emits any ${p}; TS narrows it to literals`, ctx.path, values(narrowed));
      else report(ctx.owner, ctx.field, 'type-mismatch', 'no TS member accepts this primitive', ctx.path, `(Rust ${p} vs TS ${describeUnion(tMembers)})`);
    }
    for (const p of tPrim) {
      if (rPrim.has(p)) continue;
      const emitted = [...rLit].filter((v) => scalarKind(v) === p);
      if (emitted.length) {
        report(ctx.owner, ctx.field, 'ts-wider-primitive', `Rust emits a closed set; TS accepts any ${p}`, ctx.path, `(Rust ${values(emitted)} vs TS ${p})`);
      } else {
        report(ctx.owner, ctx.field, 'variant-never-emitted', 'TS accepts a primitive Rust never emits', ctx.path, `(TS ${p})`);
      }
    }

    // Objects: internally tagged unions pair by discriminant, others by
    // compatibility; a TS alternative no Rust member uses is reported.
    const rObjects = rMembers.filter((m) => m.kind === 'object');
    const tObjects = tMembers.filter((m) => m.kind === 'object');
    const matched = new Set();
    const seenTags = new Set();
    const tag = rObjects.length > 0 ? discriminant(rObjects) : null;
    const suffix = r.name && t.name && /^\w+$/.test(t.name) && t.name !== r.name ? `(ts ${t.name})` : '';
    if (tag) {
      const tByTag = new Map();
      for (const o of tObjects) {
        const p = o.props().get(tag);
        const v = p && deref(p.shape);
        if (v?.kind !== 'literal') continue;
        for (const value of v.values) tByTag.set(value, [...(tByTag.get(value) ?? []), o]);
      }
      for (const o of rObjects) {
        const value = [...deref(o.props().get(tag).shape).values][0];
        const candidates = tByTag.get(value) ?? [];
        const match = bestMatch(o, candidates);
        if (!match) {
          if (candidates.length) report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'type-mismatch', 'no single compatible TS member', ctx.path, `(Rust ${describe(o)} vs ${candidates.length} TS members)`);
          else report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'variant-missing-in-ts', `TS union has no ${tag}=${q(value)} member`, ctx.path);
          continue;
        }
        matched.add(match);
        seenTags.add(value);
        walkObject(o, match, { ...ctx, path: `${ctx.path}<${tag}=${value}>` }, { tag, value, suffix });
      }
      // A TS member accepting several discriminant values is matched once;
      // every value no Rust variant emits is still a TS-only alternative.
      for (const value of tByTag.keys()) {
        if (!seenTags.has(value)) {
          report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'variant-never-emitted', `TS accepts ${tag}=${q(value)}; Rust never emits it`, ctx.path);
        }
      }
    } else {
      for (const o of rObjects) {
        const match = bestMatch(o, tObjects);
        if (!match) {
          report(ctx.owner, ctx.field, 'type-mismatch', 'no single compatible TS object member', ctx.path, `(Rust ${describe(o)} vs TS ${describeUnion(tObjects)})`);
          continue;
        }
        matched.add(match);
        walk(o, match, ctx);
      }
    }
    for (const o of tObjects) {
      if (matched.has(o)) continue;
      const v = tag && deref(o.props().get(tag)?.shape ?? ANY);
      if (v?.kind === 'literal') {
        // Values no Rust variant emits are reported above; this member shares
        // an emitted value with another, better matching TS member.
        const shared = [...v.values].filter((value) => seenTags.has(value));
        if (shared.length) report(ctx.owner, `${ctx.field}<${tag}=${shared.join('|')}>`, 'variant-never-emitted', 'TS accepts an object alternative Rust never emits', ctx.path, `(${describe(o)})`);
      } else {
        report(ctx.owner, `${ctx.field}<object>`, 'variant-never-emitted', 'TS accepts an object alternative Rust never emits', ctx.path, `(${describe(o)})`);
      }
    }

    // Arrays, tuples and maps: pair by compatibility in both directions.
    const structured = (members) => members.filter((m) => ['array', 'tuple', 'map'].includes(m.kind));
    const tStructured = structured(tMembers);
    for (const m of structured(rMembers)) {
      const candidates = tStructured.filter((x) => x.kind === m.kind || (x.kind === 'tuple' && m.kind === 'array') || (x.kind === 'array' && m.kind === 'tuple'));
      const match = bestMatch(m, candidates);
      if (!match) {
        report(ctx.owner, ctx.field, 'type-mismatch', 'no single compatible TS member', ctx.path, `(Rust ${describeDeep(m)} vs TS ${describeUnion(candidates.length ? candidates : tMembers)})`);
        continue;
      }
      matched.add(match);
      walk(m, match, ctx);
    }
    for (const x of tStructured) {
      if (!matched.has(x)) report(ctx.owner, ctx.field, 'variant-never-emitted', 'TS accepts an alternative Rust never emits', ctx.path, `(TS ${describeDeep(x)})`);
    }
  }

  return { walk };
}

/** Allowlist owner: the Rust type, plus the TS type it was compared with when
 * that is a different named type (so one Rust type compared against two TS
 * projections yields separate keys). Inline enum variants have no Rust type
 * name of their own: they are named by enum and tag value. */
function ownerName(r, t, variant) {
  const rust = r.name.replace(/#\d+$/, '');
  if (variant) return `${rust}<${variant.tag}=${variant.value}>${variant.suffix}`;
  return /^\w+$/.test(t.name) && t.name !== rust ? `${rust}(ts ${t.name})` : rust;
}

const values = (list) => `[${[...list].map((v) => JSON.stringify(v)).sort().join(',')}]`;

const objectIds = new WeakMap();
let nextObjectId = 0;
function objectId(o) {
  if (!objectIds.has(o)) objectIds.set(o, nextObjectId++);
  return objectIds.get(o);
}

function discriminant(objects) {
  const first = objects[0].props();
  for (const [key] of first) {
    const values = new Set();
    let ok = true;
    for (const o of objects) {
      const p = o.props().get(key);
      const s = p && deref(p.shape);
      if (!s || s.kind !== 'literal' || s.values.size !== 1 || p.optional) { ok = false; break; }
      const v = [...s.values][0];
      if (values.has(v)) { ok = false; break; }
      values.add(v);
    }
    if (ok) return key;
  }
  return null;
}

function hasPrim(s, name) {
  return (s.kind === 'prim' && s.name === name) || (s.kind === 'union' && s.members.some((m) => hasPrim(deref(m), name)));
}

function withoutPrim(s, name) {
  if (s.kind === 'prim' && s.name === name) return { kind: 'never' };
  if (s.kind !== 'union') return s;
  const members = s.members.filter((m) => !(deref(m).kind === 'prim' && deref(m).name === name));
  return members.length === 1 ? deref(members[0]) : { ...s, members };
}

function hasNull(s) {
  return (s.kind === 'prim' && s.name === 'null') || (s.kind === 'union' && s.members.some((m) => hasNull(deref(m))));
}

function stripNull(s) {
  if (s.kind !== 'union') return s;
  const members = s.members.filter((m) => !hasNull(deref(m)) || deref(m).kind === 'union');
  if (members.length === 0) return prim('null');
  return members.length === 1 ? deref(members[0]) : { kind: 'union', members };
}

function describe(s) {
  s = deref(s);
  if (s.kind === 'prim') return s.name;
  if (s.kind === 'object') return `object ${s.name}`;
  if (s.kind === 'literal') return [...s.values].map(q).join(' | ');
  if (s.kind === 'pattern') return s.text;
  return s.kind;
}

/** Order-independent description of union members. */
function describeUnion(members) {
  return members.map(describeDeep).sort().join(' | ');
}

/** `describe`, but with array/map element kinds spelled out. */
function describeDeep(s) {
  s = deref(s);
  if (s.kind === 'array') return `${describeDeep(s.item)}[]`;
  if (s.kind === 'map') return `{[key]: ${describeDeep(s.value)}}`;
  return describe(s);
}

const q = (v) => JSON.stringify(v);

// --------------------------------- Driver ----------------------------------

function rustSchema(format, schemaFile) {
  if (schemaFile) return JSON.parse(readFileSync(schemaFile, 'utf8'));
  const out = execFileSync(
    'cargo',
    ['run', '--quiet', '--locked', '-p', format.crate, '--features', 'schema', '--example', 'wire-schema'],
    { cwd: repoRoot, encoding: 'utf8', maxBuffer: 256 * 1024 * 1024, stdio: ['ignore', 'pipe', 'inherit'] },
  );
  return JSON.parse(out);
}

/** The schema cannot tell an omitted field that never carries null (e.g.
 * `Option<String>` skipped on `None`) from one that can still serialize null
 * while present (`Some(x)` where `x` itself serializes as null).
 * rustShapes drops null from omittable fields, so this scan parses the model
 * declarations (comment- and string-aware, multi-line types) and requires,
 * for every field with `skip_serializing_if`, a proof that a present value is
 * never null: for `Option<T>` skipped by `Option::is_none` the payload `T`,
 * otherwise the field type itself. Proven null-free are scalars, containers,
 * braced structs and tagged enums declared in the scanned sources, newtypes
 * over such types, and aliases of them. Everything else fails closed. */
export function scanNullProvenance(dirs) {
  const files = dirs.flatMap((dir) => listRust(path.resolve(repoRoot, dir)));
  const items = files.flatMap((file) => parseRustItems(readFileSync(file, 'utf8')).map((item) => ({ ...item, file })));
  const aliases = new Map();
  const genericAliases = new Set();
  // name -> 'braced' | 'enum' | 'unit' | 'ambiguous' | { newtype: inner }
  const decls = new Map();
  const declare = (name, kind) => decls.set(name, decls.has(name) && JSON.stringify(decls.get(name)) !== JSON.stringify(kind) ? 'ambiguous' : kind);
  for (const item of items) {
    if (item.kind === 'type') {
      if (item.generics.length) genericAliases.add(item.name);
      else aliases.set(item.name, item.rhs);
    } else if (/\b(?:untagged|transparent)\b/.test(item.attrs)) declare(item.name, 'ambiguous');
    else if (item.kind === 'enum') declare(item.name, 'enum');
    else if (item.shape === 'unit') declare(item.name, 'unit');
    else if (item.shape === 'tuple') declare(item.name, item.fields.length === 1 ? { newtype: item.fields[0].type } : 'braced');
    else declare(item.name, 'braced');
  }
  for (let changed = true; changed;) {
    changed = false;
    for (const [name, rhs] of aliases) {
      if (!genericAliases.has(name) && (rhs.match(/\w+/g) ?? []).some((t) => genericAliases.has(t))) {
        genericAliases.add(name);
        changed = true;
      }
    }
  }
  const peel = (type) => {
    let t = type.replace(/\s+/g, '');
    for (let m; (m = /^(?:(?:std|alloc)::(?:boxed|sync|rc)::)?(?:Box|Arc|Rc)<(.*)>$/.exec(t));) t = m[1];
    return t;
  };
  /** Why a present value of `type` may serialize as null, or null when it is
   * proven not to. */
  const nullable = (type, generics, depth = 0) => {
    if (depth > 16) return 'unresolved alias/newtype chain';
    const t = peel(type);
    const generic = (t.match(/\w+/g) ?? []).find((token) => genericAliases.has(token));
    if (generic) return `generic alias ${generic}`;
    if (t === '()' || t === '') return 'unit';
    // Floats are assumed finite: serde_json writes NaN/inf as null, and the
    // parsers only emit finite measurements.
    if (/^(?:String|str|&(?:'\w+)?str|bool|char|[iu](?:8|16|32|64|128|size)|f32|f64)$/.test(t)) return null;
    // Containers keep nullable elements in their item schema (never stripped).
    if (/^\[|^\(.+\)$|^(?:(?:std::)?(?:collections::|vec::)?)(?:Vec|VecDeque|BTreeMap|HashMap|BTreeSet|HashSet)</.test(t)) return null;
    const name = /^(?:\w+::)*(\w+)/.exec(t)?.[1];
    if (!name) return `unrecognized type ${t}`;
    if (name === 'Option') return 'Option';
    if (generics.has(name)) return `generic parameter ${name}`;
    if (aliases.has(name)) return nullable(aliases.get(name), generics, depth + 1);
    const decl = decls.get(name);
    if (decl === 'braced' || decl === 'enum') return null;
    if (decl === 'unit') return `unit struct ${name}`;
    if (decl === 'ambiguous') return `untagged/transparent or ambiguous type ${name}`;
    if (decl?.newtype) return nullable(decl.newtype, new Set(), depth + 1);
    return `type ${name} not found in the scanned sources`;
  };
  /** The field type with transparent wrappers peeled and aliases expanded. */
  const outer = (type) => {
    let t = peel(type);
    for (let i = 0; i < 16; i++) {
      const name = /^(?:\w+::)*(\w+)(?:<.*>)?$/.exec(t)?.[1];
      if (!name || !aliases.has(name)) break;
      t = peel(aliases.get(name));
    }
    return t;
  };

  const problems = [];
  for (const item of items) {
    for (const field of item.fields ?? []) {
      if (!/\bskip_serializing_if\b/.test(field.attrs)) continue;
      const where = `${path.relative(repoRoot, item.file)}:${field.line} ${item.name}.${field.name}`;
      // `"path"`, `r"path"` or `r#"path"#`; anything else fails closed.
      const literal = /\bskip_serializing_if\s*=\s*(?:"((?:[^"\\]|\\.)*)"|r(#*)"([\s\S]*?)"\2)/.exec(field.attrs);
      const skip = literal ? (literal[1] ?? literal[3]).trim() : null;
      if (!skip) {
        problems.push(`${where}: cannot read the skip_serializing_if predicate`);
        continue;
      }
      if (!field.type) {
        problems.push(`${where}: cannot parse the skipped field declaration`);
        continue;
      }
      const generics = new Set(item.generics);
      const option = /^(?:(?:std|core)::option::)?Option<(.*)>$/.exec(outer(field.type));
      if (option && skip !== 'Option::is_none') {
        problems.push(`${where}: Option skipped by "${skip}" can still serialize null`);
        continue;
      }
      const why = option ? nullable(option[1], generics) : nullable(field.type, generics);
      if (why) problems.push(`${where}: skipped ${field.type} can serialize null while present (${why})`);
    }
  }
  return problems;
}

/** Minimal Rust item reader for model declarations: strips comments (string
 * and char literal aware) and returns structs, enums (with the fields of
 * their struct variants and the types of tuple variants) and type aliases. */
export function parseRustItems(source) {
  const text = stripComments(source);
  const lineAt = (index) => text.slice(0, index).split('\n').length;
  const items = [];
  const re = /\b(struct|enum|type)\s+([A-Za-z_]\w*)/g;
  for (let m; (m = re.exec(text));) {
    const [, kind, name] = m;
    let i = m.index + m[0].length;
    let generics = [];
    i = skipSpace(text, i);
    if (text[i] === '<') {
      const end = matchClose(text, i);
      generics = splitTopLevel(text.slice(i + 1, end))
        .map((g) => g.trim().split(/[:\s=]/)[0])
        .filter((g) => /^[A-Z]\w*$/.test(g));
      i = end + 1;
    }
    const attrs = leadingAttributes(text, m.index);
    if (kind === 'type') {
      const eq = text.indexOf('=', i);
      const semi = text.indexOf(';', i);
      if (eq < 0 || semi < eq || /[{}]/.test(text.slice(i, eq))) continue;
      items.push({ kind, name, generics, attrs, rhs: text.slice(eq + 1, semi).trim() });
      continue;
    }
    i = skipWhere(text, skipSpace(text, i));
    if (text[i] === ';') {
      items.push({ kind, name, generics, attrs, shape: 'unit', fields: [] });
      continue;
    }
    if (text[i] !== '{' && text[i] !== '(') continue;
    const end = matchClose(text, i);
    const body = text.slice(i + 1, end);
    const base = i + 1;
    if (kind === 'struct') {
      const fields = text[i] === '{' ? namedFields(body, base, lineAt) : tupleFields(body, base, lineAt);
      items.push({ kind, name, generics, attrs, shape: text[i] === '{' ? 'braced' : 'tuple', fields });
    } else {
      const fields = [];
      for (const part of splitTopLevelWithOffsets(body)) {
        const v = part.text.replace(/^(\s*#\[[\s\S]*?\]\s*)*/, (a) => a);
        const brace = indexOfTopLevel(v, '{');
        if (brace >= 0) {
          const close = matchClose(v, brace);
          fields.push(...namedFields(v.slice(brace + 1, close), base + part.offset + brace + 1, lineAt));
        }
      }
      items.push({ kind, name, generics, attrs, shape: 'enum', fields });
    }
    re.lastIndex = end;
  }
  return items;
}

function namedFields(body, base, lineAt) {
  return splitTopLevelWithOffsets(body)
    .filter((part) => part.text.trim())
    .map((part) => {
      const { attrs, rest } = splitAttributes(part.text);
      const decl = /^\s*(?:pub(?:\s*\([^)]*\))?\s+)?([A-Za-z_]\w*)\s*:\s*([\s\S]+?)\s*$/.exec(rest);
      return {
        attrs,
        name: decl?.[1] ?? rest.trim().split(/\s/)[0],
        type: decl ? normalizeType(decl[2]) : null,
        line: lineAt(base + part.offset + (part.text.length - rest.length) + Math.max(0, rest.search(/\S/))),
      };
    });
}

function tupleFields(body, base, lineAt) {
  return splitTopLevelWithOffsets(body)
    .filter((part) => part.text.trim())
    .map((part, index) => {
      const { attrs, rest } = splitAttributes(part.text);
      const type = normalizeType(rest.replace(/^\s*pub(?:\s*\([^)]*\))?\s+/, ''));
      return { attrs, name: String(index), type, line: lineAt(base + part.offset) };
    });
}

/** Whitespace-free type text without trailing commas in generic lists. */
function normalizeType(text) {
  return text.replace(/\s+/g, '').replace(/,(?=[>)\]])/g, '');
}

function splitAttributes(text) {
  let i = skipSpace(text, 0);
  let attrs = '';
  while (text[i] === '#' && text[skipSpace(text, i + 1)] === '[') {
    const open = skipSpace(text, i + 1);
    const close = matchClose(text, open);
    attrs += ` ${text.slice(i, close + 1)}`;
    i = skipSpace(text, close + 1);
  }
  return { attrs, rest: text.slice(i) };
}

/** Attributes directly before an item keyword (`#[..] pub struct X`). */
function leadingAttributes(text, index) {
  const before = text.slice(Math.max(0, index - 4000), index);
  const m = /((?:#\[(?:[^\[\]]|\[[^\]]*\])*\]\s*)*)(?:pub(?:\s*\([^)]*\))?\s+)?$/.exec(before);
  return m?.[1] ?? '';
}

function skipSpace(text, i) {
  while (i < text.length && /\s/.test(text[i])) i += 1;
  return i;
}

function skipWhere(text, i) {
  if (!text.startsWith('where', i)) return i;
  let j = i + 5;
  while (j < text.length && text[j] !== '{' && text[j] !== ';' && text[j] !== '(') j += 1;
  return j;
}

/** Index of the bracket closing the one at `open` (string-literal aware). */
function matchClose(text, open) {
  const pairs = { '{': '}', '(': ')', '[': ']', '<': '>' };
  const stack = [pairs[text[open]]];
  for (let i = open + 1; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      i = skipString(text, i);
      continue;
    }
    if (ch === '-' && text[i + 1] === '>') {
      i += 1;
      continue;
    }
    if (ch === stack[stack.length - 1]) {
      stack.pop();
      if (stack.length === 0) return i;
    } else if (pairs[ch] && (ch !== '<' || stack[stack.length - 1] === '>' || stack.length)) {
      stack.push(pairs[ch]);
    }
  }
  return text.length;
}

function skipString(text, i) {
  for (let j = i + 1; j < text.length; j++) {
    if (text[j] === '\\') j += 1;
    else if (text[j] === '"') return j;
  }
  return text.length;
}

function indexOfTopLevel(text, needle) {
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') i = skipString(text, i);
    else if (ch === needle && depth === 0) return i;
    else if ('([<'.includes(ch)) depth += 1;
    else if (')]>'.includes(ch) && !(ch === '>' && text[i - 1] === '-')) depth -= 1;
  }
  return -1;
}

/** Remove `//` and `/* *\/` comments, keeping string and char literals (and
 * line structure) intact. */
function stripComments(source) {
  let out = '';
  for (let i = 0; i < source.length; i++) {
    const ch = source[i];
    const next = source[i + 1];
    if (ch === '/' && next === '/') {
      while (i < source.length && source[i] !== '\n') i += 1;
      out += '\n';
      continue;
    }
    if (ch === '/' && next === '*') {
      let depth = 1;
      i += 2;
      for (; i < source.length && depth > 0; i++) {
        if (source[i] === '/' && source[i + 1] === '*') { depth += 1; i += 1; }
        else if (source[i] === '*' && source[i + 1] === '/') { depth -= 1; i += 1; }
        else if (source[i] === '\n') out += '\n';
      }
      i -= 1;
      continue;
    }
    const raw = /^r(#*)"/.exec(source.slice(i, i + 12));
    if (raw && !/\w/.test(source[i - 1] ?? '')) {
      const close = `"${raw[1]}`;
      const end = source.indexOf(close, i + raw[0].length);
      const stop = end < 0 ? source.length : end + close.length;
      out += source.slice(i, stop);
      i = stop - 1;
      continue;
    }
    if (ch === '"') {
      const end = skipString(source, i);
      out += source.slice(i, end + 1);
      i = end;
      continue;
    }
    if (ch === "'") {
      // Char literal ('a', '\n', '\'') vs lifetime ('a).
      const lit = /^'(?:\\(?:x[0-9a-fA-F]{2}|u\{[0-9a-fA-F]+\}|.)|[^\\'])'/u.exec(source.slice(i, i + 12));
      if (lit) {
        out += lit[0];
        i += lit[0].length - 1;
        continue;
      }
    }
    out += ch;
  }
  return out;
}

/** Split on top-level commas (not inside <>, (), [], {} or strings). */
function splitTopLevelWithOffsets(text) {
  const parts = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') i = skipString(text, i);
    else if ('<([{'.includes(ch)) depth += 1;
    else if ('>)]}'.includes(ch) && !(ch === '>' && text[i - 1] === '-')) depth -= 1;
    else if (ch === ',' && depth === 0) {
      parts.push({ text: text.slice(start, i), offset: start });
      start = i + 1;
    }
  }
  parts.push({ text: text.slice(start), offset: start });
  return parts;
}

function splitTopLevel(text) {
  return splitTopLevelWithOffsets(text).map((p) => p.text);
}

function listRust(dir) {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) return listRust(full);
    return entry.name.endsWith('.rs') ? [full] : [];
  });
}

export function checkFormat(name, {
  schemaFile,
  list = false,
  format = FORMATS[name],
  allowlist = MODEL_WIRE_DRIFT_ALLOWLIST[name],
  widening = MODEL_WIRE_TS_WIDENING_ALLOWLIST[name],
} = {}) {
  const provenance = format.rustSources ? scanNullProvenance(format.rustSources) : [];
  if (provenance.length) {
    throw new Error(`model-wire-drift: null provenance is ambiguous for:\n  ${provenance.join('\n  ')}\n`
      + 'Teach rustShapes to keep null for such fields before allowing them.');
  }
  const findings = compareShapes(rustShapes(rustSchema(format, schemaFile)), tsShapes(format));
  if (list) return { findings, unexpected: findings, stale: [] };
  const allowed = new Set(Object.keys(allowlist ?? {}));
  for (const [category, owners] of Object.entries(widening ?? {})) {
    for (const [owner, fields] of Object.entries(owners)) {
      for (const field of fields) allowed.add(`${owner}.${field}: ${category}`);
    }
  }
  const keys = new Set(findings.map((f) => f.key));
  const unexpected = findings.filter((f) => !allowed.has(f.key));
  const stale = [...allowed].filter((k) => !keys.has(k));
  return { findings, unexpected, stale };
}

function main(argv) {
  const args = [...argv];
  let schemaFile;
  let list = false;
  const names = [];
  while (args.length) {
    const a = args.shift();
    if (a === '--schema') schemaFile = args.shift();
    else if (a === '--list') list = true;
    else names.push(a);
  }
  const selected = names.length ? names : Object.keys(FORMATS);
  let failed = false;
  for (const name of selected) {
    if (!FORMATS[name]) throw new Error(`unknown format ${name}`);
    const { findings, unexpected, stale } = checkFormat(name, { schemaFile, list });
    if (list) {
      for (const f of findings) console.log(`${f.key}\t${f.detail}\t${f.path}`);
      console.log(`${name}: ${findings.length} difference(s)`);
      continue;
    }
    for (const f of unexpected) {
      console.error(`${name}: model wire drift ${f.key}\n    ${f.detail}\n    at ${f.path}`);
    }
    for (const k of stale) {
      console.error(`${name}: stale allowlist entry "${k}" (the difference is gone; remove it)`);
    }
    if (unexpected.length || stale.length) failed = true;
    else console.log(`${name}: model wire types in sync (${findings.length} allowlisted difference(s))`);
  }
  if (failed) {
    console.error('\nFix the Rust or TS model, or (for an intentional difference) add a commented entry to scripts/model-wire-drift-allowlist.mjs.');
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2));
}
