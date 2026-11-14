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

import { MODEL_WIRE_DRIFT_ALLOWLIST } from './model-wire-drift-allowlist.mjs';

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
      return union(members);
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
          return { kind: 'tuple', items: node.prefixItems.map((p) => convert(p)) };
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
  if (nonNull.length > 0) {
    if (nonNull.every((v) => typeof v === 'string')) members.push({ kind: 'literal', values: new Set(nonNull) });
    else members.push(prim(typeof nonNull[0] === 'boolean' ? 'boolean' : typeof nonNull[0]));
  }
  if (nonNull.length !== values.length) members.push(prim('null'));
  return union(members);
}

function union(members) {
  return lazy(() => unionResolved(members));
}

function unionResolved(members) {
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
  return out.length === 1 ? deref(out[0]) : { kind: 'union', members: out };
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
      const members = type.types.filter((t) => !(t.flags & (ts.TypeFlags.Undefined | ts.TypeFlags.Void)));
      const bools = members.filter((t) => t.flags & ts.TypeFlags.BooleanLiteral);
      const rest = members.filter((t) => !(t.flags & ts.TypeFlags.BooleanLiteral));
      const shapes = rest.map(convert);
      if (bools.length > 0) shapes.push(prim('boolean'));
      // `'a' | 'b' | string` widens to string in TS already.
      return union(shapes);
    }
    if (f & ts.TypeFlags.StringLiteral) return { kind: 'literal', values: new Set([type.value]) };
    if (f & (ts.TypeFlags.String | ts.TypeFlags.TemplateLiteral | ts.TypeFlags.StringMapping)) return prim('string');
    if (f & (ts.TypeFlags.Number | ts.TypeFlags.NumberLiteral | ts.TypeFlags.Enum | ts.TypeFlags.EnumLiteral)) return prim('number');
    if (f & (ts.TypeFlags.Boolean | ts.TypeFlags.BooleanLiteral)) return prim('boolean');
    if (f & ts.TypeFlags.Null) return prim('null');
    if (checker.isTupleType(type)) {
      return { kind: 'tuple', items: checker.getTypeArguments(type).map(convert) };
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
            props.set(symbol.getName(), {
              shape: convert(propType),
              optional: (symbol.flags & ts.SymbolFlags.Optional) !== 0,
            });
          }
          return props;
        }),
      };
    }
    return ANY;
  }

  return convert(rootType);
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

function comparer(sink) {
  const visited = new Set();
  const report = (owner, field, category, detail, where, payload) => {
    const key = `${owner}.${field}: ${category}${payload ? ` ${payload}` : ''}`;
    sink({ key, category, detail, path: where });
  };

  function walk(r, t, ctx) {
    r = deref(r);
    t = deref(t);
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
    if (hasNull(r) && !hasNull(t)) {
      report(ctx.owner, ctx.field, 'rust-null', 'Rust can emit null; the TS type does not allow null', ctx.path);
    }
    r = stripNull(r);
    t = stripNull(t);
    if (r.kind === 'prim' && r.name === 'null') return;
    if (t.kind === 'prim' && t.name === 'null') {
      report(ctx.owner, ctx.field, 'type-mismatch', 'TS accepts only null', ctx.path, `(Rust ${describe(r)} vs TS null)`);
      return;
    }

    if (r.kind === 'union' || t.kind === 'union') return walkUnion(r, t, ctx);
    if (r.kind === 'literal' && t.kind === 'literal') return compareLiterals(r.values, t.values, ctx);
    if (r.kind === 'literal' && t.kind === 'prim' && t.name === 'string') return;
    if (r.kind === 'prim' && r.name === 'string' && t.kind === 'literal') {
      report(ctx.owner, ctx.field, 'ts-narrower-literal', 'Rust emits any string; TS narrows it to literals', ctx.path, values(t.values));
      return;
    }
    if (r.kind === 'array' && t.kind === 'tuple') {
      // Fixed-size Rust arrays (`[f64; 4]`) are JSON arrays typed as TS tuples.
      if (r.minItems !== t.items.length || r.maxItems !== t.items.length) {
        report(ctx.owner, ctx.field, 'type-mismatch', 'array length differs', ctx.path, `(Rust array ${r.minItems ?? 0}..${r.maxItems ?? 'n'} vs TS tuple ${t.items.length})`);
      }
      t.items.forEach((item, i) => walk(r.item, item, { ...ctx, path: `${ctx.path}[${i}]` }));
      return;
    }
    if (r.kind === 'map' && t.kind === 'object' && !t.index && [...t.props().values()].every((p) => p.optional)) {
      // A Rust map keyed by a closed vocabulary is typed as
      // `Partial<Record<Key, V>>` in TS: every value must match.
      for (const [key, p] of t.props()) walk(r.value, p.shape, { ...ctx, path: `${ctx.path}.${key}` });
      return;
    }
    if (r.kind !== t.kind) {
      report(ctx.owner, ctx.field, 'type-mismatch', 'incompatible shapes', ctx.path, `(Rust ${describe(r)} vs TS ${describe(t)})`);
      return;
    }
    switch (r.kind) {
      case 'prim':
        if (r.name !== t.name) report(ctx.owner, ctx.field, 'type-mismatch', 'incompatible primitives', ctx.path, `(Rust ${r.name} vs TS ${t.name})`);
        return;
      case 'array':
        return walk(r.item, t.item, { ...ctx, path: `${ctx.path}[]` });
      case 'map':
        return walk(r.value, t.value, { ...ctx, path: `${ctx.path}{}` });
      case 'tuple':
        if (r.items.length !== t.items.length) {
          report(ctx.owner, ctx.field, 'type-mismatch', 'tuple arity differs', ctx.path, `(Rust ${r.items.length} vs TS ${t.items.length})`);
          return;
        }
        r.items.forEach((item, i) => walk(item, t.items[i], { ...ctx, path: `${ctx.path}[${i}]` }));
        return;
      case 'object':
        return walkObject(r, t, ctx);
      default:
        return;
    }
  }

  function compareLiterals(rValues, tValues, ctx) {
    const missing = [...rValues].filter((v) => !tValues.has(v));
    const extra = [...tValues].filter((v) => !rValues.has(v));
    if (missing.length) report(ctx.owner, ctx.field, 'literal-missing-in-ts', 'Rust emits values the TS literal union lacks', ctx.path, values(missing));
    if (extra.length) report(ctx.owner, ctx.field, 'literal-never-emitted', 'TS accepts values Rust never emits', ctx.path, values(extra));
  }

  function walkObject(r, t, ctx, variant) {
    const visitKey = `${objectId(r)}|${objectId(t)}`;
    if (visited.has(visitKey)) return;
    visited.add(visitKey);
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
      walk(rf.shape, tf.shape, { owner, field, path: where });
    }
    for (const field of tp.keys()) {
      if (rp.has(field) || (variant && field === variant.tag)) continue;
      if (r.additional) walk(r.additional, tp.get(field).shape, { owner, field, path: `${ctx.path}.${field}` });
      else report(owner, field, 'ts-only', `TS ${t.name} declares "${field}"; Rust never emits it`, `${ctx.path}.${field}`);
    }
    // String-keyed extra entries (`additionalProperties` / index signature).
    if (r.additional && t.index) walk(r.additional, t.index, { owner, field: '[key]', path: `${ctx.path}[key]` });
    else if (r.additional) report(owner, '[key]', 'missing-in-ts', `Rust emits extra keys; TS ${t.name} has no index signature`, `${ctx.path}[key]`);
    else if (t.index) report(owner, '[key]', 'ts-only', `TS ${t.name} has an index signature; Rust emits only named fields`, `${ctx.path}[key]`);
  }

  function walkUnion(r, t, ctx) {
    const rMembers = r.kind === 'union' ? r.members.map(deref) : [r];
    const tMembers = t.kind === 'union' ? t.members.map(deref) : [t];
    const rLit = new Set(rMembers.filter((m) => m.kind === 'literal').flatMap((m) => [...m.values]));
    const tLit = new Set(tMembers.filter((m) => m.kind === 'literal').flatMap((m) => [...m.values]));
    const rString = rMembers.some((m) => m.kind === 'prim' && m.name === 'string');
    const tString = tMembers.some((m) => m.kind === 'prim' && m.name === 'string');
    if (rLit.size && tLit.size) compareLiterals(rLit, tLit, ctx);
    else if (rLit.size && !tString) {
      report(ctx.owner, ctx.field, 'literal-missing-in-ts', 'Rust emits string literals TS does not accept', ctx.path, values(rLit));
    } else if (tLit.size && !rString) {
      report(ctx.owner, ctx.field, 'literal-never-emitted', 'TS accepts literals Rust never emits', ctx.path, values(tLit));
    }
    if (rString && !tString && tLit.size) {
      report(ctx.owner, ctx.field, 'ts-narrower-literal', 'Rust emits any string; TS narrows it to literals', ctx.path, values(tLit));
    }

    const rObjects = rMembers.filter((m) => m.kind === 'object');
    const tObjects = tMembers.filter((m) => m.kind === 'object');
    const tag = rObjects.length > 0 ? discriminant(rObjects) : null;
    if (tag) {
      const tByTag = new Map();
      for (const o of tObjects) {
        const p = o.props().get(tag);
        const v = p && deref(p.shape);
        if (v?.kind !== 'literal') {
          report(ctx.owner, `${ctx.field}<${tag}=?>`, 'variant-never-emitted', `TS member ${o.name} has no literal "${tag}"`, ctx.path, `(${o.name})`);
          continue;
        }
        for (const value of v.values) tByTag.set(value, [...(tByTag.get(value) ?? []), o]);
      }
      const seen = new Set();
      for (const o of rObjects) {
        const value = [...deref(o.props().get(tag).shape).values][0];
        seen.add(value);
        const candidates = tByTag.get(value) ?? [];
        const match = bestMatch(o, candidates);
        if (!match) {
          if (candidates.length) report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'type-mismatch', 'no single compatible TS member', ctx.path, `(Rust ${describe(o)} vs ${candidates.length} TS members)`);
          else report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'variant-missing-in-ts', `TS union has no ${tag}=${q(value)} member`, ctx.path);
          continue;
        }
        walkObject(o, match, { ...ctx, path: `${ctx.path}<${tag}=${value}>` }, { tag, value });
      }
      for (const value of tByTag.keys()) {
        if (!seen.has(value)) report(ctx.owner, `${ctx.field}<${tag}=${value}>`, 'variant-never-emitted', `TS union accepts ${tag}=${q(value)}; Rust never emits it`, ctx.path);
      }
    } else {
      for (const o of rObjects) {
        const match = bestMatch(o, tObjects);
        if (!match) report(ctx.owner, ctx.field, 'type-mismatch', 'no single compatible TS object member', ctx.path, `(Rust ${describe(o)} vs TS ${describeUnion(tObjects)})`);
        else walk(o, match, ctx);
      }
    }
    // Remaining Rust members (primitives except string, arrays, maps, tuples):
    // pair each with its most compatible TS member of the same kind.
    for (const m of rMembers.filter((x) => !['object', 'literal'].includes(x.kind))) {
      if (m.kind === 'prim' && m.name === 'string' && (tString || tLit.size)) continue;
      const candidates = tMembers.filter((x) => x.kind === m.kind && (m.kind !== 'prim' || x.name === m.name));
      const match = bestMatch(m, candidates);
      if (!match) report(ctx.owner, ctx.field, 'type-mismatch', 'no single compatible TS member', ctx.path, `(Rust ${describeDeep(m)} vs TS ${describeUnion(candidates.length ? candidates : tMembers)})`);
      else if (m.kind !== 'prim') walk(m, match, ctx);
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
  if (variant) return `${rust}<${variant.tag}=${variant.value}>`;
  return /^\w+$/.test(t.name) && t.name !== rust ? `${rust}(ts ${t.name})` : rust;
}

const values = (list) => JSON.stringify([...list].sort());

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

/** The schema cannot tell `Option<T>` skipped on `None` (never null) from
 * `Option<Option<T>>` or an Option skipped by another predicate (both can
 * emit null). Scan the model sources and report every such field so the
 * checker's null handling stays sound. */
export function scanNullProvenance(dirs) {
  const problems = [];
  for (const file of dirs.flatMap((dir) => listRust(path.resolve(repoRoot, dir)))) {
    const lines = readFileSync(file, 'utf8').split('\n');
    let attrs = '';
    let depth = 0;
    lines.forEach((line, i) => {
      const trimmed = line.trim();
      if (depth > 0 || trimmed.startsWith('#[')) {
        attrs += ` ${trimmed}`;
        for (const ch of trimmed) depth += ch === '[' ? 1 : ch === ']' ? -1 : 0;
        return;
      }
      if (trimmed.startsWith('//') || trimmed === '') return;
      const field = /^(?:pub(?:\([^)]*\))?\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(.+?),?$/.exec(trimmed);
      if (field && !trimmed.startsWith('let ')) {
        const type = field[2].replace(/\s+/g, '');
        const skip = /skip_serializing_if\s*=\s*"([^"]+)"/.exec(attrs)?.[1];
        const where = `${path.relative(repoRoot, file)}:${i + 1} ${field[1]}`;
        if (/Option<Option</.test(type)) problems.push(`${where}: Option<Option<..>> can serialize null`);
        else if (skip && type.startsWith('Option<') && skip !== 'Option::is_none') {
          problems.push(`${where}: Option skipped by "${skip}" can still serialize null`);
        }
      }
      attrs = '';
    });
  }
  return problems;
}

function listRust(dir) {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) return listRust(full);
    return entry.name.endsWith('.rs') ? [full] : [];
  });
}

export function checkFormat(name, { schemaFile, list = false, format = FORMATS[name], allowlist = MODEL_WIRE_DRIFT_ALLOWLIST[name] } = {}) {
  const provenance = format.rustSources ? scanNullProvenance(format.rustSources) : [];
  if (provenance.length) {
    throw new Error(`model-wire-drift: null provenance is ambiguous for:\n  ${provenance.join('\n  ')}\n`
      + 'Teach rustShapes to keep null for such fields before allowing them.');
  }
  const findings = compareShapes(rustShapes(rustSchema(format, schemaFile)), tsShapes(format));
  if (list) return { findings, unexpected: findings, stale: [] };
  const entries = Object.keys(allowlist ?? {}).map((key) => ({ key, re: allowPattern(key) }));
  const unexpected = findings.filter((f) => !entries.some((e) => e.re.test(f.key)));
  const stale = entries.filter((e) => !findings.some((f) => e.re.test(f.key))).map((e) => e.key);
  return { findings, unexpected, stale };
}

/** Allowlist keys are exact finding keys; `*` matches any run of characters
 * (used only for a whole family of differences with one shared reason). */
function allowPattern(key) {
  return new RegExp(`^${key.split('*').map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('.*')}$`);
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
