#!/usr/bin/env node
// Structural feature extraction and greedy set cover for the fast private
// self-VRT tier.
//
// The full private self-VRT renders every slide/sheet of every corpus file.
// Most of that work is redundant: slides and sheets inside one file repeat the
// same XML structures. This module reduces each selection unit (one PPTX slide
// or one XLSX sheet) to the set of structural XML features reachable from it,
// then greedily picks units until every feature in the corpus is covered.
//
// A feature is one of
//   `partKind|{ns}parent>{ns}child`      element nesting (root: empty parent)
//   `partKind|{ns}element@{ns}attr`      attribute presence
//   `partKind|{ns}element@{ns}attr=val`  enum-like attribute value
//   `partKind|rel>type`                  relationship type leaving a part
//   `dirKind|*.ext`                      non-XML part (media, fonts, OLE, ...)
// where partKind is the part path with every digit run replaced by `N`, and
// names are namespace-URI qualified (Clark notation) so prefix choice does not
// create spurious features.
//
// Library policy, not an OOXML rule: structural coverage does not capture
// geometry-dependent behavior (text wrapping, autofit, pagination, numeric
// values). The fast tier is a day-to-day screen; the full tier remains the
// release gate.
//
// The emitted selection manifest names private corpus files, so it is written
// under the gitignored `packages/<fmt>/public/private/` tree.

import { createHash } from 'node:crypto';
import { readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, posix, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';
import { readRegularFile } from './checked-read.mjs';

export const FAST_TIER_SCHEMA_VERSION = 1;
/** Package-relative path of the selection manifest (gitignored location). */
export const FAST_TIER_MANIFEST = 'public/private/vrt-fast-tier.json';

const UNIT_KIND = { pptx: 'slide', xlsx: 'sheet' };
const ENUM_VALUE = /^[A-Za-z]{1,24}$/;
// Free-text identifiers (object names, alt text, table/sheet display names)
// match ENUM_VALUE but carry no rendering semantics; their presence remains a
// feature, their values do not.
const FREE_TEXT_ATTRIBUTES = new Set(['name', 'descr', 'title', 'displayName', 'altText']);
const XML_PART = /\.(xml|rels|vml)$/i;

// ── ZIP (central directory, stored/deflate, ZIP64) ───────────────────────────

/** Minimal dependency-free ZIP reader. Returns lowercase part name → entry. */
export function readZip(bytes) {
  const eocdMin = Math.max(0, bytes.length - (22 + 0xffff));
  let eocd = -1;
  for (let offset = bytes.length - 22; offset >= eocdMin; offset -= 1) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) { eocd = offset; break; }
  }
  if (eocd < 0) throw new Error('ZIP end-of-central-directory record is missing');
  let count = bytes.readUInt16LE(eocd + 10);
  let cdOffset = bytes.readUInt32LE(eocd + 16);
  if (count === 0xffff || cdOffset === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || bytes.readUInt32LE(locator) !== 0x07064b50) {
      throw new Error('ZIP64 end-of-central-directory locator is missing');
    }
    const zip64 = Number(bytes.readBigUInt64LE(locator + 8));
    if (bytes.readUInt32LE(zip64) !== 0x06064b50) throw new Error('invalid ZIP64 record');
    count = Number(bytes.readBigUInt64LE(zip64 + 32));
    cdOffset = Number(bytes.readBigUInt64LE(zip64 + 48));
  }
  const entries = new Map();
  let p = cdOffset;
  for (let i = 0; i < count; i += 1) {
    if (bytes.readUInt32LE(p) !== 0x02014b50) throw new Error('invalid ZIP central directory');
    const method = bytes.readUInt16LE(p + 10);
    let compressed = bytes.readUInt32LE(p + 20);
    let uncompressed = bytes.readUInt32LE(p + 24);
    const nameLength = bytes.readUInt16LE(p + 28);
    const extraLength = bytes.readUInt16LE(p + 30);
    const commentLength = bytes.readUInt16LE(p + 32);
    let local = bytes.readUInt32LE(p + 42);
    const name = bytes.toString('utf8', p + 46, p + 46 + nameLength);
    for (let e = p + 46 + nameLength; e + 4 <= p + 46 + nameLength + extraLength;) {
      const id = bytes.readUInt16LE(e);
      const size = bytes.readUInt16LE(e + 2);
      if (id === 0x0001) {
        let f = e + 4;
        if (uncompressed === 0xffffffff) { uncompressed = Number(bytes.readBigUInt64LE(f)); f += 8; }
        if (compressed === 0xffffffff) { compressed = Number(bytes.readBigUInt64LE(f)); f += 8; }
        if (local === 0xffffffff) local = Number(bytes.readBigUInt64LE(f));
      }
      e += 4 + size;
    }
    p += 46 + nameLength + extraLength + commentLength;
    if (name.endsWith('/')) continue;
    entries.set(name.replace(/^\//, '').toLowerCase(), { name, method, compressed, local });
  }
  return {
    names: [...entries.keys()].sort(),
    has: (name) => entries.has(name.toLowerCase()),
    read(name) {
      const entry = entries.get(name.toLowerCase());
      if (!entry) return null;
      if (bytes.readUInt32LE(entry.local) !== 0x04034b50) throw new Error(`invalid ZIP local header: ${entry.name}`);
      const start = entry.local + 30 + bytes.readUInt16LE(entry.local + 26) + bytes.readUInt16LE(entry.local + 28);
      const data = bytes.subarray(start, start + entry.compressed);
      if (entry.method === 0) return data;
      if (entry.method === 8) return inflateRawSync(data);
      throw new Error(`unsupported ZIP compression method ${entry.method}: ${entry.name}`);
    },
  };
}

// ── XML scanner ──────────────────────────────────────────────────────────────

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const ATTRIBUTE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function resolveName(qname, scope, isAttribute) {
  const colon = qname.indexOf(':');
  const prefix = colon < 0 ? '' : qname.slice(0, colon);
  const local = colon < 0 ? qname : qname.slice(colon + 1);
  if (prefix === '' && isAttribute) return local;
  const uri = prefix === 'xml' ? XML_NS : scope.get(prefix);
  if (uri === undefined) return prefix === '' ? local : `{?${prefix}}${local}`;
  return uri === '' ? local : `{${uri}}${local}`;
}

/** Tolerant streaming XML scanner with namespace resolution. Calls
 * `onStart(element, stack)` where element = { name, local, attrs: [[name, value]] }
 * and stack holds the open ancestors (excluding element), then `onEnd(stack)`.
 * `onText(raw)` receives raw character data when provided. Mismatched end tags
 * (legacy VML) close up to the nearest matching ancestor. */
export function scanXml(xml, { onStart, onEnd, onText }) {
  const stack = [];
  const scopes = [new Map()];
  let p = 0;
  while (p < xml.length) {
    const lt = xml.indexOf('<', p);
    if (lt < 0) break;
    if (onText && lt > p) onText(xml.slice(p, lt));
    if (xml.startsWith('<!--', lt)) { p = skipTo(xml, '-->', lt); continue; }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt);
      if (onText) onText(xml.slice(lt + 9, end < 0 ? xml.length : end));
      p = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<?', lt)) { p = skipTo(xml, '?>', lt); continue; }
    if (xml.startsWith('<!', lt)) { p = skipTo(xml, '>', lt); continue; }
    let gt = lt + 1;
    let quote = '';
    for (; gt < xml.length; gt += 1) {
      const ch = xml[gt];
      if (quote) { if (ch === quote) quote = ''; } else if (ch === '"' || ch === "'") quote = ch;
      else if (ch === '>') break;
    }
    const body = xml.slice(lt + 1, gt);
    p = gt + 1;
    if (body.startsWith('/')) {
      const qname = body.slice(1).trim();
      let match = stack.length - 1;
      while (match >= 0 && stack[match].qname !== qname) match -= 1;
      if (match < 0) continue;
      while (stack.length > match) {
        stack.pop();
        scopes.pop();
        onEnd?.(stack);
      }
      continue;
    }
    const selfClosing = body.endsWith('/');
    const inner = selfClosing ? body.slice(0, -1) : body;
    const nameEnd = inner.search(/[\s]/);
    const qname = nameEnd < 0 ? inner : inner.slice(0, nameEnd);
    const raw = [];
    let scope = scopes[scopes.length - 1];
    let copied = false;
    ATTRIBUTE.lastIndex = 0;
    for (let m; (m = ATTRIBUTE.exec(inner));) {
      const [, name, dq, sq] = m;
      const value = dq ?? sq;
      if (name === 'xmlns' || name.startsWith('xmlns:')) {
        if (!copied) { scope = new Map(scope); copied = true; }
        scope.set(name === 'xmlns' ? '' : name.slice(6), value);
      } else raw.push([name, value]);
    }
    const name = resolveName(qname, scope, false);
    const colon = qname.indexOf(':');
    const element = {
      qname,
      name,
      local: colon < 0 ? qname : qname.slice(colon + 1),
      attrs: raw.map(([attr, value]) => [resolveName(attr, scope, true), value]),
    };
    onStart(element, stack);
    if (selfClosing) {
      onEnd?.(stack);
    } else {
      stack.push(element);
      scopes.push(scope);
    }
  }
}

function skipTo(xml, token, from) {
  const end = xml.indexOf(token, from);
  return end < 0 ? xml.length : end + token.length;
}

// ── Features ─────────────────────────────────────────────────────────────────

export function partKind(path) {
  return path.toLowerCase().replace(/\d+/g, 'N');
}

function elementFeatures(kind, element, parent, add) {
  add(`${kind}|${parent ? parent.name : ''}>${element.name}`);
  for (const [attr, value] of element.attrs) {
    add(`${kind}|${element.name}@${attr}`);
    if (ENUM_VALUE.test(value) && !FREE_TEXT_ATTRIBUTES.has(attr)) add(`${kind}|${element.name}@${attr}=${value}`);
  }
}

/** Features of one XML part. `bucketOf(element, stack)` may route features of
 * an element subtree to a named bucket (e.g. one `cellXfs/xf`); everything else
 * lands in bucket ''. Returns Map bucket → Set<feature>. */
function xmlPartFeatures(path, xml, bucketOf = () => '', hooks = {}) {
  const kind = partKind(path);
  const buckets = new Map();
  const bucketStack = [];
  scanXml(xml, {
    onStart(element, stack) {
      const inherited = bucketStack.length > 0 ? bucketStack[bucketStack.length - 1] : '';
      const bucket = inherited || bucketOf(element, stack);
      let set = buckets.get(bucket);
      if (!set) buckets.set(bucket, (set = new Set()));
      elementFeatures(kind, element, stack[stack.length - 1], (f) => set.add(f));
      bucketStack.push(bucket);
      hooks.onStart?.(element, stack);
    },
    onEnd(stack) {
      bucketStack.pop();
      hooks.onEnd?.(stack);
    },
    onText: hooks.onText,
  });
  return buckets;
}

function binaryFeature(path) {
  const slash = path.lastIndexOf('/');
  const dot = path.lastIndexOf('.');
  const ext = dot > slash ? path.slice(dot + 1).toLowerCase() : '';
  return `${partKind(path.slice(0, slash + 1))}|*.${ext}`;
}

// ── Package relationships ────────────────────────────────────────────────────

function relsPathOf(part) {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

function relationships(zip, part, xmlOf) {
  const xml = xmlOf(relsPathOf(part));
  if (xml === null) return [];
  const base = part.includes('/') ? part.slice(0, part.lastIndexOf('/')) : '';
  const rels = [];
  scanXml(xml, {
    onStart(element) {
      if (element.local !== 'Relationship') return;
      const attrs = Object.fromEntries(element.attrs);
      if (attrs.TargetMode === 'External' || !attrs.Target) {
        rels.push({ id: attrs.Id, type: attrs.Type ?? '', target: null });
        return;
      }
      let target = attrs.Target.replace(/#.*$/, '');
      try { target = decodeURI(target); } catch { /* keep the raw target */ }
      target = target.startsWith('/')
        ? posix.normalize(target.slice(1))
        : posix.normalize(posix.join(base, target));
      rels.push({ id: attrs.Id, type: attrs.Type ?? '', target: target.toLowerCase() });
    },
  });
  return rels;
}

const relType = (type) => type.slice(type.lastIndexOf('/') + 1);
const isRelationshipAttr = (name) => /\/relationships\}id$/.test(name);

// ── Per-format unit discovery ────────────────────────────────────────────────

function officeDocumentPart(zip, xmlOf) {
  const main = relationships(zip, '', xmlOf).find((rel) => relType(rel.type) === 'officeDocument');
  if (!main?.target) throw new Error('package has no officeDocument relationship');
  return main.target;
}

/** Ordered unit roots, matching the renderer's slide/sheet indices:
 * PPTX `p:sldIdLst/p:sldId` order (ECMA-376 §19.2.1.33) and XLSX
 * `workbook/sheets/sheet` order (§18.2.19), chartsheets included. A malformed
 * reference keeps its slot with no root part so indices never shift. */
function unitRoots(format, zip, xmlOf) {
  const main = officeDocumentPart(zip, xmlOf);
  const byId = new Map(relationships(zip, main, xmlOf).map((rel) => [rel.id, rel.target]));
  const listName = format === 'pptx' ? 'sldIdLst' : 'sheets';
  const itemName = format === 'pptx' ? 'sldId' : 'sheet';
  const roots = [];
  scanXml(xmlOf(main) ?? '', {
    onStart(element, stack) {
      // Depth-anchored: PowerPoint sections (p14:sectionLst) nest their own
      // p14:sldIdLst/p14:sldId inside extLst and must not add slots.
      if (stack.length !== 2 || element.local !== itemName || stack[1].local !== listName) return;
      const id = element.attrs.find(([name]) => isRelationshipAttr(name))?.[1];
      const target = byId.get(id);
      roots.push(target && zip.has(target) ? target : null);
    },
  });
  return { main, roots };
}

function descend(format, rel, from, root) {
  const type = relType(rel.type);
  if (format === 'pptx') {
    // A master lists every layout; only the slide's own layout belongs to it.
    if (type === 'slideLayout' && from !== root) return false;
  }
  return true;
}

// ── XLSX shared parts: route styles/shared strings to the referencing sheet ──

const STYLE_CONTAINERS = new Set(['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'dxfs']);

function styleBuckets(path, xml) {
  const counters = new Map();
  return xmlPartFeatures(path, xml, (element, stack) => {
    if (stack.length !== 2 || !STYLE_CONTAINERS.has(stack[1].local)) return '';
    const container = stack[1].local;
    if (container === 'numFmts') {
      return `numFmt#${element.attrs.find(([name]) => name === 'numFmtId')?.[1]}`;
    }
    const index = counters.get(container) ?? 0;
    counters.set(container, index + 1);
    return `${container}#${index}`;
  });
}

function styleIndex(xml) {
  // xf index → referenced font/fill/border/numFmt/cellStyleXf ids.
  const xfs = { cellXfs: [], cellStyleXfs: [] };
  scanXml(xml, {
    onStart(element, stack) {
      if (element.local !== 'xf' || stack.length !== 2) return;
      const list = xfs[stack[1].local];
      if (list) list.push(Object.fromEntries(element.attrs));
    },
  });
  return xfs;
}

function styleClosure(styles, xfIds, dxfIds) {
  const buckets = new Set();
  const addXf = (list, index) => {
    const xf = styles.index[list][index];
    if (!xf) return;
    buckets.add(`${list}#${index}`);
    if (xf.fontId !== undefined) buckets.add(`fonts#${xf.fontId}`);
    if (xf.fillId !== undefined) buckets.add(`fills#${xf.fillId}`);
    if (xf.borderId !== undefined) buckets.add(`borders#${xf.borderId}`);
    if (xf.numFmtId !== undefined) buckets.add(`numFmt#${xf.numFmtId}`);
    if (list === 'cellXfs' && xf.xfId !== undefined) addXf('cellStyleXfs', Number(xf.xfId));
  };
  for (const id of xfIds) addXf('cellXfs', id);
  for (const id of dxfIds) buckets.add(`dxfs#${id}`);
  return buckets;
}

// ── File analysis ────────────────────────────────────────────────────────────

/** Analyze one package. Returns { unitCount, units: Set<feature>[], fileLevel:
 * Set<feature> }. File-level features (parts no unit reaches, plus unreferenced
 * style/shared-string entries) are covered by whichever unit of the file is
 * selected first. */
export function analyzePackage(format, bytes) {
  const zip = readZip(bytes);
  const xmlCache = new Map();
  const xmlOf = (path) => {
    if (!xmlCache.has(path)) {
      const data = zip.read(path);
      xmlCache.set(path, data === null ? null : data.toString('utf8'));
    }
    return xmlCache.get(path);
  };
  const featureCache = new Map();
  const partFeatures = (path) => {
    if (!featureCache.has(path)) {
      let set;
      if (!XML_PART.test(path)) set = new Set([binaryFeature(path)]);
      else {
        set = xmlPartFeatures(path, xmlOf(path) ?? '').get('') ?? new Set();
        const kind = partKind(path);
        const source = path.replace(/_rels\/([^/]+)\.rels$/, '$1');
        if (path.endsWith('.rels')) {
          for (const rel of relationships(zip, source === path ? '' : source, xmlOf)) {
            set.add(`${kind}|rel>${rel.type}`);
          }
        }
      }
      featureCache.set(path, set);
    }
    return featureCache.get(path);
  };

  const { main, roots } = unitRoots(format, zip, xmlOf);
  const rootSet = new Set(roots.filter(Boolean));
  const shared = new Set();
  let styles = null;
  let sharedStrings = null;
  if (format === 'xlsx') {
    for (const rel of relationships(zip, main, xmlOf)) {
      if (!rel.target || !zip.has(rel.target)) continue;
      const type = relType(rel.type);
      if (type === 'styles') {
        const xml = xmlOf(rel.target);
        styles = { path: rel.target, buckets: styleBuckets(rel.target, xml), index: styleIndex(xml) };
        shared.add(rel.target);
      } else if (type === 'sharedStrings') {
        let index = -1;
        sharedStrings = {
          path: rel.target,
          buckets: xmlPartFeatures(rel.target, xmlOf(rel.target), (element, stack) =>
            (stack.length === 1 && element.local === 'si' ? `si#${(index += 1)}` : '')),
        };
        shared.add(rel.target);
      }
    }
  }

  const reached = new Set();
  const usedBuckets = { styles: new Set(), sharedStrings: new Set() };
  const units = roots.map((root) => {
    const features = new Set();
    if (!root) return features;
    const visited = new Set([root]);
    const queue = [root];
    const xfIds = new Set([0]);
    const dxfIds = new Set();
    const stringIds = new Set();
    while (queue.length > 0) {
      const part = queue.shift();
      reached.add(part);
      for (const f of partFeatures(part)) features.add(f);
      const relsPath = relsPathOf(part);
      if (zip.has(relsPath)) {
        reached.add(relsPath.toLowerCase());
        for (const f of partFeatures(relsPath.toLowerCase())) features.add(f);
      }
      if (format === 'xlsx' && part === root) collectSheetStyleRefs(xmlOf(part), xfIds, dxfIds, stringIds);
      for (const rel of relationships(zip, part, xmlOf)) {
        const target = rel.target;
        if (!target || visited.has(target) || !zip.has(target)) continue;
        if (rootSet.has(target) || shared.has(target) || !descend(format, rel, part, root)) continue;
        visited.add(target);
        queue.push(target);
      }
    }
    if (styles) {
      for (const bucket of styleClosure(styles, xfIds, dxfIds)) {
        usedBuckets.styles.add(bucket);
        for (const f of styles.buckets.get(bucket) ?? []) features.add(f);
      }
    }
    if (sharedStrings) {
      for (const id of stringIds) {
        usedBuckets.sharedStrings.add(`si#${id}`);
        for (const f of sharedStrings.buckets.get(`si#${id}`) ?? []) features.add(f);
      }
    }
    return features;
  });

  const fileLevel = new Set();
  for (const name of zip.names) {
    if (reached.has(name) || shared.has(name)) continue;
    for (const f of partFeatures(name)) fileLevel.add(f);
  }
  for (const [key, part] of [['styles', styles], ['sharedStrings', sharedStrings]]) {
    if (!part) continue;
    for (const [bucket, set] of part.buckets) {
      if (bucket !== '' && usedBuckets[key].has(bucket)) continue;
      for (const f of set) fileLevel.add(f);
    }
  }
  return { unitCount: roots.length, units, fileLevel };
}

/** Cell/row/column style indices, conditional-format dxf ids and shared-string
 * indices referenced by one worksheet (ECMA-376 §18.3.1.4 c@s, §18.3.1.73
 * row@s, §18.3.1.13 col@style, §18.3.1.10 cfRule@dxfId, c@t="s"). */
function collectSheetStyleRefs(xml, xfIds, dxfIds, stringIds) {
  if (!xml) return;
  let sharedCell = false;
  let inValue = false;
  let text = '';
  scanXml(xml, {
    onStart(element) {
      const local = element.local;
      for (const [name, value] of element.attrs) {
        if (name === 's' && (local === 'c' || local === 'row')) xfIds.add(Number(value));
        else if (name === 'style' && local === 'col') xfIds.add(Number(value));
        else if (name === 'dxfId') dxfIds.add(Number(value));
      }
      if (local === 'c') sharedCell = element.attrs.some(([name, value]) => name === 't' && value === 's');
      inValue = local === 'v' && sharedCell;
      text = '';
    },
    onEnd() {
      if (inValue && text.trim() !== '') stringIds.add(Number(text.trim()));
      inValue = false;
    },
    onText(raw) { if (inValue) text += raw; },
  });
}

// ── Greedy cover ─────────────────────────────────────────────────────────────

/** Greedy set cover over all (file, unit) candidates. Each candidate covers its
 * own features plus, while its file has no selected unit yet, the file-level
 * features. Ties prefer the smaller file, then the name, then the lower index.
 * Lazy evaluation is exact here because gains never increase. */
export function selectUnits(files) {
  const ids = new Map();
  const id = (feature) => {
    let value = ids.get(feature);
    if (value === undefined) ids.set(feature, (value = ids.size));
    return value;
  };
  const candidates = [];
  const fileLevel = files.map((file) => Int32Array.from(file.fileLevel, id));
  files.forEach((file, fileIndex) => {
    file.units.forEach((features, unit) => {
      candidates.push({ fileIndex, unit, features: Int32Array.from(features, id), gain: 0 });
    });
  });
  const covered = new Uint8Array(ids.size);
  const fileSelected = new Uint8Array(files.length);
  const gainOf = (candidate) => {
    let gain = 0;
    for (const f of candidate.features) if (!covered[f]) gain += 1;
    if (!fileSelected[candidate.fileIndex]) {
      for (const f of fileLevel[candidate.fileIndex]) if (!covered[f]) gain += 1;
    }
    return gain;
  };
  const before = (a, b) => {
    if (a.gain !== b.gain) return a.gain > b.gain;
    const fa = files[a.fileIndex];
    const fb = files[b.fileIndex];
    if (fa.size !== fb.size) return fa.size < fb.size;
    if (fa.name !== fb.name) return fa.name < fb.name;
    return a.unit < b.unit;
  };
  const heap = new Heap(before);
  for (const candidate of candidates) {
    candidate.gain = gainOf(candidate);
    if (candidate.gain > 0) heap.push(candidate);
  }
  const selected = files.map(() => []);
  while (heap.size > 0) {
    const top = heap.pop();
    top.gain = gainOf(top);
    if (top.gain === 0) continue;
    if (heap.size > 0 && before(heap.peek(), top)) { heap.push(top); continue; }
    selected[top.fileIndex].push(top.unit);
    for (const f of top.features) covered[f] = 1;
    if (!fileSelected[top.fileIndex]) {
      fileSelected[top.fileIndex] = 1;
      for (const f of fileLevel[top.fileIndex]) covered[f] = 1;
    }
  }
  let uncovered = 0;
  for (const value of covered) if (!value) uncovered += 1;
  return { selected: selected.map((units) => units.sort((a, b) => a - b)), features: ids.size, uncovered };
}

class Heap {
  constructor(before) { this.before = before; this.items = []; }
  get size() { return this.items.length; }
  peek() { return this.items[0]; }
  push(item) {
    const items = this.items;
    items.push(item);
    for (let i = items.length - 1; i > 0;) {
      const parent = (i - 1) >> 1;
      if (!this.before(items[i], items[parent])) break;
      [items[i], items[parent]] = [items[parent], items[i]];
      i = parent;
    }
  }
  pop() {
    const items = this.items;
    const top = items[0];
    const last = items.pop();
    if (items.length > 0) {
      items[0] = last;
      for (let i = 0; ;) {
        const l = 2 * i + 1;
        const r = l + 1;
        let best = i;
        if (l < items.length && this.before(items[l], items[best])) best = l;
        if (r < items.length && this.before(items[r], items[best])) best = r;
        if (best === i) break;
        [items[i], items[best]] = [items[best], items[i]];
        i = best;
      }
    }
    return top;
  }
}

// ── Manifest ─────────────────────────────────────────────────────────────────

/** Build the fast-tier selection manifest for corpus files
 * `[{ name, bytes }]` (name as the VRT spec lists it, e.g. `pptx/x.pptx`).
 * The manifest is deterministic (no timestamps), so regenerating an unchanged
 * corpus reproduces it byte for byte. */
export function buildFastTierManifest(format, corpus) {
  if (!UNIT_KIND[format]) throw new Error(`unsupported fast-tier format: ${format}`);
  const files = corpus.map(({ name, bytes }) => ({
    name,
    size: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    ...analyzePackage(format, bytes),
  }));
  const { selected, features, uncovered } = selectUnits(files);
  if (uncovered > 0) {
    // Only possible for a file with zero units; its package-level features
    // cannot be rendered by any slide/sheet.
    throw new Error(`${uncovered} features are not reachable from any ${UNIT_KIND[format]}`);
  }
  return {
    schemaVersion: FAST_TIER_SCHEMA_VERSION,
    format,
    unitKind: UNIT_KIND[format],
    summary: {
      files: files.length,
      features,
      unitsTotal: files.reduce((sum, file) => sum + file.unitCount, 0),
      unitsSelected: selected.reduce((sum, units) => sum + units.length, 0),
    },
    files: files.map((file, index) => ({
      name: file.name,
      sha256: file.sha256,
      unitCount: file.unitCount,
      units: selected[index],
    })),
  };
}

/** Resolve `VRT_PRIVATE_TIER` (unset/`full` → full corpus, `fast` → the
 * selection manifest). Unknown values fail rather than silently running full. */
export function privateCorpusTier(value = process.env.VRT_PRIVATE_TIER) {
  if (value === undefined || value === '' || value === 'full') return 'full';
  if (value === 'fast') return 'fast';
  throw new Error(`invalid VRT_PRIVATE_TIER: ${value} (expected full or fast)`);
}

/** Load the fast-tier selection for the private corpus `files` (as the VRT
 * spec lists them, relative to `directory`). Fails closed when the manifest is
 * missing, stale for any input, or does not list exactly the current corpus:
 * a silently skipped document is not coverage. Returns Map name → { unitCount,
 * units } and the manifest digest that binds fast-tier baselines. */
export function loadFastTierSelection({
  format, files, manifestPath = FAST_TIER_MANIFEST, directory = 'public/private',
}) {
  const regenerate = `regenerate it with \`node tests/visual/corpus-features.mjs --format ${format}\``
    + ' (or `pnpm vrt:private:fast:select`)';
  let text;
  try {
    text = readRegularFile(manifestPath).toString('utf8');
  } catch {
    throw new Error(`missing ${format} fast-tier selection manifest ${manifestPath}; ${regenerate}`);
  }
  const manifest = JSON.parse(text);
  if (manifest.schemaVersion !== FAST_TIER_SCHEMA_VERSION || manifest.format !== format) {
    throw new Error(`incompatible ${format} fast-tier selection manifest ${manifestPath}; ${regenerate}`);
  }
  const entries = new Map(manifest.files.map((entry) => [entry.name, entry]));
  const problems = [];
  const digests = new Map();
  for (const name of files) {
    const entry = entries.get(name);
    if (!entry) { problems.push(`${name} is not in the manifest`); continue; }
    const sha256 = createHash('sha256').update(readRegularFile(join(directory, name))).digest('hex');
    if (entry.sha256 !== sha256) problems.push(`${name} changed since selection`);
    digests.set(name, entry.sha256);
  }
  for (const name of entries.keys()) {
    if (!files.includes(name)) problems.push(`${name} is no longer in the corpus`);
  }
  if (problems.length > 0) {
    throw new Error(`stale ${format} fast-tier selection manifest: ${problems.join('; ')}; ${regenerate}`);
  }
  return {
    digest: createHash('sha256').update(text).digest('hex'),
    // `sha256` is the digest validated above; corpus preparation must read
    // the same bytes.
    files: new Map(manifest.files.map(({ name, unitCount, units }) => [name, { unitCount, units, sha256: digests.get(name) }])),
  };
}

/** Corpus files exactly as the private VRT specs enumerate them. */
export function listCorpus(directory, format) {
  return readdirSync(directory)
    .filter((file) => file.endsWith(`.${format}`) && !file.startsWith('~$'))
    .filter((file) => statSync(join(directory, file)).isFile())
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

function main(argv) {
  const args = new Map();
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith('--')) args.set(argv[i].slice(2), argv[i + 1]?.startsWith('--') ? '' : argv[++i] ?? '');
  }
  const format = args.get('format');
  if (!UNIT_KIND[format]) {
    console.error('usage: node tests/visual/corpus-features.mjs --format pptx|xlsx [--corpus DIR] [--out FILE]');
    process.exit(2);
  }
  const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
  const directory = resolve(args.get('corpus') || join(repo, `packages/${format}/public/private/${format}`));
  const out = resolve(args.get('out') || join(repo, `packages/${format}/`, FAST_TIER_MANIFEST));
  const started = performance.now();
  const corpus = listCorpus(directory, format).map((file) => ({
    name: `${format}/${file}`,
    bytes: readRegularFile(join(directory, file)),
  }));
  const manifest = buildFastTierManifest(format, corpus);
  writeFileSync(out, `${JSON.stringify(manifest, null, 2)}\n`);
  const { summary } = manifest;
  console.log(
    `${format}: ${summary.unitsSelected}/${summary.unitsTotal} ${manifest.unitKind}s selected `
    + `covering ${summary.features} features across ${summary.files} files `
    + `in ${((performance.now() - started) / 1000).toFixed(1)}s`,
  );
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2));
}
