// Structural validity of a built package: OPC content types and relationships
// (ECMA-376 Part 2) and XML 1.0 well-formedness with namespace prefixes bound.
// This is a guard against builder mistakes, not a schema validator.

import { posix } from 'node:path';
import { readZip } from '../../visual/corpus-features.mjs';

const NAME = /^[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?$/;
const ATTR = /\s+([^\s=]+)\s*=\s*("[^"<]*"|'[^'<]*')/;
const PREDEFINED = new Set(['lt', 'gt', 'amp', 'quot', 'apos']);

/** Strict well-formedness check. Returns the parsed elements as
 * [{ name, attrs: Map, ns: Map(prefix → uri), depth }] in document order. */
export function parseXmlStrict(xml) {
  const elements = [];
  const stack = [];
  const scopes = [new Map([['xml', 'http://www.w3.org/XML/1998/namespace']])];
  let roots = 0;
  let p = 0;
  const fail = (message) => { throw new Error(`${message} at offset ${p}`); };
  const checkText = (text) => {
    if (/[<]/.test(text)) fail('unescaped <');
    for (const m of text.matchAll(/&([^;\s]*);?/g)) {
      const ref = m[1];
      if (!m[0].endsWith(';') || !(PREDEFINED.has(ref) || /^#(\d+|x[\da-fA-F]+)$/.test(ref))) fail(`bad entity &${ref}`);
    }
  };
  if (xml.startsWith('<?xml')) p = xml.indexOf('?>') + 2;
  while (p < xml.length) {
    const lt = xml.indexOf('<', p);
    const text = xml.slice(p, lt < 0 ? xml.length : lt);
    if (stack.length === 0 && text.trim()) fail('text outside the root element');
    checkText(text);
    if (lt < 0) break;
    p = lt;
    if (xml.startsWith('<!--', p)) { p = xml.indexOf('-->', p) + 3; continue; }
    if (xml.startsWith('<?', p) || xml.startsWith('<!', p)) fail('unexpected declaration');
    const gt = xml.indexOf('>', p);
    if (gt < 0) fail('unterminated tag');
    if (xml[p + 1] === '/') {
      const name = xml.slice(p + 2, gt).trim();
      const open = stack.pop();
      if (open !== name) fail(`</${name}> closes <${open}>`);
      scopes.pop();
      p = gt + 1;
      continue;
    }
    const selfClosing = xml[gt - 1] === '/';
    const body = xml.slice(p + 1, selfClosing ? gt - 1 : gt);
    const name = /^[^\s/>]+/.exec(body)?.[0] ?? '';
    if (!NAME.test(name)) fail(`bad element name ${name}`);
    const attrs = new Map();
    const re = new RegExp(ATTR.source, 'y');
    re.lastIndex = name.length;
    let end = name.length;
    for (let m; (m = re.exec(body)); end = re.lastIndex) {
      if (!NAME.test(m[1]) && !/^xmlns(:[A-Za-z_][\w.-]*)?$/.test(m[1])) fail(`bad attribute name ${m[1]}`);
      if (attrs.has(m[1])) fail(`duplicate attribute ${m[1]}`);
      checkText(m[2].slice(1, -1));
      attrs.set(m[1], m[2].slice(1, -1));
    }
    if (body.slice(end).trim()) fail(`malformed attributes in <${name}>: ${body.slice(end).trim()}`);
    const ns = new Map(scopes[scopes.length - 1]);
    for (const [attr, value] of attrs) {
      if (attr === 'xmlns') ns.set('', value);
      else if (attr.startsWith('xmlns:')) ns.set(attr.slice(6), value);
    }
    for (const qname of [name, ...[...attrs.keys()].filter((a) => !a.startsWith('xmlns'))]) {
      const colon = qname.indexOf(':');
      if (colon > 0 && !ns.has(qname.slice(0, colon))) fail(`unbound prefix in ${qname}`);
    }
    if (stack.length === 0 && ++roots > 1) fail('more than one root element');
    elements.push({ name, attrs, ns, depth: stack.length });
    if (!selfClosing) { stack.push(name); scopes.push(ns); }
    p = gt + 1;
  }
  if (stack.length) throw new Error(`unclosed <${stack[stack.length - 1]}>`);
  if (roots !== 1) throw new Error('no root element');
  return elements;
}

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const sourceOfRels = (rels) => rels.replace(/(^|\/)_rels\/([^/]*)\.rels$/, '$1$2');

/** Throws on the first structural problem of the package bytes. */
export function checkPackage(bytes) {
  const zip = readZip(bytes);
  const names = new Set(zip.names);
  const xmlOf = (name) => zip.read(name).toString('utf8');
  const parsed = new Map();
  for (const name of names) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    try { parsed.set(name, parseXmlStrict(xmlOf(name))); } catch (e) { throw new Error(`${name}: ${e.message}`); }
  }

  // Content types: every part has one; every override names a part.
  const types = parsed.get('[content_types].xml');
  if (!types) throw new Error('missing [Content_Types].xml');
  const defaults = new Map();
  const overrides = new Map();
  for (const el of types) {
    if (el.name === 'Default') defaults.set(el.attrs.get('Extension').toLowerCase(), el.attrs.get('ContentType'));
    if (el.name === 'Override') overrides.set(el.attrs.get('PartName').replace(/^\//, '').toLowerCase(), el.attrs.get('ContentType'));
  }
  for (const part of overrides.keys()) if (!names.has(part)) throw new Error(`override for missing part ${part}`);
  for (const name of names) {
    if (name === '[content_types].xml') continue;
    const ext = name.slice(name.lastIndexOf('.') + 1);
    if (!overrides.has(name) && !defaults.has(ext)) throw new Error(`no content type for ${name}`);
  }

  // Relationships: targets exist, ids are unique, r:* references resolve, and
  // every part other than the content types and rels parts is a target.
  const targeted = new Set();
  const idsBySource = new Map();
  for (const [name, elements] of parsed) {
    if (!name.endsWith('.rels')) continue;
    const source = sourceOfRels(name);
    if (source !== '' && !names.has(source)) throw new Error(`${name} belongs to missing part ${source}`);
    const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/')) : '';
    const ids = new Set();
    for (const el of elements) {
      if (el.name !== 'Relationship') continue;
      const id = el.attrs.get('Id');
      if (ids.has(id)) throw new Error(`${name}: duplicate relationship id ${id}`);
      ids.add(id);
      if (el.attrs.get('TargetMode') === 'External') continue;
      const target = posix.normalize(posix.join(base, el.attrs.get('Target'))).toLowerCase();
      if (!names.has(target)) throw new Error(`${name}: ${id} targets missing part ${target}`);
      targeted.add(target);
    }
    idsBySource.set(source, ids);
  }
  const root = parsed.get('_rels/.rels') ?? [];
  if (!root.some((el) => el.name === 'Relationship' && el.attrs.get('Type').endsWith('/officeDocument'))) {
    throw new Error('package has no officeDocument relationship');
  }
  for (const name of names) {
    if (name === '[content_types].xml' || name.endsWith('.rels')) continue;
    if (!targeted.has(name)) throw new Error(`part ${name} is not the target of any relationship`);
  }
  for (const [name, elements] of parsed) {
    if (name.endsWith('.rels') || name === '[content_types].xml') continue;
    const ids = idsBySource.get(name) ?? new Set();
    for (const el of elements) {
      for (const [attr, value] of el.attrs) {
        const colon = attr.indexOf(':');
        if (colon < 0 || el.ns.get(attr.slice(0, colon)) !== REL_NS) continue;
        if (!ids.has(value)) throw new Error(`${name}: <${el.name} ${attr}="${value}"> has no relationship`);
      }
    }
  }
}
