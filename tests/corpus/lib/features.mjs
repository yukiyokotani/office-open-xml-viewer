// Feature vocabulary of the corpus `covers` lists.
//
// Extraction is delegated to tests/visual/corpus-features.mjs, the extractor
// the private fast self-VRT tier uses, so a case's `covers` and the private
// corpus speak the same vocabulary:
//   partKind|{ns}parent>{ns}child
//   partKind|{ns}element@{ns}attr
//   partKind|{ns}element@{ns}attr=value   (enum-like values only)
// partKind is the lower-cased part path with digit runs replaced by N.
//
// Case modules write features with prefixes for readability
// ('word/document.xml|w:tab@w:val=decimal'); `canonicalFeature` expands them
// with the fixed prefix map in xml.mjs.

import { analyzePackage, partKind } from '../../visual/corpus-features.mjs';
import { NS } from './xml.mjs';

function qualify(qname, isAttribute) {
  const colon = qname.indexOf(':');
  if (colon < 0) {
    if (isAttribute) return qname;
    throw new Error(`element name needs a prefix: ${qname}`);
  }
  const prefix = qname.slice(0, colon);
  const uri = NS[prefix];
  if (!uri) throw new Error(`unknown prefix in feature: ${prefix}`);
  return `{${uri}}${qname.slice(colon + 1)}`;
}

export function canonicalFeature(shorthand) {
  const bar = shorthand.indexOf('|');
  if (bar < 0) throw new Error(`feature needs 'part|selector': ${shorthand}`);
  const kind = partKind(shorthand.slice(0, bar).trim());
  const selector = shorthand.slice(bar + 1).trim();
  const at = selector.indexOf('@');
  if (at >= 0) {
    const element = qualify(selector.slice(0, at), false);
    const rest = selector.slice(at + 1);
    const eq = rest.indexOf('=');
    const attr = qualify(eq < 0 ? rest : rest.slice(0, eq), true);
    return `${kind}|${element}@${attr}${eq < 0 ? '' : `=${rest.slice(eq + 1)}`}`;
  }
  const gt = selector.indexOf('>');
  if (gt < 0) throw new Error(`feature selector needs '>' or '@': ${shorthand}`);
  const parent = selector.slice(0, gt);
  return `${kind}|${parent ? qualify(parent, false) : ''}>${qualify(selector.slice(gt + 1), false)}`;
}

/** Every feature present anywhere in the package. analyzePackage splits
 * features into slide/sheet units and file-level leftovers; the union is the
 * whole package (for DOCX there are no units, so everything is file-level). */
export function packageFeatures(format, bytes) {
  const { units, fileLevel } = analyzePackage(format, bytes);
  const all = new Set(fileLevel);
  for (const unit of units) for (const feature of unit) all.add(feature);
  return all;
}
