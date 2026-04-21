#!/usr/bin/env node
// Extracts ECMA-376 preset shape definitions from LibreOffice's
// presetShapeDefinitions.xml into a compact JSON the runtime engine consumes.
//
// Usage: node scripts/extract-presets.mjs <input.xml> <output.json>

import fs from 'node:fs';

const [, , inPath, outPath] = process.argv;
if (!inPath || !outPath) {
  console.error('Usage: extract-presets.mjs <input.xml> <output.json>');
  process.exit(1);
}

const xml = fs.readFileSync(inPath, 'utf8');

// ── Parse each top-level <shapeName>…</shapeName> block ─────────────────────
const out = {};

// Shape block regex: matches opening tag at col 2 + closing tag at col 2.
// The LibreOffice file uses 2-space indentation at top level.
const shapeRe = /^  <([a-zA-Z][a-zA-Z0-9]*)>\s*\n([\s\S]*?)^  <\/\1>/gm;
let m;
while ((m = shapeRe.exec(xml)) !== null) {
  const name = m[1];
  // Skip action button placeholder names used for nested elements, not shapes.
  if (name === 'pathLst' || name === 'avLst' || name === 'gdLst' ||
      name === 'rect' || name === 'ahLst' || name === 'cxnLst') continue;
  out[name.toLowerCase()] = parseShape(m[2]);
}

fs.writeFileSync(outPath, JSON.stringify(out));
console.error(`Wrote ${Object.keys(out).length} presets → ${outPath}`);

function parseShape(body) {
  return {
    adj: extractGuides(extractBlock(body, 'avLst')),
    gd:  extractGuides(extractBlock(body, 'gdLst')),
    paths: extractPaths(extractBlock(body, 'pathLst')),
  };
}

function extractBlock(body, tag) {
  // <tag ...>…</tag> or <tag ... />
  const re = new RegExp(`<${tag}\\b[^>]*(?:/>|>([\\s\\S]*?)</${tag}>)`);
  const m = body.match(re);
  return m ? (m[1] ?? '') : '';
}

function extractGuides(block) {
  const out = [];
  const re = /<gd\s+name="([^"]+)"\s+fmla="([^"]+)"\s*\/>/g;
  let m;
  while ((m = re.exec(block)) !== null) out.push([m[1], m[2]]);
  return out;
}

function extractPaths(block) {
  const out = [];
  const re = /<path\b([^>]*)>([\s\S]*?)<\/path>/g;
  let m;
  while ((m = re.exec(block)) !== null) {
    const attrs = parseAttrs(m[1]);
    out.push({
      w:    attrs.w    ? +attrs.w    : null,
      h:    attrs.h    ? +attrs.h    : null,
      fill: attrs.fill ?? null,                              // null = "norm"
      stroke: attrs.stroke !== 'false',                      // default true
      extrusionOk: attrs.extrusionOk !== 'false',
      cmds: extractCommands(m[2]),
    });
  }
  return out;
}

function parseAttrs(s) {
  const out = {};
  const re = /(\w+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(s)) !== null) out[m[1]] = m[2];
  return out;
}

function extractCommands(s) {
  const out = [];
  // Walk through commands in order. Match whichever comes first.
  const tokenRe = /<(moveTo|lnTo|arcTo|cubicBezTo|quadBezTo|close)\b([^>]*?)(?:\s*\/>|>([\s\S]*?)<\/\1>)/g;
  let m;
  while ((m = tokenRe.exec(s)) !== null) {
    const type = m[1];
    const attrs = parseAttrs(m[2] ?? '');
    const inner = m[3] ?? '';
    if (type === 'close') {
      out.push(['c']);
    } else if (type === 'arcTo') {
      out.push(['a', attrs.wR, attrs.hR, attrs.stAng, attrs.swAng]);
    } else {
      const pts = [];
      const ptRe = /<pt\s+x="([^"]+)"\s+y="([^"]+)"\s*\/>/g;
      let pm;
      while ((pm = ptRe.exec(inner)) !== null) pts.push([pm[1], pm[2]]);
      const code = type === 'moveTo' ? 'm' : type === 'lnTo' ? 'l'
                 : type === 'cubicBezTo' ? 'C' : 'Q';
      out.push([code, ...pts.flat()]);
    }
  }
  return out;
}
