#!/usr/bin/env node
// Extracts ECMA-376 §20.1.9.19 preset TEXT-WARP definitions from the spec's
// presetTextWarpDefinitions.xml into the same compact JSON shape the shared
// preset-geometry evaluator/path-executor already consume for preset SHAPES.
//
// A text warp's <pathLst> is NOT a closed silhouette: it is a set of open
// "envelope" paths (typically a top edge + a bottom edge) that the renderer
// maps glyphs between. The command/guide-formula grammar is identical to a
// preset shape's, so the extracted `{ adj, gd, paths }` records feed the very
// same evaluator (createEvaluator) and path flattener.
//
// Usage (from packages/pptx):
//   node scripts/extract-text-warps.mjs scripts/preset-text-warp-definitions.xml \
//     ../core/src/shape/preset-geometry/text-warp-presets.json

import fs from 'node:fs';

const [, , inPath, outPath] = process.argv;
if (!inPath || !outPath) {
  console.error('Usage: extract-text-warps.mjs <input.xml> <output.json>');
  process.exit(1);
}

const xml = fs.readFileSync(inPath, 'utf8');

// The spec ships this file as a single line; strip the XML declaration and the
// outer <presetTextWarpDefinitions> wrapper, then walk its direct children —
// each is one warp preset whose immediate contents open with <avLst>.
const inner = /<presetTextWarpDefinitions>([\s\S]*)<\/presetTextWarpDefinitions>/.exec(xml);
if (!inner) {
  console.error('Could not find <presetTextWarpDefinitions> root');
  process.exit(1);
}

const out = {};
// Each warp block: <name><avLst …>…</name>. The `(?=<avLst)` lookahead anchors
// on the first child so we never match a nested element as a top-level preset.
const shapeRe = /<([a-zA-Z][a-zA-Z0-9]*)>(?=<avLst)([\s\S]*?)<\/\1>/g;
let m;
while ((m = shapeRe.exec(inner[1])) !== null) {
  out[m[1].toLowerCase()] = parseWarp(m[2]);
}

fs.writeFileSync(outPath, JSON.stringify(out));
console.error(`Wrote ${Object.keys(out).length} text-warp presets → ${outPath}`);

function parseWarp(body) {
  return {
    adj: extractGuides(extractBlock(body, 'avLst')),
    gd: extractGuides(extractBlock(body, 'gdLst')),
    paths: extractPaths(extractBlock(body, 'pathLst')),
  };
}

function extractBlock(body, tag) {
  const re = new RegExp(`<${tag}\\b[^>]*(?:/>|>([\\s\\S]*?)</${tag}>)`);
  const mm = body.match(re);
  return mm ? (mm[1] ?? '') : '';
}

function extractGuides(block) {
  const res = [];
  // The spec file writes each guide as <gd …></gd> (not self-closing).
  const re = /<gd\s+name="([^"]+)"\s+fmla="([^"]+)"\s*(?:\/>|><\/gd>)/g;
  let mm;
  while ((mm = re.exec(block)) !== null) res.push([mm[1], mm[2]]);
  return res;
}

function extractPaths(block) {
  const res = [];
  const re = /<path\b([^>]*)>([\s\S]*?)<\/path>/g;
  let mm;
  while ((mm = re.exec(block)) !== null) {
    const attrs = parseAttrs(mm[1]);
    res.push({
      w: attrs.w ? +attrs.w : null,
      h: attrs.h ? +attrs.h : null,
      fill: attrs.fill ?? null,
      stroke: attrs.stroke !== 'false',
      extrusionOk: attrs.extrusionOk !== 'false',
      cmds: extractCommands(mm[2]),
    });
  }
  return res;
}

function parseAttrs(s) {
  const res = {};
  const re = /(\w+)="([^"]*)"/g;
  let mm;
  while ((mm = re.exec(s)) !== null) res[mm[1]] = mm[2];
  return res;
}

function extractCommands(s) {
  const res = [];
  const tokenRe =
    /<(moveTo|lnTo|arcTo|cubicBezTo|quadBezTo|close)\b([^>]*?)(?:\s*\/>|>([\s\S]*?)<\/\1>)/g;
  let mm;
  while ((mm = tokenRe.exec(s)) !== null) {
    const type = mm[1];
    const attrs = parseAttrs(mm[2] ?? '');
    const inner2 = mm[3] ?? '';
    if (type === 'close') {
      res.push(['c']);
    } else if (type === 'arcTo') {
      res.push(['a', attrs.wR, attrs.hR, attrs.stAng, attrs.swAng]);
    } else {
      const pts = [];
      const ptRe = /<pt\s+x="([^"]+)"\s+y="([^"]+)"\s*(?:\/>|><\/pt>)/g;
      let pm;
      while ((pm = ptRe.exec(inner2)) !== null) pts.push([pm[1], pm[2]]);
      const code =
        type === 'moveTo' ? 'm' : type === 'lnTo' ? 'l' : type === 'cubicBezTo' ? 'C' : 'Q';
      res.push([code, ...pts.flat()]);
    }
  }
  return res;
}
