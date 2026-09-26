// Source experiments only: never infer an Office result from this manifest.
import { pathToFileURL } from 'node:url';
import { rulerProbeManifest } from './legacy-ppt-ruler-probes.mjs';

export function textProbeManifest() {
  const original = rulerProbeManifest();
  const baseline = original.conditions[0].first;
  const paragraphs = texts => texts.map(text => ({ runs: [{ text }] }));
  const body = (properties, content) => ({ properties: structuredClone(properties), paragraphs: structuredClone(content) });
  const conditions = original.conditions.map(c => ({
    id: c.id.replace('T', 'S'), family: 'separate-rulers', sourceCondition: c.id,
    name: c.name,
    bodies: [body(c.first, paragraphs([original.text])), body(c.second, paragraphs([original.text]))],
  }));
  const slots = [
    ['Baseline text', paragraphs(['ABCD'])],
    ['Equal length different text', paragraphs(['WXYZ'])],
    ['Literal spaces', paragraphs([' A  '])],
    ['Literal underscores', paragraphs(['____'])],
    ['Tab in run', paragraphs(['A\tB'])],
    ['Trailing tab', paragraphs(['AB\t'])],
    ['Combining sequence', paragraphs(['Ae\u0301B'])],
    ['Supplementary character', paragraphs(['A\u{1f600}B'])],
    ['CJK characters', paragraphs(['A漢字B'])],
    ['Empty paragraph boundaries', paragraphs(['', 'AB', ''])],
    ['Unequal paragraph lengths', paragraphs(['A', 'BCD'])],
    ['Reversed paragraph lengths', paragraphs(['BCD', 'A'])],
    ['Styled run boundary', [{ runs: [{ text: 'AB' }, { text: 'CD', bold: true }] }]],
    ['Explicit line break', [{ runs: [{ text: 'AB' }, { break: true }, { text: 'CD' }] }]],
    ['Empty run', [{ runs: [{ text: 'AB' }, { text: '' }, { text: 'CD' }] }]],
    ['Unchanged text repeat', paragraphs(['ABCD'])],
  ];
  for (const [index, [name, content]] of slots.entries()) {
    conditions.push({
      id: `M${String(index + 1).padStart(3, '0')}`, family: 'text-slots', name,
      bodies: [body(baseline, paragraphs(['ABCD'])), body(baseline, content)],
    });
  }
  return {
    schemaVersion: 1, units: 'EMU', arrangement: 'two-separate-text-bodies',
    origin: 'synthetic authored OOXML; not an observed Office result',
    conditions,
  };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.stdout.write(`${JSON.stringify(textProbeManifest(), null, 2)}\n`);
}
