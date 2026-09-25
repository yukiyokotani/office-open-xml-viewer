import { describe, expect, it } from 'vitest';
import { mathMLToSvg } from './engine.js';

describe('MathJax engine outside Window', () => {
  it('typesets through LiteDOM when document is unavailable', async () => {
    expect(typeof document).toBe('undefined');

    const output = await mathMLToSvg(
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi><mo>+</mo><mn>1</mn></math>',
    );

    expect(output.svg).toContain('<svg');
    expect(output.widthEm).toBeGreaterThan(0);
    expect(output.ascentEm).toBeGreaterThan(0);
  });

  it('keeps representative equations inside the worker Canvas SVG vocabulary', async () => {
    const equations = [
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>',
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></mrow></msqrt></math>',
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfenced><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></mfenced></math>',
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mover><mi>x</mi><mo>¯</mo></mover><mo>=</mo><munderover><mo>∑</mo><mn>1</mn><mi>n</mi></munderover><msub><mi>x</mi><mi>i</mi></msub></math>',
    ];
    const names = new Set<string>();
    for (const equation of equations) {
      const output = await mathMLToSvg(equation);
      for (const match of output.svg.matchAll(/<\/?([A-Za-z][\w:-]*)/g)) {
        names.add(match[1].toLowerCase());
      }
    }

    expect([...names].sort()).toEqual(['g', 'path', 'rect', 'svg']);
  });

  it('renders accented Latin glyphs as STIX2 paths in every OMML math style', async () => {
    for (const variant of ['normal', 'italic', 'bold', 'bold-italic']) {
      const output = await mathMLToSvg(
        `<math xmlns="http://www.w3.org/1998/Math/MathML"><mi mathvariant="${variant}">á</mi></math>`,
      );
      expect(output.svg).toMatch(/<path[^>]*data-c="E1"/);
      expect(output.svg).not.toContain('<text');
    }
  });

  it('uses a synchronous fallback for an excluded range on repeated conversion', async () => {
    const equation = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>Ж</mi></math>';
    for (let attempt = 0; attempt < 2; attempt++) {
      const output = await mathMLToSvg(equation);
      expect(output.svg).toContain('>Ж</text>');
    }
  });
});
