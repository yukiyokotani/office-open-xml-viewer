import { describe, expect, it } from 'vitest';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import { wordExternalLinkSyntaxBreakOffsets } from './layout/line-compatibility.js';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';

const breaks = (text: string, protectedOffsets: readonly number[] = []) =>
  wordExternalLinkSyntaxBreakOffsets(
    text,
    new Set(graphemeClusterOffsets(text)),
    new Set(protectedOffsets),
  );

describe('Word-observed external-link syntax breaks', () => {
  it('keeps a hyphenated authority intact and exposes path/query separators', () => {
    const text = 'https://hyphen-host.example/path/long-name.pdf?part=one&part=two';
    const authorityHyphen = text.indexOf('-') + 1;
    const pathSlash = text.indexOf('/', text.indexOf('/path') + 1) + 1;
    const pathHyphen = text.indexOf('-', text.indexOf('/path')) + 1;
    expect(breaks(text)).not.toContain(authorityHyphen);
    expect(breaks(text)).toEqual(expect.arrayContaining([
      pathSlash,
      pathHyphen,
      text.indexOf('?') + 1,
      text.indexOf('&') + 1,
    ]));
  });

  it('does not split a separator from its combining mark', () => {
    const text = 'https://example.test/path/a-\u0301combined/next-part';
    const clusteredHyphen = text.indexOf('-') + 1;
    expect(breaks(text)).not.toContain(clusteredHyphen);
    expect(breaks(text)).toContain(text.lastIndexOf('-') + 1);
  });

  it('excludes an authored noBreakHyphen while retaining other URL breaks', () => {
    const text = 'https://example.test/path/no-break/ordinary-break';
    const protectedHyphen = text.indexOf('-') + 1;
    const ordinaryHyphen = text.lastIndexOf('-') + 1;
    expect(breaks(text, [protectedHyphen])).not.toContain(protectedHyphen);
    expect(breaks(text, [protectedHyphen])).toContain(ordinaryHyphen);
  });

  it('checks every candidate when negative spacing makes prefix widths non-monotone', () => {
    const text = 'https://example.test/path/i-i-W-W-document.pdf';
    const offsets = breaks(text);
    const ctx = {
      font: '10px serif',
      letterSpacing: '0px',
      fontKerning: 'none',
      measureText(value: string) {
        const width = [...value].reduce(
          (sum, char) => sum + (char === 'i' ? 2 : char === 'W' ? 20 : 8),
          0,
        );
        return {
          width,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const segment: LayoutTextSeg = {
      text,
      bold: false,
      italic: false,
      underline: true,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'serif',
      vertAlign: null,
      measuredWidth: 0,
      charSpacing: -5,
      hyperlink: { kind: 'external', url: text },
      externalLinkBreakOffsets: offsets,
      src: { segIndex: 0, charOffset: 0 },
    };

    const lines = layoutLines(ctx, [segment], 94, 0, 1);
    const pieces = lines.flatMap((line) => line.segments)
      .filter((item): item is LayoutTextSeg => 'text' in item);
    expect(pieces.map((piece) => piece.text).join('')).toBe(text);
    expect(pieces.length).toBeGreaterThan(1);
    expect(offsets).toContain(pieces[0]!.text.length);
  });
});
