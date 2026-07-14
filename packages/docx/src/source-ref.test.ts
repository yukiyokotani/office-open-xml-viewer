import { describe, expect, it } from 'vitest';
import { buildSegments, type LayoutTextSeg } from './line-layout.js';
import type { DocRun, DocxTextRun } from './types.js';

const ENV = { pageIndex: 0, totalPages: 1 };

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: 'serif',
    isLink: false,
    background: null,
    vertAlign: null,
    allCaps: false,
    smallCaps: false,
    doubleStrikethrough: false,
    ...extra,
  } as unknown as DocRun;
}

describe('DOCX text source references', () => {
  it('clips one XML text-node range across visual script segments', () => {
    const sourceRefs = [{
      partName: 'word/document.xml',
      path: [{ namespaceUri: 'urn:w', localName: 't', index: 0 }],
      textStart: 0,
      textEnd: 4,
      sourceStart: 0,
      sourceEnd: 4,
    }];
    const segments = buildSegments(
      [textRun('A😀中', { sourceRefs })],
      ENV,
    ).filter((segment): segment is LayoutTextSeg => 'text' in segment);

    expect(segments.map((segment) => segment.text).join('')).toBe('A😀中');
    expect(segments.flatMap((segment) => segment.sourceRefs ?? []).map((ref) => [
      ref.sourceStart,
      ref.sourceEnd,
    ])).toEqual([[0, 3], [3, 4]]);
    for (const segment of segments) {
      expect(segment.sourceRefs).toEqual([expect.objectContaining({
        textStart: 0,
        textEnd: segment.text.length,
      })]);
    }
  });
});
