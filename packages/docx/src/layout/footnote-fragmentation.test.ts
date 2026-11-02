import { describe, expect, it } from 'vitest';
import { createFootnotePartitioner, partitionFootnote } from './footnote-fragmentation.js';
import type { NoteLayout, NoteSeparatorLayout, ParagraphLayout } from './types.js';
import { deepFreezePlainData } from './plain-data.js';

const source = { story: 'footnote' as const, storyInstance: '7', path: [0] };

function note(pageIndex: number): NoteLayout {
  const bounds = { xPt: 0, yPt: 6, widthPt: 180, heightPt: 40 };
  const paragraph = {
    kind: 'paragraph', id: `footnote:7:page:${pageIndex}:paragraph`, source,
    flowDomainId: `notes:page:${pageIndex}:footnote:7`, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: 40,
    spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
    lines: Array.from({ length: 4 }, (_, index) => ({
      range: { start: index, end: index + 1 },
      bounds: { xPt: 0, yPt: 6 + index * 10, widthPt: 20, heightPt: 10 },
      baselinePt: 14 + index * 10, advancePt: 10,
      placements: [{
        kind: 'text' as const, range: { start: index, end: index + 1 },
        origin: { xPt: 0, yPt: 14 + index * 10 },
        bounds: { xPt: 0, yPt: 6 + index * 10, widthPt: 20, heightPt: 10 },
        advancePt: 20, decorations: [],
      }],
    })),
  } as unknown as ParagraphLayout;
  return {
    kind: 'note', id: `footnote:7:page:${pageIndex}`, source,
    flowDomainId: `notes:page:${pageIndex}`, ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt: 46 },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt: 46 },
    advancePt: 46, separator: [],
    story: {
      story: 'footnote', blocks: [paragraph], flowBounds: bounds,
      inkBounds: bounds, advancePt: 40, diagnostics: [],
    },
  };
}

function notice(advancePt: number): NoteSeparatorLayout {
  const bounds = { xPt: 0, yPt: 0, widthPt: 180, heightPt: advancePt };
  const source = { story: 'footnote' as const, storyInstance: 'reserved:continuation-notice', path: [] };
  return deepFreezePlainData({
    role: 'continuationNotice', source, flowBounds: bounds, advancePt,
    paragraph: {
      kind: 'paragraph', id: 'notice', source: { ...source, path: [0] }, flowDomainId: 'notice',
      ordinaryFlow: true, flowBounds: bounds, inkBounds: bounds, advancePt,
      spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false, lines: [],
      borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
    } as unknown as ParagraphLayout,
  }) as NoteSeparatorLayout;
}

describe('page-bottom footnote continuation', () => {
  it('retains an indivisible over-band line and its notice, then resumes every remaining line once', () => {
    const original = note(0);
    const paragraph = original.story.blocks[0] as ParagraphLayout;
    const lines = paragraph.lines.map((line, index) => ({ ...line,
      advancePt: index === 0 ? 120 : 10,
      bounds: { ...line.bounds, yPt: index === 0 ? 6 : 126 + (index - 1) * 10,
        heightPt: index === 0 ? 120 : 10 },
    }));
    const bounds = { ...paragraph.flowBounds, heightPt: 150 };
    const acquired: NoteLayout = { ...original, advancePt: 156,
      flowBounds: { ...original.flowBounds, heightPt: 156 },
      story: { ...original.story, advancePt: 150, blocks: [{ ...paragraph,
        flowBounds: bounds, inkBounds: bounds, advancePt: 150, lines }] },
    };
    expect(partitionFootnote(acquired, null, 80)).toBeNull();
    const head = partitionFootnote(acquired, null, 80, () => notice(4), 'overflow')!;
    expect(head.fragment.advancePt).toBe(130);
    expect(head.fragment.trailing?.advancePt).toBe(4);
    const tail = partitionFootnote(acquired, head.nextCursor, 80)!;
    expect(tail.nextCursor).toBeNull();
    expect([head, tail].flatMap(part => part.fragment.story.blocks.flatMap(block =>
      block.kind === 'paragraph' ? block.lines.map(line => line.range.start) : []))).toEqual([0, 1, 2, 3]);
  });

  it('does not rescan a retained long paragraph for every continuation page', () => {
    const total = 2_000;
    let advanceReads = 0;
    const original = note(0);
    const paragraph = original.story.blocks[0] as ParagraphLayout;
    const line = paragraph.lines[0]!;
    const lines = Array.from({ length: total }, (_, index) => ({
      ...line, range: { start: index, end: index + 1 },
      bounds: { ...line.bounds, yPt: 6 + index * 10 }, baselinePt: 14 + index * 10,
      get advancePt() { advanceReads++; return 10; },
      placements: [{ ...line.placements[0]!, range: { start: index, end: index + 1 },
        sourceRunIndex: 0 }],
    }));
    const bounds = { ...paragraph.flowBounds, heightPt: total * 10 };
    const block = deepFreezePlainData({ ...paragraph, lines, advancePt: total * 10, flowBounds: bounds });
    const acquired = deepFreezePlainData({ ...original, advancePt: total * 10 + 6,
      flowBounds: { ...original.flowBounds, heightPt: total * 10 + 6 },
      story: { ...original.story, advancePt: total * 10, flowBounds: bounds,
        blocks: [block] },
    }) as NoteLayout;
    advanceReads = 0;
    const partition = createFootnotePartitioner();
    let cursor = null;
    const retained: number[] = [];
    do {
      const result = partition(acquired, cursor, 106);
      if (!result) throw new Error('Expected source progress');
      retained.push(...result.fragment.story.blocks.flatMap(block => block.kind === 'paragraph'
        ? block.lines.map(line => line.range.start) : []));
      cursor = result.nextCursor;
    } while (cursor);
    expect(retained).toEqual(Array.from({ length: total }, (_, index) => index));
    // Includes indexing and retained-fragment finalization. This bound detects
    // whole-source work per page without a machine-dependent timing threshold.
    expect(advanceReads).toBeLessThan(total * 25);
  });

  it('retains each source line exactly once across a page boundary', () => {
    const first = partitionFootnote(note(0), null, 27)!;
    expect(first.nextCursor).not.toBeNull();
    expect(first.fragment.story.blocks[0]?.kind).toBe('paragraph');
    expect(first.fragment.advancePt).toBe(26);
    const second = partitionFootnote(note(1), first.nextCursor, 100)!;
    expect(second.nextCursor).toBeNull();
    const ranges = [first.fragment, second.fragment].flatMap((part) =>
      part.story.blocks.flatMap((block) => block.kind === 'paragraph'
        ? block.lines.map((line) => line.range.start) : []));
    expect(ranges).toEqual([0, 1, 2, 3]);
    expect(second.fragment.flowDomainId).toBe('notes:page:1');
  });

  it('keeps fractional line gaps and before/after ownership at a capacity boundary', () => {
    const original = note(0);
    const paragraph = original.story.blocks[0] as ParagraphLayout;
    const block = deepFreezePlainData({ ...paragraph, advancePt: 46,
      flowBounds: { ...paragraph.flowBounds, heightPt: 46 },
      spacing: { beforePt: 2, afterPt: 3 },
      lines: paragraph.lines.map((line, index) => ({ ...line, advancePt: 10.1,
        bounds: { ...line.bounds, yPt: 8 + index * 10.3, heightPt: 10.1 },
      })),
    });
    const acquired = { ...original, advancePt: 52,
      flowBounds: { ...original.flowBounds, heightPt: 52 },
      story: { ...original.story, blocks: [block], advancePt: 46 },
    } as NoteLayout;
    const partition = createFootnotePartitioner();
    const first = partition(acquired, null, 28.4)!;
    expect(first.nextCursor?.lineIndex).toBe(2);
    expect(first.fragment.advancePt).toBeCloseTo(28.4);
    expect(partition(acquired, null, 28.3)!.nextCursor?.lineIndex).toBe(1);
    const last = partition(acquired, first.nextCursor, 29.4)!;
    expect(last.nextCursor).toBeNull();
    expect(last.fragment.advancePt).toBeCloseTo(29.4);
    expect(last.fragment.story.blocks[0]?.kind === 'paragraph'
      && last.fragment.story.blocks[0].spacing).toEqual({ beforePt: 0, afterPt: 3 });
  });

  it('relocates a reference if no note line can fit', () => {
    expect(partitionFootnote(note(0), null, 15)).toBeNull();
  });

  it('rejects a changed continuation width before a line cursor can skip text', () => {
    const first = partitionFootnote(note(0), null, 27)!;
    const second = { ...note(1), flowBounds: { ...note(1).flowBounds, widthPt: 160 } };
    expect(() => partitionFootnote(second, first.nextCursor, 100)).toThrow(/different text widths/);
  });

  it('keeps a fitting note whole', () => {
    const whole = partitionFootnote(note(0), null, 46)!;
    expect(whole.nextCursor).toBeNull();
    expect(whole.fragment.advancePt).toBe(46);
  });

  it('preserves independently positioned paragraphs whole instead of folding their frame as a margin', () => {
    const original = note(0);
    const acquired = { ...original, story: { ...original.story,
      blocks: original.story.blocks.map(block => ({ ...block, ordinaryFlow: false })),
    } } as NoteLayout;
    expect(partitionFootnote(acquired, null, 46)?.fragment).toBe(acquired);
    expect(partitionFootnote(acquired, null, 27)).toBeNull();
  });

  it('retains a trailing empty paragraph when it continues alone', () => {
    const original = note(0);
    const empty = {
      ...original.story.blocks[0],
      id: 'empty', source: { ...source, path: [1] }, lines: [],
      flowBounds: { xPt: 0, yPt: 46, widthPt: 180, heightPt: 11 },
      inkBounds: { xPt: 0, yPt: 46, widthPt: 180, heightPt: 11 },
      advancePt: 11,
    } as ParagraphLayout;
    const story = {
      ...original.story,
      blocks: [original.story.blocks[0]!, empty],
      advancePt: 51,
      flowBounds: { ...original.story.flowBounds, heightPt: 51 },
    };
    const acquired: NoteLayout = {
      ...original, story, advancePt: 57,
      flowBounds: { ...original.flowBounds, heightPt: 57 },
    };
    const first = partitionFootnote(acquired, null, 46)!;
    expect(first.nextCursor).not.toBeNull();
    const second = partitionFootnote(acquired, first.nextCursor, 100)!;
    expect(second.nextCursor).toBeNull();
    expect(second.fragment.story.blocks.map((block) => block.source.path)).toEqual([[1]]);
  });

  it('reserves a measured continuation notice only on a continuing fragment', () => {
    let acquisitions = 0;
    const provider = () => { acquisitions += 1; return notice(8); };
    const intact = partitionFootnote(note(0), null, 46, provider)!;
    expect(intact.nextCursor).toBeNull();
    expect(intact.fragment.trailing).toBeUndefined();
    expect(acquisitions).toBe(0);
    // Three lines fit without the notice (36pt); its 8pt reserve leaves two.
    const first = partitionFootnote(note(0), null, 37, provider)!;
    expect(first.nextCursor?.lineIndex).toBe(2);
    expect(first.fragment.advancePt).toBe(34);
    expect(first.fragment.trailing).toMatchObject({
      role: 'continuationNotice', advancePt: 8, flowBounds: { yPt: 26, heightPt: 8 },
      paragraph: { flowDomainId: 'notes:page:0', flowBounds: { yPt: 26 } },
    });
    // The page occurrence is projected under the continuing note's identity.
    expect(first.fragment.trailing?.paragraph?.id.startsWith('footnote:7:page:0:notice')).toBe(true);
    const last = partitionFootnote(note(1), first.nextCursor, 100, provider)!;
    expect(last.nextCursor).toBeNull();
    expect(last.fragment.trailing).toBeUndefined();
    // A notice that leaves no real note line relocates the reference rather
    // than retaining a separator/notice-only fragment.
    expect(partitionFootnote(note(0), null, 23, provider)).toBeNull();
    // An empty or hidden notice story owns no advance.
    expect(partitionFootnote(note(0), null, 37, () => null)!.fragment.advancePt).toBe(36);
  });
});
