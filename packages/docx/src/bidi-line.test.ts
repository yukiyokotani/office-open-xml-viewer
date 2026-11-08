import { describe, it, expect } from 'vitest';
import { getDefaultBidiEngine, setBidiEngine } from '@silurus/ooxml-core';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  createLineOrderResolver,
  resolveAlignEdge,
  jcIsFullyJustified,
  jcStretchesLastLine,
} from './bidi-line.js';

describe('segmentsHaveRtl', () => {
  it('detects Arabic/Hebrew, ignores Latin and objects', () => {
    expect(segmentsHaveRtl([{ text: 'Hello world' }])).toBe(false);
    expect(segmentsHaveRtl([{ text: 'price 99' }, {}])).toBe(false); // object seg has no text
    expect(segmentsHaveRtl([{ text: 'مرحبا' }])).toBe(true);
    expect(segmentsHaveRtl([{ text: 'abc ' }, { text: 'שלום' }])).toBe(true);
  });

  it('detects a run-level rtl mark even on neutral-only text', () => {
    // "1. " has no strong-RTL char, but the run carries <w:rtl> (§17.3.2.30).
    expect(segmentsHaveRtl([{ text: '1. ', rtl: true }])).toBe(true);
    expect(segmentsHaveRtl([{ text: '1. ' }])).toBe(false);
  });
});

describe('authored opposite-direction run ordering', () => {
  it('keeps a leading Latin label at the left of its unmarked Arabic run', () => {
    const segments = [{ text: 'Label ' }, { text: 'هذا ' }, { text: 'نص ' }, { text: 'عربي ' }];
    const resolve = createLineOrderResolver();
    const visual = resolve(segments, true, 'word');
    expect(visual.order.map(index => segments[index]!.text.trim()))
      .toEqual(['Label', 'عربي', 'نص', 'هذا']);
    // Acquisition callers and unauthored shape text retain their existing
    // line-level resolution; this change only orders authored final lines.
    expect(resolve(segments, true).order.at(-1)).toBe(0);
    expect(resolve(segments, true, 'word').order[0]).toBe(0);
    segments[0]!.text = 'אחר ';
    expect(resolve(segments, true, 'word').order).toEqual([3, 2, 1, 0]);
  });

  it('honors same-object adapter lifecycle and method changes during ordering reuse', () => {
    const engine = getDefaultBidiEngine();
    let reverseAdapterOrder = false;
    const adapter = {
      computeLevels: (...args: Parameters<typeof engine.computeLevels>) => engine.computeLevels(...args),
      reorderVisual: (...args: Parameters<typeof engine.reorderVisual>) => {
        const order = engine.reorderVisual(...args);
        return reverseAdapterOrder ? order.reverse() : order;
      },
      getMirror: (...args: Parameters<typeof engine.getMirror>) => engine.getMirror(...args),
    };
    const resolve = createLineOrderResolver();
    const words = [{ text: 'שלום ' }, { text: 'עולם' }];
    try {
      setBidiEngine(adapter);
      expect(resolve(words, true).order).toEqual([1, 0]);
      // A provider can change captured configuration while retaining its
      // object and methods; the explicit setter is its invalidation contract.
      reverseAdapterOrder = true;
      setBidiEngine(adapter);
      expect(resolve(words, true).order).toEqual([0, 1]);
      adapter.reorderVisual = (...args) => engine.reorderVisual(...args);
      expect(resolve(words, true).order).toEqual([1, 0]);
      adapter.computeLevels = (text) => ({ levels: new Uint8Array(text.length), paragraphLevel: 0 });
      expect(resolve(words, true).order).toEqual([0, 1]);
    } finally { setBidiEngine(engine); }
  });

  it('honors replaced native methods and proxy providers for uniform RTL slices', () => {
    const engine = getDefaultBidiEngine();
    const reorder = engine.reorderVisual;
    const words = [{ text: 'שלום ', rtl: true }, { text: 'עולם', rtl: true }];
    const resolve = createLineOrderResolver();
    try {
      expect(resolve(words, true).order).toEqual([1, 0]);
      engine.reorderVisual = (...args) => reorder.call(engine, ...args).reverse();
      expect(resolve(words, true).order).toEqual([0, 1]);
      engine.reorderVisual = reorder;
      const proxy = new Proxy(engine, {
        get(target, key, receiver) {
          return key === 'reorderVisual'
            ? (...args: Parameters<typeof reorder>) => reorder.call(target, ...args).reverse()
            : Reflect.get(target, key, receiver);
        },
      });
      setBidiEngine(proxy);
      expect(resolve(words, true).order).toEqual([0, 1]);
    } finally {
      engine.reorderVisual = reorder;
      setBidiEngine(engine);
    }
  });

  it('preserves unresolved space and terminal-neutral anchor contexts', () => {
    const resolve = createLineOrderResolver();
    // L1 resets the tab, not the space AFTER it: the earlier L context can
    // still give that space a level-2 anchor.
    expect(resolve([{ text: 'AB', rtl: true }, { isTab: true },
      { text: ' ', rtl: true }, { text: 'CD', rtl: true }], true, 'word'))
      .toEqual({ order: [2, 3, 1, 0], rtl: [false, true, false, false] });
    // An opposite unmarked space shares the LRI span with the following L.
    expect(resolve([{ text: 'א', rtl: true }, { text: ' ' },
      { text: 'AB' }, { text: 'ב', rtl: true }], true, 'word'))
      .toEqual({ order: [3, 1, 2, 0], rtl: [true, false, false, true] });
    // A terminal neutral is not bounded by L outside an isolate.
    expect(resolve([{ text: 'AB!' }, { text: 'א', rtl: true }], true, 'line-level'))
      .toEqual({ order: [1, 0], rtl: [true, true] });
  });

  it('leaves paragraph-direction Arabic ordering and paint direction unchanged', () => {
    const segments = [{ text: 'هذا ', rtl: true }, { text: 'نص ', rtl: true }, { text: 'عربي ', rtl: true }];
    expect(computeLineVisualOrder(segments, true, 'word'))
      .toEqual(computeLineVisualOrder(segments, true));
  });
});

  it('keeps LTR word order for English text in rtl-marked runs', () => {
    // w:rtl on a run of pure Latin (e.g. "first leader name" in sample-7's
    // signature block): the RLE embedding must NOT reverse the words — only
    // weak/neutral content (digits, punctuation) takes the RTL embedding.
    const segs = [
      { text: 'first ', rtl: true },
      { text: 'leader ', rtl: true },
      { text: 'name', rtl: true },
    ];
    const { order, rtl } = computeLineVisualOrder(segs, false);
    expect(order).toEqual([0, 1, 2]);
    expect(rtl).toEqual([false, false, false]);
  });

describe('resolveAlignEdge', () => {
  it('resolves logical start/end against base direction', () => {
    expect(resolveAlignEdge(undefined, false)).toBe('left');
    expect(resolveAlignEdge(undefined, true)).toBe('right');
    expect(resolveAlignEdge('start', true)).toBe('right');
    expect(resolveAlignEdge('end', true)).toBe('left');
    expect(resolveAlignEdge('end', false)).toBe('right');
    expect(resolveAlignEdge('center', true)).toBe('center');
    expect(resolveAlignEdge('both', true)).toBe('justify');
    // Transitional left/right are "semantically equivalent to start/end"
    // (ECMA-376 Part 4 §14.11.2) — logical, so they flip under an RTL base.
    expect(resolveAlignEdge('left', true)).toBe('right');
    expect(resolveAlignEdge('right', true)).toBe('left');
    expect(resolveAlignEdge('left', false)).toBe('left');
    expect(resolveAlignEdge('right', false)).toBe('right');
  });

  // ECMA-376 §17.18.44 — the Arabic kashida variants and thaiDistribute are all
  // forms of full justification; their physical edge is "justify", not the
  // leading-edge default they previously fell through to.
  it('maps kashida + thaiDistribute jc values to the justify edge', () => {
    expect(resolveAlignEdge('lowKashida', false)).toBe('justify');
    expect(resolveAlignEdge('mediumKashida', false)).toBe('justify');
    expect(resolveAlignEdge('highKashida', false)).toBe('justify');
    expect(resolveAlignEdge('thaiDistribute', false)).toBe('justify');
    // Independent of base direction.
    expect(resolveAlignEdge('lowKashida', true)).toBe('justify');
    expect(resolveAlignEdge('thaiDistribute', true)).toBe('justify');
  });
});

describe('jcIsFullyJustified (§17.18.44)', () => {
  it('is true for every full-justification jc value incl. kashida/thai', () => {
    for (const jc of ['both', 'justify', 'distribute', 'lowKashida', 'mediumKashida', 'highKashida', 'thaiDistribute']) {
      expect(jcIsFullyJustified(jc), jc).toBe(true);
    }
  });
  it('is false for non-justifying values', () => {
    for (const jc of ['left', 'right', 'center', 'start', 'end', undefined, 'numTab']) {
      expect(jcIsFullyJustified(jc), String(jc)).toBe(false);
    }
  });
});

describe('jcStretchesLastLine (§17.18.44)', () => {
  it('stretches the last line only for distribute', () => {
    expect(jcStretchesLastLine('distribute')).toBe(true);
    // thaiDistribute leaves its final line ragged like `both` (Word GT, #959):
    // only non-final lines get the Thai cluster distribution.
    expect(jcStretchesLastLine('thaiDistribute')).toBe(false);
    // both/justify/kashida leave the last line as-is.
    for (const jc of ['both', 'justify', 'lowKashida', 'mediumKashida', 'highKashida', 'left', undefined]) {
      expect(jcStretchesLastLine(jc), String(jc)).toBe(false);
    }
  });
});

describe('computeLineVisualOrder', () => {
  it('keeps pure-LTR segments in logical order', () => {
    const { order, rtl } = computeLineVisualOrder([{ text: 'Hello ' }, { text: 'world' }], false);
    expect(order).toEqual([0, 1]);
    expect(rtl).toEqual([false, false]);
  });

  it('reverses pure-RTL segments and marks them RTL', () => {
    const { order, rtl } = computeLineVisualOrder([{ text: 'שלום ' }, { text: 'עולם' }], true);
    expect(order).toEqual([1, 0]); // last word is visually leftmost
    expect(rtl).toEqual([true, true]);
  });

  it('keeps leading neutral punctuation with its RTL word under an LTR paragraph base', () => {
    const { order, rtl } = computeLineVisualOrder(
      [{ text: ':مرحبا ' }, { text: 'في ' }, { text: 'العالم' }],
      false,
    );
    // The colon resolves to the LTR paragraph level, but Canvas paints the
    // whole first slice RTL. Its ordering level must therefore come from the
    // same odd-level content rather than stranding the slice at the left edge.
    expect(order).toEqual([2, 1, 0]);
    expect(rtl).toEqual([true, true, true]);
  });

  it('orders an embedded LTR run inside an RTL line', () => {
    // logical: Hebrew, Latin, Hebrew (base RTL) -> visual L→R: [2,1,0]
    const { order, rtl } = computeLineVisualOrder(
      [{ text: 'שלום ' }, { text: 'world ' }, { text: 'טוב' }],
      true,
    );
    expect(order).toEqual([2, 1, 0]);
    expect(rtl[1]).toBe(false); // the Latin segment is LTR
    expect(rtl[0]).toBe(true);
    expect(rtl[2]).toBe(true);
  });

  it('resolves a leading-digit run-level rtl run with digits at the trailing (right) end', () => {
    // "1. " + Hebrew, both run-level rtl, base LTR (no w:bidi). The renderer
    // marks rtl-run digits `digitsAsAN` (w:lang w:bidi RTL or the rtl-marked
    // fallback, §17.3.2.20), classifying them AN; the rtl mark resolves the
    // "." to R (§17.3.2.30). Visually the Hebrew is leftmost and the "1. "
    // (read RTL → ".1") sits at the right end.
    const { order, rtl } = computeLineVisualOrder(
      [{ text: '1. ', rtl: true, digitsAsAN: true }, { text: 'תוכן', rtl: true }],
      false,
    );
    expect(order).toEqual([1, 0]); // Hebrew visually left, "1." visually right
    expect(rtl).toEqual([true, true]);
  });

  it('keeps a trailing non-rtl numeric run leftmost in a base-RTL title (§17.3.2.30)', () => {
    // sample-7 title "المشاريع لعام 2026": three runs, the Arabic ones carry
    // w:rtl, "2026" does not, in a base-RTL (w:bidi) paragraph. Word renders
    // "2026" at the LEFT (logically last in the RTL line). A w:rtl run that
    // already has strong-RTL content must NOT be over-embedded above the base.
    const { order, rtl } = computeLineVisualOrder(
      [
        { text: 'المشاريع ', rtl: true },
        { text: 'لعام ', rtl: true },
        { text: '2026', rtl: false },
      ],
      true,
    );
    // `order` is logical indices in visual LEFT→RIGHT order, so the LEFTMOST
    // segment is order[0] (cf. the pure-RTL case where order=[1,0] puts the
    // logically-last word leftmost).
    expect(order).toEqual([2, 1, 0]);
    expect(order[0]).toBe(2); // "2026" visually leftmost
    expect(rtl[2]).toBe(false);
  });

  it('keeps the LTR fast path byte-identical when no rtl marks and no strong-RTL', () => {
    const { order, rtl } = computeLineVisualOrder(
      [{ text: '1. ' }, { text: 'item' }],
      false,
    );
    expect(order).toEqual([0, 1]);
    expect(rtl).toEqual([false, false]);
  });

  it('places a neutral object by surrounding direction', () => {
    // object between two Hebrew words, base RTL -> object stays in RTL flow
    const { order } = computeLineVisualOrder([{ text: 'שלום ' }, {}, { text: 'טוב' }], true);
    expect(order).toEqual([2, 1, 0]);
  });

  it('orders an AN-classified date as 2026-02-28', () => {
    // A logical "28-02-2026" in an Arabic complex-script run is pre-split into
    // digit-group and separator segments tagged `digitsAsAN`. Under the RTL
    // base, UAX #9 §4.3 HL1 reorders those groups to "2026-02-28".
    const segs = [
      { text: '28', rtl: false, digitsAsAN: false },
      { text: '-', rtl: false, digitsAsAN: false },
      { text: '02', rtl: false, digitsAsAN: false },
      { text: '-', rtl: false, digitsAsAN: false },
      { text: '2026', rtl: false, digitsAsAN: false },
    ];
    const resolve = createLineOrderResolver();
    expect(resolve(segs, true).order.map(i => segs[i]!.text).join('')).toBe('28-02-2026');
    for (const segment of segs) segment.digitsAsAN = true;
    const { order } = resolve(segs, true);
    // Visual L→R: 2026 - 02 - 28  →  logical indices [4,3,2,1,0].
    const visual = order.map((i) => segs[i].text).join('');
    expect(visual).toBe('2026-02-28');
    for (const segment of segs) segment.rtl = true;
    expect(resolve(segs, true).order.map(i => segs[i]!.text).join('')).toBe('2026-02-28');
  });

  it('reorders date groups even without AN classification when the run is rtl-marked', () => {
    // Same split, base RTL, no digitsAsAN: the digits stay EN, but the
    // rtl-marked "-" separators are ambiguous punctuation and resolve R per
    // §17.3.2.30, so the EN groups still reorder to Word's visual order.
    const segs = [
      { text: '28', rtl: false },
      { text: '-', rtl: false },
      { text: '02', rtl: false },
      { text: '-', rtl: false },
      { text: '2026', rtl: false },
    ];
    const resolve = createLineOrderResolver();
    expect(resolve(segs, true).order.map(i => segs[i]!.text).join('')).toBe('28-02-2026');
    for (const segment of segs) segment.rtl = true;
    const { order } = resolve(segs, true);
    const visual = order.map((i) => segs[i].text).join('');
    expect(visual).toBe('2026-02-28');
  });
});

describe('rtl-marked run punctuation (§17.3.2.30 ambiguous → RTL)', () => {
  it('places a trailing rtl-marked "." at the line start edge, not after the number', () => {
    // sample-7 page 1: logical [Arabic… (rtl)]["2022"][". " (rtl)] under w:bidi.
    // Word renders the period at the visual LEFT end (".2022 …"); the old
    // RLE…PDF wrap over-embedded the "." to level base+2 and produced "2022.".
    const segs = [
      { text: 'يناير ', rtl: true },
      { text: '2022' },
      { text: '.', rtl: true },
    ];
    const { order, rtl } = computeLineVisualOrder(segs, true);
    // Visual left→right: [.][2022][Arabic]
    expect(order).toEqual([2, 1, 0]);
    expect(rtl[2]).toBe(true); // the "." draws in RTL context
  });

  it('still mirrors a literal "1. " prefix to ".1" (digit group + period)', () => {
    // digitsAsAN pre-split of an rtl-marked "1. " run in a base-LTR paragraph
    // (sample-7's zero-height numbered headings).
    const segs = [
      { text: '1', rtl: true, digitsAsAN: true },
      { text: '. ', rtl: true, digitsAsAN: true },
    ];
    const { order } = computeLineVisualOrder(segs, false);
    // Visual left→right: [". "]["1"] → rendered ".1"
    expect(order).toEqual([1, 0]);
  });

  it('does not reorder English words in rtl-marked runs (spaces untouched)', () => {
    const segs = [
      { text: 'first ', rtl: true },
      { text: 'leader ', rtl: true },
      { text: 'name', rtl: true },
    ];
    const { order, rtl } = computeLineVisualOrder(segs, false);
    expect(order).toEqual([0, 1, 2]);
    expect(rtl).toEqual([false, false, false]);
  });

  it('keeps a non-rtl trailing number at the left end of an RTL title', () => {
    // sample-7 page-1 title: rtl Arabic runs + plain "2026" run, w:bidi center.
    const segs = [
      { text: 'المشاريع ', rtl: true },
      { text: 'لعام', rtl: true },
      { text: ' ', rtl: true },
      { text: '2026' },
    ];
    const { order } = computeLineVisualOrder(segs, true);
    // "2026" must be visually leftmost (first in visual order).
    expect(order[0]).toBe(3);
  });

  it('keeps the date group order under digitsAsAN', () => {
    const segs = [
      { text: '28', rtl: true, digitsAsAN: true },
      { text: '-', rtl: true, digitsAsAN: true },
      { text: '02', rtl: true, digitsAsAN: true },
      { text: '-', rtl: true, digitsAsAN: true },
      { text: '2026', rtl: true, digitsAsAN: true },
    ];
    const { order } = computeLineVisualOrder(segs, true);
    // Visual left→right: 2026 - 02 - 28
    expect(order).toEqual([4, 3, 2, 1, 0]);
  });
});


describe('authored numeric common separator', () => {
  it('leaves numeric type matching to W2 before W4 after an AL strong source', () => {
    // Explicitly no digitsAsAN: the existing UAX W2 owns EN→AN here. The
    // interior CS is numeric membership, not a host-side W2 reimplementation.
    const segments = [{ text: 'ا', rtl: true }, { text: '1', rtl: true },
      { text: '.', rtl: true }, { text: '٢', rtl: true }];
    const visual = computeLineVisualOrder(segments, true);
    expect(visual.order.map(i => segments[i]!.text).join('')).toBe('1.٢ا');
    expect(visual.rtl[2]).toBe(false);
  });

  it('uses original scalar numeric neighbors, including astral EN', () => {
    const segments = [{ text: '\u{1D7CE}', rtl: true },
      { text: '.', rtl: true }, { text: '\u{1D7CF}', rtl: true }];
    const visual = computeLineVisualOrder(segments, true);
    expect(visual.order.map(i => segments[i]!.text).join('')).toBe('\u{1D7CE}.\u{1D7CF}');
    expect(visual.rtl).toEqual([false, false, false]);
  });

  it('does not bridge tabs, objects, controls, author isolates or differing digit classes', () => {
    const numeric = { text: '12', rtl: true };
    const period = { text: '.', rtl: true };
    const last = { text: '5', rtl: true };
    // Each boundary prevents immediate matching numeric adjacency for a
    // distinct reason; none is skipped when constructing the complete host.
    for (const barrier of [{ isTab: true }, {}, { text: '\u200f' }]) {
      expect(computeLineVisualOrder([numeric, barrier, period, last], true).rtl[2]).toBe(true);
    }
    const authored = [{ text: '12' }, period, last];
    expect(computeLineVisualOrder(authored, true, 'word').rtl[1]).toBe(true);
    expect(computeLineVisualOrder([{ ...numeric, digitsAsAN: true }, period, last], true).rtl[1]).toBe(true);
  });

  it.each([false, true])('keeps a formatted numeric CS bridge in numeric source order, AN=%s', digitsAsAN => {
    // A formatting seam does not end a Unicode number. Same effective numeric
    // class on each side permits W4; retain EN versus authored AN precedence.
    const segments = [{ text: '12', rtl: true, digitsAsAN },
      { text: '.', rtl: true, digitsAsAN }, { text: '5', rtl: true, digitsAsAN }];
    const visual = computeLineVisualOrder(segments, true);
    expect(visual.order.map(i => segments[i]!.text).join('')).toBe('12.5');
    expect(visual.rtl).toEqual([false, false, false]);
  });
});
