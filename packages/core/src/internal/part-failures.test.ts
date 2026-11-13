import { describe, expect, it, vi } from 'vitest';
import {
  clampPartFailureMessage,
  errorPlaceholderCanvasLines,
  errorPlaceholderDetail,
  errorPlaceholderTitle,
  normalizeErrorPlaceholderText,
  PART_FAILURE_MESSAGES_MAX_UTF8,
  PartFailureDelivery,
  PartFailureLog,
  prepareErrorPlaceholder,
  resolveErrorPlaceholderFor,
  resolvedFromWire,
  wireErrorPlaceholder,
  wrapErrorPlaceholderText,
} from './part-failures';

interface Failure { readonly kind: 'slide'; readonly slideIndex: number; readonly message: string }
const failure = (slideIndex: number, message = `ppt/slides/slide${slideIndex + 1}.xml: bad`): Failure =>
  ({ kind: 'slide', slideIndex, message });
const newLog = () => new PartFailureLog<Failure>((f) => f.slideIndex, (a, b) => a.slideIndex - b.slideIndex);

describe('normalizeErrorPlaceholderText', () => {
  it('keeps the built-in text for undefined and {}', () => {
    expect(normalizeErrorPlaceholderText(undefined)).toBeNull();
    expect(normalizeErrorPlaceholderText({})).toBeNull();
  });

  it('hides the detail for null and empty detail', () => {
    expect(normalizeErrorPlaceholderText({ detail: null })).toEqual({ title: null, detail: null });
    expect(normalizeErrorPlaceholderText({ detail: '' })).toEqual({ title: null, detail: null });
  });

  it('rejects invalid shapes', () => {
    expect(() => normalizeErrorPlaceholderText(null)).toThrow(TypeError);
    expect(() => normalizeErrorPlaceholderText('x')).toThrow(TypeError);
    expect(() => normalizeErrorPlaceholderText({ title: 1 })).toThrow(TypeError);
    expect(() => normalizeErrorPlaceholderText({ detail: 2 })).toThrow(TypeError);
  });

  it('reads each property exactly once', () => {
    const reads = { title: 0, detail: 0 };
    const value = {
      get title() { reads.title += 1; return 'T'; },
      get detail() { reads.detail += 1; return 'D'; },
    };
    expect(normalizeErrorPlaceholderText(value)).toEqual({ title: 'T', detail: 'D' });
    expect(reads).toEqual({ title: 1, detail: 1 });
  });

  it('paints the built-in text where a field is omitted', () => {
    const resolved = normalizeErrorPlaceholderText({ title: 'Custom' });
    expect(errorPlaceholderTitle(resolved, 'Built-in')).toBe('Custom');
    expect(errorPlaceholderDetail(resolved, 'detail')).toBe('detail');
    expect(errorPlaceholderTitle(null, 'Built-in')).toBe('Built-in');
  });
});

describe('prepared options', () => {
  it('passes the failure to a function and validates its result', () => {
    const prepared = prepareErrorPlaceholder<Failure>((f: Failure) => ({ title: `Slide ${f.slideIndex + 1} failed` }));
    expect(resolveErrorPlaceholderFor(prepared, failure(2))).toEqual({ title: 'Slide 3 failed', detail: undefined });
    const invalid = prepareErrorPlaceholder<Failure>(() => ({ title: 3 }) as never);
    expect(() => resolveErrorPlaceholderFor(invalid, failure(0))).toThrow(TypeError);
    const undefinedResult = prepareErrorPlaceholder<Failure>(() => undefined);
    expect(resolveErrorPlaceholderFor(undefinedResult, failure(0))).toBeNull();
    expect(wireErrorPlaceholder(null)).toEqual({ state: 'builtin' });
  });

  it('rejects an invalid static option up front', () => {
    expect(() => prepareErrorPlaceholder({ title: false })).toThrow(TypeError);
  });

  it('round-trips the wire form', () => {
    const resolved = normalizeErrorPlaceholderText({ title: 'T', detail: null });
    expect(resolvedFromWire(wireErrorPlaceholder(resolved))).toEqual({ title: 'T', detail: null });
    expect(resolvedFromWire(undefined)).toBeNull();
    expect(() => resolvedFromWire({ state: 'resolve-on-main' })).toThrow();
  });
});

describe('clampPartFailureMessage', () => {
  it('keeps short messages unchanged', () => {
    expect(clampPartFailureMessage('xl/worksheets/sheet1.xml: bad')).toBe('xl/worksheets/sheet1.xml: bad');
  });

  it('clamps to the byte budget at a grapheme boundary', () => {
    const clamped = clampPartFailureMessage('あ'.repeat(2000), 100);
    expect(new TextEncoder().encode(clamped).length).toBeLessThanOrEqual(100);
    expect(clamped.endsWith('…')).toBe(true);
    const emoji = clampPartFailureMessage('👨‍👩‍👧'.repeat(50), 40);
    expect(emoji).toBe(`${'👨‍👩‍👧'.repeat(2)}…`);
  });
});

describe('PartFailureLog', () => {
  it('dedupes by key and sorts the snapshot', () => {
    const log = newLog();
    log.add(failure(3));
    log.add(failure(1));
    expect(log.add(failure(3, 'other'))).toBeUndefined();
    expect(log.list.map((f) => f.slideIndex)).toEqual([1, 3]);
    expect(Object.isFrozen(log.list)).toBe(true);
    expect(Object.isFrozen(log.list[0])).toBe(true);
  });

  it('bounds retained message bytes and keeps identities on overflow', () => {
    const log = newLog();
    const long = 'x'.repeat(10_000);
    const count = Math.ceil(PART_FAILURE_MESSAGES_MAX_UTF8 / 1024) + 5;
    for (let index = 0; index < count; index += 1) log.add(failure(index, long));
    const list = log.list;
    expect(list).toHaveLength(count);
    const bytes = list.reduce((sum, f) => sum + new TextEncoder().encode(f.message).length, 0);
    expect(bytes).toBeLessThanOrEqual(PART_FAILURE_MESSAGES_MAX_UTF8);
    expect(list.at(-1)!.message).toBe('');
    expect(list[0].message.length).toBeGreaterThan(0);
  });
});

describe('PartFailureDelivery', () => {
  it('replays existing failures, then delivers live ones once', () => {
    const log = newLog();
    log.add(failure(0));
    const seen: number[] = [];
    const delivery = new PartFailureDelivery<Failure>('Viewer', (f) => seen.push(f.slideIndex), (f) => f.slideIndex);
    delivery.attach(log);
    log.add(failure(2));
    delivery.attach(log);
    expect(seen).toEqual([0, 2]);
  });

  it('drops the rest of a replay when a callback closes the viewer', () => {
    const log = newLog();
    log.add(failure(0));
    log.add(failure(1));
    const seen: number[] = [];
    const delivery: PartFailureDelivery<Failure> = new PartFailureDelivery<Failure>('Viewer', (f) => {
      seen.push(f.slideIndex);
      delivery.close();
    }, (f) => f.slideIndex);
    delivery.attach(log);
    expect(seen).toEqual([0]);
  });

  it('drops the old engine when a callback attaches a new one', () => {
    const first = newLog();
    first.add(failure(0));
    first.add(failure(1));
    const second = newLog();
    second.add(failure(5));
    const seen: number[] = [];
    const delivery: PartFailureDelivery<Failure> = new PartFailureDelivery<Failure>('Viewer', (f) => {
      seen.push(f.slideIndex);
      if (f.slideIndex === 0) delivery.attach(second);
    }, (f) => f.slideIndex);
    delivery.attach(first);
    first.add(failure(2));
    expect(seen).toEqual([0, 5]);
  });

  it('isolates callback errors', () => {
    const error = vi.spyOn(console, 'error').mockImplementation(() => {});
    const log = newLog();
    log.add(failure(0));
    log.add(failure(1));
    const seen: number[] = [];
    const delivery = new PartFailureDelivery<Failure>('Viewer', (f) => {
      seen.push(f.slideIndex);
      throw new Error('boom');
    }, (f) => f.slideIndex);
    expect(() => delivery.attach(log)).not.toThrow();
    expect(seen).toEqual([0, 1]);
    expect(error).toHaveBeenCalledTimes(2);
    error.mockRestore();
  });

  it('a reload started from a callback drops the rest of the replay', () => {
    const log = newLog();
    log.add(failure(0));
    log.add(failure(1));
    const seen: number[] = [];
    let epoch = 0;
    const delivery: PartFailureDelivery<Failure> = new PartFailureDelivery<Failure>('Viewer', (f) => {
      seen.push(f.slideIndex);
      if (f.slideIndex === 0) epoch = delivery.beginLoad();
    }, (f) => f.slideIndex);
    delivery.attach(log);
    expect(seen).toEqual([0]);
    log.add(failure(2));
    expect(seen).toEqual([0]);
    // The reload failed before commit: the retained engine resumes, without repeats.
    delivery.endLoad(epoch, log);
    expect(seen).toEqual([0, 1, 2]);
  });

  it('ignores a superseded load when it settles', () => {
    const log = newLog();
    log.add(failure(0));
    const seen: number[] = [];
    const delivery = new PartFailureDelivery<Failure>('Viewer', (f) => seen.push(f.slideIndex), (f) => f.slideIndex);
    const first = delivery.beginLoad();
    delivery.beginLoad();
    delivery.endLoad(first, log);
    expect(seen).toEqual([]);
  });

  it('isolates an async observer rejection', async () => {
    const error = vi.spyOn(console, 'error').mockImplementation(() => {});
    const log = newLog();
    log.add(failure(0));
    const delivery = new PartFailureDelivery<Failure>(
      'Viewer',
      (async () => { throw new Error('async boom'); }) as unknown as (f: Failure) => void,
      (f) => f.slideIndex,
    );
    delivery.attach(log);
    await vi.waitFor(() => expect(error).toHaveBeenCalledTimes(1));
    error.mockRestore();
  });
});


describe('placeholder text bounds', () => {
  const width = (value: string) => [...value].length * 10;

  it('caps caller and parser text before it is painted', () => {
    const huge = 'word '.repeat(200_000);
    const title = errorPlaceholderTitle({ title: huge, detail: undefined }, 'Built-in');
    const detail = errorPlaceholderDetail(null, huge)!;
    for (const text of [title, detail]) {
      expect(new TextEncoder().encode(text).length).toBeLessThanOrEqual(1024);
      expect(text.endsWith('…')).toBe(true);
    }
    expect(errorPlaceholderDetail({ title: null, detail: null }, huge)).toBeNull();
  });

  it('fits every line to the width, breaking long words and marking cut text', () => {
    const title = wrapErrorPlaceholderText('Sheet could not be displayed at all', 120, 1, width);
    expect(title).toHaveLength(1);
    expect(title[0].endsWith('…')).toBe(true);
    const detail = wrapErrorPlaceholderText(`${'x'.repeat(30)} tail`, 100, 4, width);
    expect(detail).toEqual(['x'.repeat(10), 'x'.repeat(10), 'x'.repeat(10), 'tail']);
    for (const line of [...title, ...wrapErrorPlaceholderText('y'.repeat(100), 100, 4, width)]) {
      expect(width(line)).toBeLessThanOrEqual(120);
    }
  });

  it('reuses the wrapped lines while the text, font and width are unchanged', () => {
    const measureText = vi.fn((value: string) => ({ width: width(value) }));
    const ctx = { font: '10px sans-serif', measureText };
    const first = errorPlaceholderCanvasLines(ctx, 'a b c d e f', 30, 4);
    const calls = measureText.mock.calls.length;
    expect(errorPlaceholderCanvasLines(ctx, 'a b c d e f', 30, 4)).toBe(first);
    expect(measureText).toHaveBeenCalledTimes(calls);
    ctx.font = '20px sans-serif';
    errorPlaceholderCanvasLines(ctx, 'a b c d e f', 30, 4);
    expect(measureText.mock.calls.length).toBeGreaterThan(calls);
  });
});
