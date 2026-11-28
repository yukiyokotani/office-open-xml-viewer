import { describe, expect, it } from 'vitest';
import {
  TEXT_ADVANCE_REFERENCE_PX,
  createReferenceAdvanceContext,
  createTextAdvancePlanner,
  fillTextAtReference,
  referenceTextFont,
  withReferenceAdvances,
  type ReferenceAdvanceTarget,
} from './reference-advances';

const PROBES = ['a', 'e', 'm', 'W', '0', 'あ', '一'];
const FULL_WIDTH = /[　-鿿]/u;

/** A measuring double. `advance(character, px)` is the per-character width. */
function measuringContext(advance: (character: string, px: number) => number) {
  const measured: string[] = [];
  const target: ReferenceAdvanceTarget = {
    font: '10px serif',
    letterSpacing: '0px',
    fontKerning: 'auto',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    measureText(text: string) {
      measured.push(`${target.font}|${text}`);
      const px = Number(/(\d+(?:\.\d+)?)px/.exec(target.font)?.[1] ?? 10);
      const spacing = Number(/^(-?[\d.]+)px$/.exec(target.letterSpacing)?.[1] ?? 0);
      const width = [...text].reduce((sum, character) => sum + advance(character, px) + spacing, 0);
      return { width, fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2 } as TextMetrics;
    },
  };
  return { target, measured };
}

/** Grid fitting below 64px: half-width cells round to whole pixels. */
const gridFitted = (character: string, px: number) => {
  const half = px < 64 ? Math.floor(px / 2 + 0.5) : px / 2;
  return FULL_WIDTH.test(character) ? 2 * half : half;
};
const linear = (character: string, px: number) => (FULL_WIDTH.test(character) ? px : px / 2);

describe('reference text advances', () => {
  it('rewrites only the absolute px size of a font shorthand', () => {
    expect(referenceTextFont('normal 400 11px "MS Mincho", serif')).toEqual({
      font: `normal 400 ${TEXT_ADVANCE_REFERENCE_PX}px "MS Mincho", serif`,
      scale: 11 / TEXT_ADVANCE_REFERENCE_PX,
    });
    expect(referenceTextFont('italic bold 10.5px/1.2 Foo')?.font)
      .toBe(`italic bold ${TEXT_ADVANCE_REFERENCE_PX}px/1.2 Foo`);
    expect(referenceTextFont('1em serif')).toBeNull();
  });

  it('plans the reference em exactly for sizes the host grid-fits', () => {
    const { target, measured } = measuringContext(gridFitted);
    const planner = createTextAdvancePlanner(target, { probeClusters: PROBES });
    expect(planner.plan('normal 400 11px X')).toEqual(referenceTextFont('normal 400 11px X'));
    // Every probe cluster is an integer at an even size: grid fitting is a no-op.
    expect(planner.plan('normal 400 12px X')).toBeNull();
    // Spelling variants of one face and size share one decision and one probe.
    const probes = measured.length;
    expect(planner.plan('bold 11px X')).toEqual(referenceTextFont('bold 11px X'));
    expect(planner.plan('normal 700 11px X')).toEqual(referenceTextFont('normal 700 11px X'));
    // Reference-em clusters, the reference and doubled probe strings, and
    // every requested-size cluster (each departure must be a whole pixel).
    expect(measured.length - probes).toBe(PROBES.length * 2 + 2);
    expect(planner.planned('700 11px X')).toEqual(referenceTextFont('700 11px X'));
    expect(planner.planned('normal 400 13px X')).toBeUndefined();
    // The probe restores the measuring state it borrowed.
    expect(target.font).toBe('10px serif');
    expect(target.letterSpacing).toBe('0px');
    // Only the given probe scalars are measured.
    const scalars = new Set(measured.flatMap((entry) => [...entry.slice(entry.indexOf('|') + 1)]));
    expect([...scalars].every((scalar) => PROBES.includes(scalar))).toBe(true);
  });

  it('leaves proportional hosts and size-independent doubles unplanned', () => {
    const proportional = createTextAdvancePlanner(measuringContext(linear).target, { probeClusters: PROBES });
    expect(proportional.plan('normal 400 11px X')).toBeNull();
    const fixed = createTextAdvancePlanner(measuringContext(() => 6).target, { probeClusters: PROBES });
    expect(fixed.plan('normal 400 11px X')).toBeNull();
  });

  it('keeps optical-size designs at their own size', () => {
    // An opsz axis that tops out below the reference em: wider, fractional
    // advances at small sizes, exactly proportional from 96px up (the shape
    // of macOS system-ui: 'a' is 6.13px at 11px against 5.52 from 256px).
    const optical = (character: string, px: number) =>
      linear(character, px) * (px >= 96 ? 1 : 1 + 0.11 * (96 - px) / 85);
    expect(createTextAdvancePlanner(measuringContext(optical).target, { probeClusters: PROBES })
      .plan('normal 400 11px system-ui')).toBeNull();
    // A route mixing a whole-pixel departure with a fractional one has a
    // size-dependent design in it as well: no proof of grid fitting.
    const mixed = (character: string, px: number) => (character === '0'
      ? (px < 64 ? Math.round(px * 1.3) : px)
      : optical(character, px));
    expect(createTextAdvancePlanner(measuringContext(mixed).target, { probeClusters: PROBES })
      .plan('normal 400 11px X')).toBeNull();
    // A colour bitmap face that only supplies the digit still proves it.
    const strikeDigit = (character: string, px: number) => (character === '0'
      ? (px < 64 ? Math.round(px * 1.3) : px)
      : linear(character, px));
    expect(createTextAdvancePlanner(measuringContext(strikeDigit).target, { probeClusters: PROBES })
      .plan('normal 400 11px X')).toEqual(referenceTextFont('normal 400 11px X'));
  });

  it('measures a planned shorthand at the reference em with requested-px results', () => {
    const { target, measured } = measuringContext(gridFitted);
    const planner = createTextAdvancePlanner(target, { probeClusters: PROBES });
    const context = createReferenceAdvanceContext(target, planner);
    context.letterSpacing = '1px';
    context.font = 'normal 400 11px X';
    expect(context.font).toBe('normal 400 11px X');
    expect(context.letterSpacing).toBe('1px');
    expect(target.font).toBe(`normal 400 ${TEXT_ADVANCE_REFERENCE_PX}px X`);
    const before = measured.length;
    // Linear 11 + 11 plus one px of tracking after each character.
    expect(context.measureText('一一').width).toBeCloseTo(24, 12);
    expect(measured.length - before).toBe(1);
    context.font = 'normal 400 12px X';
    expect(target.font).toBe('normal 400 12px X');
    expect(target.letterSpacing).toBe('1px');
    expect(context.measureText('一').width).toBe(13);
  });

  it('serves the view through the concrete context type', () => {
    const { target } = measuringContext(gridFitted);
    const canvas = { id: 'canvas' };
    const owned = Object.assign(target, { canvas });
    const planner = createTextAdvancePlanner(owned, { probeClusters: PROBES });
    const view = createReferenceAdvanceContext(owned, planner);
    const typed = withReferenceAdvances(owned, view);
    typed.font = 'normal 400 11px X';
    typed.textAlign = 'center';
    expect(typed.font).toBe('normal 400 11px X');
    expect(owned.font).toBe(`normal 400 ${TEXT_ADVANCE_REFERENCE_PX}px X`);
    expect(owned.textAlign).toBe('center');
    expect(typed.canvas).toBe(canvas);
    expect(typed.measureText('\u4E00').width).toBe(11);
  });

  it('draws at the reference em under the linear scale and restores the state', () => {
    const calls: string[] = [];
    const ctx = {
      font: 'normal 400 11px X',
      letterSpacing: '2px',
      save() { calls.push('save'); },
      restore() { calls.push('restore'); },
      translate(x: number, y: number) { calls.push(`translate ${x} ${y}`); },
      scale(x: number, y: number) { calls.push(`scale ${x} ${y}`); },
      fillText(text: string, x: number, y: number) {
        calls.push(`fill ${text} ${x} ${y} ${ctx.font} ${ctx.letterSpacing}`);
      },
    };
    const reference = referenceTextFont('normal 400 11px X')!;
    fillTextAtReference(ctx, reference, 'abc', 3, 4);
    expect(calls).toEqual([
      'save',
      'translate 3 4',
      `scale ${reference.scale} ${reference.scale}`,
      `fill abc 0 0 normal 400 ${TEXT_ADVANCE_REFERENCE_PX}px X ${2 / reference.scale}px`,
      'restore',
    ]);
  });

  it('keeps a selected language across its save() and leaves untagged text a plain save()', () => {
    const paint = (lang: string | undefined) => {
      const accesses: string[] = [];
      const stack: string[] = [];
      const state = { lang: lang ?? 'inherit' };
      const fills: string[] = [];
      const target = {
        font: 'normal 400 11px X',
        letterSpacing: '0px',
        // Chrome resets the active state's language on save().
        save() { stack.push(state.lang); state.lang = 'inherit'; },
        restore() { state.lang = stack.pop() ?? state.lang; },
        translate() {},
        scale() {},
        fillText() { fills.push(state.lang); },
        get lang() { accesses.push('get'); return state.lang; },
        set lang(value: string) { accesses.push('set'); state.lang = value; },
      };
      const ctx = new Proxy(target, {
        has(object, key) {
          if (key === 'lang') accesses.push('has');
          return Reflect.has(object, key);
        },
      });
      fillTextAtReference(ctx, referenceTextFont('normal 400 11px X')!, 'a', 0, 0, lang);
      return { accesses, fills, lang: state.lang };
    };
    expect(paint(undefined)).toEqual({ accesses: [], fills: ['inherit'], lang: 'inherit' });
    const tagged = paint('sr-RS');
    expect(tagged.fills).toEqual(['sr-RS']);
    expect(tagged.lang).toBe('sr-RS');
  });
});
