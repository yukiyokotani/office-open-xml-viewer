import { describe, expect, it } from 'vitest';
import { pathFillModeOverlay, renderPresetShape } from './index';

// ECMA-376 §20.1.10.37 ST_PathFillMode amounts, from PowerPoint's PDF of
// custom-geometry paths in each mode (colours read from the content stream).
describe('pathFillModeOverlay', () => {
  const OFFICE: Record<string, Record<string, string>> = {
    '4472C4': { lighten: '8FAADC', lightenLess: '698ED0', darken: '294476', darkenLess: '375C9E' },
    ED7D31: { lighten: 'F4B183', lightenLess: 'F19659', darken: '8E4B1D', darkenLess: 'BF6427' },
    '808080': { lighten: 'B3B3B3', lightenLess: '999999', darken: '4D4D4D', darkenLess: '676767' },
    '000000': { lighten: '666666', lightenLess: '323232', darken: '000000', darkenLess: '000000' },
    FFFFFF: { lighten: 'FFFFFF', lightenLess: 'FFFFFF', darken: '999999', darkenLess: 'CDCDCD' },
    '70AD47': { lighten: 'A9CE91', lightenLess: '8CBD6B', darken: '43682B', darkenLess: '5A8B39' },
  };
  it('composites to the colours PowerPoint paints', () => {
    for (const [base, modes] of Object.entries(OFFICE)) {
      for (const [mode, expected] of Object.entries(modes)) {
        const [, r, g, b, a] = /rgba\((\d+),(\d+),(\d+),([\d.]+)\)/.exec(pathFillModeOverlay(mode)!)!.map(Number);
        for (let channel = 0; channel < 3; channel++) {
          const v = parseInt(base.slice(channel * 2, channel * 2 + 2), 16);
          const over = [r, g, b][channel];
          const out = Math.round(v * (1 - a) + over * a);
          const office = parseInt(expected.slice(channel * 2, channel * 2 + 2), 16);
          expect(Math.abs(out - office), `${base} ${mode}`).toBeLessThanOrEqual(1);
        }
      }
    }
    expect(pathFillModeOverlay('norm')).toBeNull();
    expect(pathFillModeOverlay('none')).toBeNull();
  });
});

function makeContext() {
  const events: string[] = [];
  const states: Array<{ fillStyle: unknown }> = [];
  let fillStyle: unknown = 'initial';
  const target = {
    beginPath: () => events.push('beginPath'),
    closePath: () => events.push('closePath'),
    moveTo: () => events.push('path'),
    lineTo: () => events.push('path'),
    bezierCurveTo: () => events.push('path'),
    quadraticCurveTo: () => events.push('path'),
    ellipse: () => events.push('path'),
    save: () => { states.push({ fillStyle }); events.push('save'); },
    restore: () => { fillStyle = states.pop()!.fillStyle; events.push('restore'); },
    fill: () => events.push(`fill:${String(fillStyle)}`),
    stroke: () => events.push('stroke'),
    get fillStyle() { return fillStyle; },
    set fillStyle(value: unknown) { fillStyle = value; events.push(`style:${String(value)}`); },
  };
  return { ctx: target as unknown as CanvasRenderingContext2D, events, get fillStyle() { return fillStyle; } };
}

describe('renderPresetShape paintFill', () => {
  it('paints the traced path before its tint overlay and stroke', () => {
    const a = makeContext();
    renderPresetShape(a.ctx, 'bordercallout2', 0, 0, 200, 100, [], '#base', () => {
      a.events.push('stroke');
    }, () => a.events.push('clearShadow'), {
      paintFill: () => { a.events.push('paintFill'); return true; },
    });

    expect(a.events.indexOf('beginPath')).toBeLessThan(a.events.indexOf('paintFill'));
    expect(a.events.indexOf('paintFill')).toBeLessThan(a.events.indexOf('clearShadow'));
    expect(a.events.indexOf('clearShadow')).toBeLessThan(a.events.indexOf('stroke'));
    expect(a.events).not.toContain('fill:#base');
  });

  it('does not invoke the painter for fill=none paths', () => {
    const a = makeContext();
    let paints = 0;
    renderPresetShape(a.ctx, 'straightconnector1', 0, 0, 200, 100, [], '#base', () => {
      a.events.push('stroke');
    }, () => a.events.push('clearShadow'), {
      paintFill: () => { paints++; return true; },
    });

    expect(paints).toBe(0);
    expect(a.events.filter((event) => event === 'stroke')).toHaveLength(1);
    expect(a.events).not.toContain('clearShadow');
  });

  it('reuses each multipath geometry for paint and tint without duplicating strokes', () => {
    const a = makeContext();
    let paints = 0;
    renderPresetShape(a.ctx, 'can', 0, 0, 200, 100, [], '#base', () => {
      a.events.push('stroke');
    }, () => a.events.push('clearShadow'), {
      paintFill: () => { paints++; a.events.push('paintFill'); return true; },
    });

    expect(paints).toBe(2);
    expect(a.events.filter((event) => event === 'fill:rgba(255,255,255,0.4)')).toHaveLength(1);
    expect(a.events.filter((event) => event === 'stroke')).toHaveLength(1);
    expect(a.events.filter((event) => event === 'clearShadow')).toHaveLength(1);
  });

  it('does not fill, tint, or clear shadow when the painter returns false', () => {
    const a = makeContext();
    renderPresetShape(a.ctx, 'can', 0, 0, 200, 100, [], '#base', null,
      () => a.events.push('clearShadow'), { paintFill: () => false });

    expect(a.events.some((event) => event.startsWith('fill:'))).toBe(false);
    expect(a.events).not.toContain('clearShadow');
  });

  it('combines the painter with skipTrailingStroke without losing the body stroke', () => {
    const a = makeContext();
    let paints = 0;
    renderPresetShape(a.ctx, 'bordercallout2', 0, 0, 200, 100, [], null, () => {
      a.events.push('stroke');
    }, () => {}, {
      paintFill: () => { paints++; return true; },
      skipTrailingStroke: true,
    });
    expect(paints).toBe(1);
    expect(a.events.filter((event) => event === 'stroke')).toHaveLength(1);
  });

  it('restores canvas state when the painter throws', () => {
    const a = makeContext();
    expect(() => renderPresetShape(a.ctx, 'parallelogram', 0, 0, 200, 100, [], '#base', null, () => {}, {
      paintFill: (ctx) => {
        ctx.fillStyle = 'temporary';
        throw new Error('paint failed');
      },
    })).toThrow('paint failed');
    expect(a.fillStyle).toBe('initial');
    expect(a.events.at(-1)).toBe('restore');
  });

  it('preserves the legacy baseFill sequence when no painter is supplied', () => {
    const a = makeContext();
    renderPresetShape(a.ctx, 'can', 0, 0, 200, 100, [], '#base', () => {
      a.events.push('stroke');
    }, () => a.events.push('clearShadow'));

    expect(a.events.filter((event) => event === 'fill:#base')).toHaveLength(2);
    expect(a.events.filter((event) => event === 'fill:rgba(255,255,255,0.4)')).toHaveLength(1);
    expect(a.events.filter((event) => event === 'clearShadow')).toHaveLength(1);
    expect(a.events.filter((event) => event === 'stroke')).toHaveLength(1);
    expect(a.events.filter((event) => event === 'save')).toHaveLength(2);
    expect(a.events.filter((event) => event === 'restore')).toHaveLength(2);
    expect(a.fillStyle).toBe('initial');
  });
});
