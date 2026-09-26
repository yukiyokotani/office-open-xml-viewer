import { describe, expect, it } from 'vitest';
import { pathFillModeOverlay } from './index';

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
