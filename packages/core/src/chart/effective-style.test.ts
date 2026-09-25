import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import {
  chartDataPointStyleRole,
  effectiveChartStyleRole,
  withEffectiveChartStyleRoles,
} from './effective-style.js';

describe('effective classic chart style cascade', () => {
  it('keeps numeric line and text when a linked role supplies only fill', () => {
    expect(effectiveChartStyleRole(
      { fillColors: ['111111'], lineColors: ['222222'], lineWidthEmu: 12_700,
        fontColor: '333333' },
      { fillColors: ['AAAAAA'] },
    )).toEqual({
      fillColors: ['AAAAAA'],
      lineColors: ['222222'],
      lineWidthEmu: 12_700,
      fontColor: '333333',
    });
  });

  it('treats linked NoStyle as fallthrough and linked noFill as suppression', () => {
    const numeric = { fillColors: ['111111'], lineColors: ['222222'] };
    expect(effectiveChartStyleRole(numeric, {
      fillNoStyle: true, lineNoStyle: true,
    })).toEqual(numeric);
    expect(effectiveChartStyleRole(numeric, {
      fillHidden: true, fillPaintAuthored: true,
      lineHidden: true, linePaintAuthored: true,
    })).toEqual({
      fillHidden: true, fillPaintAuthored: true,
      lineHidden: true, linePaintAuthored: true,
    });
  });

  it('keeps numeric paint while applying geometry from a linked NoStyle line', () => {
    expect(effectiveChartStyleRole(
      { lineColors: ['111111'], lineWidthEmu: 12_700, lineDash: 'solid' },
      {
        lineHidden: true,
        lineNoStyle: true,
        lineWidthEmu: 25_400,
        lineDash: 'dash',
        lineDashAuthored: true,
      },
    )).toEqual({
      lineColors: ['111111'],
      lineWidthEmu: 25_400,
      lineDash: 'dash',
      lineDashAuthored: true,
    });
  });

  it('preserves a numeric NoStyle line sentinel until linked paint replaces it', () => {
    const numeric = {
      lineHidden: true,
      lineNoStyle: true,
      lineWidthEmu: 12_700,
    };
    expect(effectiveChartStyleRole(numeric, {
      fillColors: ['AAAAAA'],
    })).toEqual({
      fillColors: ['AAAAAA'],
      lineHidden: true,
      lineNoStyle: true,
      lineWidthEmu: 12_700,
    });
    expect(effectiveChartStyleRole(numeric, {
      lineColors: ['BBBBBB'],
    })).toEqual({
      lineColors: ['BBBBBB'],
      lineWidthEmu: 12_700,
    });
  });

  it('preserves linked style-entry override modifiers independently of paint', () => {
    expect(effectiveChartStyleRole(
      { fillColors: ['111111'], lineColors: ['222222'] },
      {
        shapePropertiesPresent: true,
        allowNoFillOverride: true,
        allowNoLineOverride: false,
      },
    )).toEqual({
      fillColors: ['111111'],
      lineColors: ['222222'],
      shapePropertiesPresent: true,
      allowNoFillOverride: true,
      allowNoLineOverride: false,
    });
  });

  it('retains numeric line geometry under linked line paint', () => {
    expect(effectiveChartStyleRole(
      { lineColors: ['111111'], lineWidthEmu: 25_400, lineDash: 'dash' },
      { lineColors: ['AAAAAA'] },
    )).toEqual({
      lineColors: ['AAAAAA'], lineWidthEmu: 25_400, lineDash: 'dash',
    });
  });

  it('retains numeric line paint under linked line geometry', () => {
    expect(effectiveChartStyleRole(
      { lineColors: ['111111'], lineWidthEmu: 12_700, lineDash: 'solid' },
      { lineWidthEmu: 25_400, lineDash: 'dash' },
    )).toEqual({
      lineColors: ['111111'], lineWidthEmu: 25_400, lineDash: 'dash',
    });
  });

  it('replaces the numeric line palette domain only with linked paint', () => {
    expect(effectiveChartStyleRole(
      { lineColors: ['111111', '222222'], lineFormattingIndices: [10, 20] },
      { lineColors: ['AAAAAA', 'BBBBBB'], lineWidthEmu: 25_400 },
    )).toEqual({
      lineColors: ['AAAAAA', 'BBBBBB'], lineWidthEmu: 25_400,
    });
    expect(effectiveChartStyleRole(
      { lineColors: ['111111', '222222'], lineFormattingIndices: [10, 20] },
      { lineWidthEmu: 25_400 },
    )).toEqual({
      lineColors: ['111111', '222222'], lineFormattingIndices: [10, 20],
      lineWidthEmu: 25_400,
    });
  });

  it('treats preset/custom dash as one choice across style layers', () => {
    expect(effectiveChartStyleRole(
      {
        lineColors: ['111111'],
        lineCustomDash: [{ dash: 2, space: 1 }],
        lineDashAuthored: true,
      },
      { lineDash: 'dash', lineDashAuthored: true },
    )).toEqual({
      lineColors: ['111111'],
      lineDash: 'dash',
      lineDashAuthored: true,
    });
  });

  it('does not expose numeric text paint through an authored linked paint', () => {
    expect(effectiveChartStyleRole(
      { fontColor: '111111', fontSizeHpt: 2_000 },
      { fontPaintAuthored: true, fontFace: 'Aptos' },
    )).toEqual({
      fontPaintAuthored: true, fontSizeHpt: 2_000, fontFace: 'Aptos',
    });
  });

  it('replaces effects as one component and preserves only linked NoStyle fallthrough', () => {
    const numeric = {
      shadows: [{ color: '111111', alpha: 1, blur: 1, dist: 2, dir: 3 }],
      effectAuthored: true,
    };
    expect(effectiveChartStyleRole(numeric, {
      effectNoStyle: true,
    })).toEqual(numeric);
    expect(effectiveChartStyleRole(numeric, {
      effectAuthored: true,
    })).toEqual({ effectAuthored: true });
    expect(effectiveChartStyleRole(numeric, {
      effectAuthored: true,
      effectUnsupported: true,
    })).toEqual({ effectAuthored: true, effectUnsupported: true });
  });

  it('adapts classic mark roles without replacing the source layers', () => {
    const chart = {
      chartStyleRoles: { dataPoint: { fillColors: ['AAAAAA'] } },
      classicChartStyleRoles: {
        dataPoint: { fillColors: ['111111'], lineColors: ['222222'] },
        dataPointLine: { lineColors: ['333333'] },
      },
    } as ChartModel;
    const effective = withEffectiveChartStyleRoles(chart);
    expect(effective.chartStyleRoles?.dataPoint).toEqual({
      fillColors: ['AAAAAA'], lineColors: ['222222'],
    });
    expect(effective.chartexDataPointStyle).toEqual(effective.chartStyleRoles?.dataPoint);
    expect(effective.chartexDataPointLineStyle).toEqual(
      effective.chartStyleRoles?.dataPointLine,
    );
    expect(effective.classicChartStyleRoles).toBe(chart.classicChartStyleRoles);
    expect(effective.linkedChartStyleRoles).toBe(chart.chartStyleRoles);
    expect(withEffectiveChartStyleRoles(effective).chartStyleRoles)
      .toEqual(effective.chartStyleRoles);
  });

  it('replaces raw classic linked aliases with linked-over-numeric adapters', () => {
    const rawLinked = { fillNoStyle: true, fillPaintAuthored: true };
    const rawLine = { lineNoStyle: true, linePaintAuthored: true };
    const chart = {
      chartStyleRoles: { dataPoint: rawLinked, dataPointLine: rawLine },
      classicChartStyleRoles: {
        dataPoint: { fillColors: ['112233'] },
        dataPointLine: { lineColors: ['445566'], lineWidthEmu: 25_400 },
      },
      // Classic parsers historically populated these with the raw linked
      // roles. They must not bypass numeric fallback in shared painters.
      chartexDataPointStyle: rawLinked,
      chartexDataPointLineStyle: rawLine,
    } as ChartModel;
    const effective = withEffectiveChartStyleRoles(chart);
    expect(effective.chartexDataPointStyle).toEqual({ fillColors: ['112233'] });
    expect(effective.chartexDataPointLineStyle).toEqual({
      lineColors: ['445566'], lineWidthEmu: 25_400,
    });
  });

  it('selects point-domain roles only for the owning varying plot group', () => {
    const chart = withEffectiveChartStyleRoles({
      chartType: 'clusteredBar',
      series: [{ values: [1, 2] }, { values: [3, 4], chartexFormatIdx: 10 }],
      plotGroups: [
        { kind: 'bar', seriesStart: 0, seriesCount: 1, categoryAxis: 'primary',
          valueAxis: 'primary', seriesAxis: 'none', varyColors: true },
        { kind: 'line', seriesStart: 1, seriesCount: 1, categoryAxis: 'primary',
          valueAxis: 'primary', seriesAxis: 'none' },
      ],
      classicChartStyleRoles: { dataPoint: { fillColors: ['SERIES'] } },
      classicVaryingPointChartStyleRoles: { dataPoint: { fillColors: ['POINT'] } },
    } as ChartModel);
    expect(chartDataPointStyleRole(chart, 'dataPoint', 0)?.fillColors).toEqual(['POINT']);
    expect(chartDataPointStyleRole(chart, 'dataPoint', 1)?.fillColors).toEqual(['SERIES']);
  });

  it('keeps point-domain roles and refusal sentinels local to each group', () => {
    const chart = withEffectiveChartStyleRoles({
      chartType: 'line',
      series: [{ values: [1, 2, 3] }, { values: [4, 5] }],
      plotGroups: [
        { kind: 'line', seriesStart: 0, seriesCount: 1, categoryAxis: 'primary',
          valueAxis: 'primary', seriesAxis: 'none', varyColors: true },
        { kind: 'radar', seriesStart: 1, seriesCount: 1, categoryAxis: 'primary',
          valueAxis: 'primary', seriesAxis: 'none', varyColors: true },
      ],
      classicChartStyleRoles: { dataPoint: { fillColors: ['SERIES'] } },
      classicVaryingPointChartStyleRolesByGroup: [
        { dataPoint: { fillColors: ['LINE-0', 'LINE-1', 'LINE-2'] } },
        {},
      ],
    } as ChartModel);
    expect(chartDataPointStyleRole(chart, 'dataPoint', 0)?.fillColors).toEqual([
      'LINE-0', 'LINE-1', 'LINE-2',
    ]);
    expect(chartDataPointStyleRole(chart, 'dataPoint', 1)).toBeUndefined();
  });
});
