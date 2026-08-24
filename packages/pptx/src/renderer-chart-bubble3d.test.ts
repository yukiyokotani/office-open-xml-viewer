import { describe, expect, it, vi } from 'vitest';
import type { ChartExRenderer, ChartModel } from '@silurus/ooxml-core';
import { renderSlide } from './renderer.js';
import type { Slide } from './types.js';

function recordingCanvas() {
  const rotations: number[] = [];
  const radialGradients: Array<{ args: number[]; stops: Array<[number, string]> }> = [];
  const compositeModes: string[] = [];
  const state: Record<string, unknown> = {
    fillStyle: '#000000', strokeStyle: '#000000', globalAlpha: 1,
    globalCompositeOperation: 'source-over', font: '10px sans-serif',
    lineWidth: 1, textAlign: 'start', textBaseline: 'alphabetic',
  };
  const context = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      if (property === 'measureText') return (text: string) => ({ width: String(text).length * 6 });
      if (property === 'getLineDash') return () => [];
      if (property === 'rotate') return (angle: number) => { rotations.push(angle); };
      if (property === 'createLinearGradient') return () => ({ addColorStop() {} });
      if (property === 'createRadialGradient') return (...args: number[]) => {
        const gradient = { args, stops: [] as Array<[number, string]> };
        radialGradients.push(gradient);
        return {
          addColorStop(position: number, color: string) {
            gradient.stops.push([position, color]);
          },
        };
      };
      return vi.fn();
    },
    set(target, property: string, value) {
      target[property] = value;
      if (property === 'globalCompositeOperation') compositeModes.push(String(value));
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  const canvas = {
    width: 0, height: 0, style: {}, offsetWidth: 960,
    getContext: () => context,
  } as unknown as HTMLCanvasElement;
  return { canvas, rotations, radialGradients, compositeModes };
}

describe('PPTX rotated bubble3D chart', () => {
  it('rotates the chart frame while keeping the complete material in local coordinates', async () => {
    const chart = {
      chartType: 'bubble', categories: ['1'], showLegend: false,
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
      series: [{
        name: 'Bubble', color: '4472C4', values: [1], bubbleSizes: [100], bubble3D: true,
      }],
    } as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [{
        type: 'chart', x: 500_000, y: 500_000, width: 4_000_000, height: 3_000_000,
        rotation: 90, flipH: false, flipV: false, chart,
      }],
    } as Slide;
    const rec = recordingCanvas();

    await renderSlide(rec.canvas, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1 });

    expect(rec.rotations).toContain(Math.PI / 2);
    expect(rec.radialGradients).toHaveLength(3);
    expect(rec.radialGradients.map(gradient => gradient.stops.length)).toEqual([4, 5, 6]);
    expect(rec.compositeModes.filter(mode => mode === 'source-atop')).toHaveLength(3);
  });

  it('forwards the opt-in ChartEx module through the PPTX chart element boundary', async () => {
    const chart = {
      chartType: 'boxWhisker', categories: [], series: [], showLegend: false,
      chartexBox: { categories: [], series: [] },
    } as unknown as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [{
        type: 'chart', x: 500_000, y: 500_000, width: 4_000_000, height: 3_000_000,
        rotation: 0, flipH: false, flipV: false, chart,
      }],
    } as Slide;
    const rec = recordingCanvas();
    const render = vi.fn<ChartExRenderer['render']>(() => true);
    const chartEx = { render } as ChartExRenderer;

    await renderSlide(rec.canvas, slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, chartEx,
    });

    expect(render).toHaveBeenCalledOnce();
    expect(render.mock.calls[0]?.[1]).toBe(chart);
    expect(render.mock.calls[0]?.[4]).toBe(0);
  });
});
