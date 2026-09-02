import { describe, expect, it } from 'vitest';
import { PptxFontPreloadAccumulator, pptxFontPreloadNames } from './font-plan.js';
import type { Presentation, Slide } from './types.js';

describe('PPTX font plan', () => {
  it('includes slide-local paragraph and run families', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        textBody: { paragraphs: [{
          defFontFamily: 'Franklin Gothic Medium',
          runs: [{
            type: 'text', text: 'Title', fontFamily: null,
            fontFamilyEa: 'Yu Gothic', fontFamilySym: null,
          }],
        }] },
      }],
    } as unknown as Slide;
    const accumulator = new PptxFontPreloadAccumulator('Aptos Display', 'Aptos');
    accumulator.addSlide(slide);

    expect(accumulator.names()).toEqual(expect.arrayContaining([
      'Aptos Display', 'Aptos', 'Franklin Gothic Medium', 'Yu Gothic',
    ]));
  });

  it('matches full-presentation script planning incrementally', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [
        { type: 'shape', textBody: { paragraphs: [{ runs: [{ type: 'text', text: '日本語' }] }] } },
        { type: 'table', rows: [{ cells: [{ textBody: { paragraphs: [{ runs: [{ type: 'text', text: 'العربية' }] }] } }] }] },
        {
          type: 'chart',
          chart: {
            title: 'Заголовок', categories: ['หมวด'], series: [{ name: 'סדרה' }],
          },
        },
      ],
    } as unknown as Slide;
    const presentation: Presentation = {
      slideWidth: 1,
      slideHeight: 1,
      slides: [slide],
      defaultTextColor: null,
      majorFont: 'Yu Gothic',
      minorFont: 'Aptos',
    };
    const incremental = new PptxFontPreloadAccumulator(
      presentation.majorFont,
      presentation.minorFont,
    );
    incremental.addSlide(slide);

    expect(incremental.names()).toEqual(pptxFontPreloadNames(presentation));
    expect(incremental.names()).toEqual([
      'Yu Gothic', 'Aptos',
      'Noto Sans JP', 'Noto Serif JP',
      'Noto Sans', 'Noto Serif',
      'Noto Naskh Arabic', 'Noto Sans Arabic',
      'Noto Sans Thai',
      'Noto Sans Hebrew', 'Noto Serif Hebrew',
    ]);
  });
});
