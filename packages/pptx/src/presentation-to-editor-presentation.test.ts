import { describe, it, expect, vi } from 'vitest';

import { PptxPresentation } from './presentation';
import type { Slide } from './types';

/**
 * `toEditorPresentation` assembles editor JSON from preflight + the main-mode
 * slide repository. The constructor opens a real Worker, so tests build an
 * off-prototype instance and inject only the private fields the method reads.
 */
describe('PptxPresentation.toEditorPresentation', () => {
  const slide = (index: number): Slide => ({
    index,
    slideNumber: index + 1,
    partName: `ppt/slides/slide${index + 1}.xml`,
    background: null,
    elements: [],
    elementSources: [],
  });

  function make(args: {
    mode: 'main' | 'worker';
    slideCount?: number;
    withSlide?: (index: number, consume: (slide: Slide) => unknown) => Promise<unknown>;
    waitUntilLayoutComplete?: () => Promise<void>;
    replaceAll?: (slides: readonly Slide[]) => void;
  }) {
    const slideCount = args.slideCount ?? 2;
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = args.mode;
    instance._resourceFailure = null;
    instance._bootstrap = {
      slideCount,
      slideWidth: 9144000,
      slideHeight: 6858000,
      defaultTextColor: '383838',
      majorFont: 'Aptos Display',
      minorFont: 'Aptos',
      hlinkColor: '0563C1',
      folHlinkColor: null,
      embeddedFonts: [],
      slides: Array.from({ length: slideCount }, (_, index) => ({
        index,
        partName: `ppt/slides/slide${index + 1}.xml`,
      })),
    };
    instance._preflight = {
      slideCount,
      slideWidth: 9144000,
      slideHeight: 6858000,
      defaultTextColor: '383838',
      majorFont: 'Aptos Display',
      minorFont: 'Aptos',
      hlinkColor: '0563C1',
      folHlinkColor: null,
      embeddedFonts: [],
      slides: Array.from({ length: slideCount }, (_, index) => ({
        index,
        partName: `ppt/slides/slide${index + 1}.xml`,
        notes: null,
        hidden: false,
        mediaElements: [],
      })),
      fontPreloadNames: [],
    };
    instance._slides = {
      withSlide: args.withSlide ?? (async (index, consume) => consume(slide(index))),
      replaceAll: args.replaceAll ?? vi.fn(),
    };
    instance._availableSlideCount = slideCount;
    instance.waitUntilLayoutComplete = args.waitUntilLayoutComplete ?? vi.fn().mockResolvedValue(undefined);
    return instance as unknown as PptxPresentation;
  }

  it('assembles a detached Presentation from preflight and every slide', async () => {
    const loads: number[] = [];
    const original = slide(0);
    const pres = make({
      mode: 'main',
      withSlide: async (index, consume) => {
        loads.push(index);
        return consume(index === 0 ? original : slide(index));
      },
    });

    const model = await pres.toEditorPresentation();

    expect(loads).toEqual([0, 1]);
    expect(model).toEqual({
      slideWidth: 9144000,
      slideHeight: 6858000,
      slides: [slide(0), slide(1)],
      defaultTextColor: '383838',
      majorFont: 'Aptos Display',
      minorFont: 'Aptos',
      hlinkColor: '0563C1',
    });
    expect(model.slides[0]).not.toBe(original);
    model.slides[0].partName = 'mutated';
    expect(original.partName).toBe('ppt/slides/slide1.xml');
  });

  it('waits for progressive slide preparation before exporting', async () => {
    let ready = false;
    const pres = make({
      mode: 'main',
      waitUntilLayoutComplete: vi.fn(async () => { ready = true; }),
      withSlide: async (index, consume) => {
        expect(ready).toBe(true);
        return consume(slide(index));
      },
    });

    await pres.toEditorPresentation();
  });

  it('keeps bootstrap and paintable slide counts aligned after replacing the list', () => {
    const replaceAll = vi.fn();
    const pres = make({ mode: 'main', replaceAll });
    const replacement = [slide(0)];

    pres.replaceSlideList(replacement);

    expect(replaceAll).toHaveBeenCalledWith(replacement);
    expect(pres.slideCount).toBe(1);
    expect(pres.availableSlideCount).toBe(1);
  });

  it('rejects worker mode', async () => {
    const withSlide = vi.fn();
    const pres = make({ mode: 'worker', withSlide });
    await expect(pres.toEditorPresentation()).rejects.toThrow(/mode: 'worker'/);
    expect(withSlide).not.toHaveBeenCalled();
  });

  it('rejects when the presentation is not loaded', async () => {
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance._resourceFailure = null;
    instance._preflight = null;
    instance._slides = null;
    const pres = instance as unknown as PptxPresentation;
    await expect(pres.toEditorPresentation()).rejects.toThrow('Presentation not loaded');
  });
});
