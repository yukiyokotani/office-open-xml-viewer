import { afterEach, describe, expect, it, vi } from 'vitest';
import { installDom, makeEl } from '../viewer-destroy-test-dom.js';
import { CanvasSurface, SheetOverlayHost } from './sheet-surface.js';

afterEach(() => vi.unstubAllGlobals());

describe('XLSX sheet surface roles', () => {
  it('CanvasSurface owns focus-scoped input listener teardown and local coordinates', () => {
    installDom();
    const canvas = makeEl('canvas');
    const area = makeEl('div');
    const input = makeEl('div');
    const surface = new CanvasSurface(
      canvas as unknown as HTMLCanvasElement,
      area as unknown as HTMLDivElement,
      input as unknown as HTMLDivElement,
    );
    const pointer = vi.fn();
    const keyboard = vi.fn();
    surface.on('pointerdown', pointer);
    surface.on('keydown', keyboard);
    input.dispatch('pointerdown', {});
    input.dispatch('keydown', {});
    expect(pointer).toHaveBeenCalledOnce();
    expect(keyboard).toHaveBeenCalledOnce();

    surface.destroy();
    input.dispatch('pointerdown', {});
    input.dispatch('keydown', {});
    expect(pointer).toHaveBeenCalledOnce();
    expect(keyboard).toHaveBeenCalledOnce();
  });

  it('SheetOverlayHost owns the shared stacking order and overlay visibility', () => {
    installDom();
    const area = makeEl('div');
    const canvas = makeEl('canvas');
    const input = makeEl('div');
    const overlays = new SheetOverlayHost(
      area as unknown as HTMLDivElement,
      canvas as unknown as HTMLCanvasElement,
      input as unknown as HTMLDivElement,
      { commentMaxWidth: 280, commentMaxHeight: 200, validationMaxWidth: 240, validationMaxHeight: 200 },
    );
    expect(area.children).toEqual([
      canvas,
      overlays.selection,
      overlays.find,
      input,
      overlays.comment,
      overlays.commentStatus,
      overlays.validation,
    ]);
    overlays.announceComment('Comment on B2 by Ada: Review this');
    expect(overlays.commentStatus.textContent).toBe('Comment on B2 by Ada: Review this');
    overlays.showComment(12, 34);
    expect(overlays.comment.style.left).toBe('12px');
    expect(overlays.comment.style.top).toBe('34px');
    overlays.hideComment();
    expect(overlays.comment.style.display).toBe('none');
    expect(overlays.commentStatus.textContent).toBe('');
  });
});
