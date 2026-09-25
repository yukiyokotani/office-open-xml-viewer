/**
 * Renderer-owned host-layout measurement for realms that load the XLSX
 * renderer lazily (the Node facade). It measures a model source's Normal font
 * with the same `computeMdw` that sizes the painted grid.
 */
import { computeMdw } from '../renderer.js';
import type { HostLayoutFont } from './host-layout.js';

export function measureHostLayoutFont(
  font: HostLayoutFont,
  context: CanvasRenderingContext2D,
): number {
  return computeMdw(
    font.family,
    font.sizePt,
    undefined,
    false,
    font.bold ? 700 : 400,
    font.italic ? 'italic' : 'normal',
    context,
  );
}
