// Test-only source module: the ODP source with `odfRowHeights` removed from
// every pulled slide, standing in for an application-supplied provider whose
// model omits the field.
import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import { openModelSource as openOdp } from '../odp-source-module.js';

function stripRowHeightMarkers(bytes: Uint8Array): Uint8Array {
  const slide = JSON.parse(new TextDecoder().decode(bytes)) as unknown;
  JSON.stringify(slide, (_key, value: unknown) => {
    if (value && typeof value === 'object' && 'odfRowHeights' in value) {
      delete (value as { odfRowHeights?: boolean }).odfRowHeights;
    }
    return value;
  });
  return new TextEncoder().encode(JSON.stringify(slide));
}

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
  transfer?: readonly Transferable[],
  context?: ModelSourceOpenContext,
) {
  const opened = await openOdp(bytes, config, signal, transfer, context);
  const archive = new Proxy(opened.archive, {
    get(target, key) {
      const value = Reflect.get(target, key, target) as unknown;
      if (key === 'pull_slide' && typeof value === 'function') {
        return (...args: unknown[]) => stripRowHeightMarkers((value as (...a: unknown[]) => Uint8Array).apply(target, args));
      }
      return typeof value === 'function' ? (value as (...a: unknown[]) => unknown).bind(target) : value;
    },
  });
  return Object.freeze({ archive, close: () => opened.close() });
}
