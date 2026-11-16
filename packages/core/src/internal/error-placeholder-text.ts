import type { ResolvedErrorPlaceholder, WireErrorPlaceholder } from './part-failures.js';

/**
 * Text of the "could not be displayed" placeholder. Only a failed part needs
 * it, so no format entry imports this module statically: the DOCX error page
 * and the PPTX and XLSX canvas cards that use it are loaded with the first
 * failed part (#2248).
 *
 * The module imports nothing at run time. A module of the static graph that a
 * lazily loaded chunk imports is moved, by the bundler, into a static chunk of
 * its own (#2249), which costs every load more than the lazy code saves. What
 * this code needs from the static graph is handed in as
 * {@link ErrorPlaceholderTextHost} instead.
 */

/** The functions of the static graph the placeholder text is built on. */
export interface ErrorPlaceholderTextHost {
  /** `clampPartFailureMessage` with its default budget (`part-failures.ts`). */
  readonly clampPartFailureMessage: (message: string) => string;
  /** `graphemeClusterOffsets` (`text/sea-break.ts`). */
  readonly graphemeClusterOffsets: (text: string) => number[];
}

/** Read a wire field back into resolved text (`builtin`/absent → `null`).
 *  `resolve-on-main` must be handled by the caller first. @internal */
export function resolvedFromWire(wire: WireErrorPlaceholder | undefined): ResolvedErrorPlaceholder | null {
  if (wire === undefined || wire.state === 'builtin') return null;
  if (wire.state !== 'resolved') throw new Error('errorPlaceholder was not resolved on the main thread');
  const title = typeof wire.title === 'string' ? wire.title : null;
  const detail = wire.detail === null || typeof wire.detail === 'string' ? wire.detail : undefined;
  return Object.freeze({ title, detail: detail === '' ? null : detail });
}

/** Lines painted for a placeholder detail; the title is one line. */
export const ERROR_PLACEHOLDER_DETAIL_LINES = 4;

interface PlaceholderMeasureContext {
  readonly font: string;
  measureText(text: string): { readonly width: number };
}

const CANVAS_LINE_CACHE_ENTRIES = 8;
const canvasLineCaches = new WeakMap<object, Map<string, readonly string[]>>();

/** The placeholder text functions, bound to the static graph's `host`. */
export function errorPlaceholderText(host: ErrorPlaceholderTextHost) {
  /** The painted title for a failed part, at most
   *  `PART_FAILURE_MESSAGE_MAX_UTF8` bytes. */
  function errorPlaceholderTitle(resolved: ResolvedErrorPlaceholder | null, builtin: string): string {
    return host.clampPartFailureMessage(resolved?.title ?? builtin);
  }

  /** The painted detail for a failed part; `null` paints no detail line.
   *  Caller text and the file-derived parser message are both capped at
   *  `PART_FAILURE_MESSAGE_MAX_UTF8` bytes: at most a few lines are
   *  painted, so wrapping never has to scan an unbounded string. */
  function errorPlaceholderDetail(
    resolved: ResolvedErrorPlaceholder | null,
    builtin: string,
  ): string | null {
    const detail = resolved === null || resolved.detail === undefined ? builtin : resolved.detail;
    return detail === null ? null : host.clampPartFailureMessage(detail);
  }

  /**
   * Word-wrap placeholder text to `maxWidth` in at most `maxLines` lines. A word
   * wider than a line is broken at grapheme boundaries, and text that does not
   * fit ends in `…`. Shared by the DOCX error page (shaped at layout) and the
   * PPTX/XLSX canvas cards so every format truncates the same way.
   */
  function wrapErrorPlaceholderText(
    text: string,
    maxWidth: number,
    maxLines: number,
    widthOf: (value: string) => number,
  ): string[] {
    const words = text.trim().split(/\s+/).filter(Boolean);
    const lines: string[] = [];
    let line = '';
    const ellipsize = (value: string): string => {
      const ellipsis = '…';
      const offsets = [0, ...host.graphemeClusterOffsets(value), value.length]
        .filter((offset, index, values) => index === 0 || offset !== values[index - 1]);
      let end = offsets.length - 1;
      while (end > 0 && widthOf(`${value.slice(0, offsets[end])}${ellipsis}`) > maxWidth) end -= 1;
      return `${value.slice(0, offsets[end] ?? 0)}${ellipsis}`;
    };
    const truncateLast = (): void => {
      if (lines.length === 0) lines.push(ellipsize(''));
      else lines[lines.length - 1] = ellipsize(lines[lines.length - 1]);
    };
    const pushOverlong = (word: string): boolean => {
      const offsets = [0, ...host.graphemeClusterOffsets(word), word.length]
        .filter((offset, index, values) => index === 0 || offset !== values[index - 1]);
      let start = 0;
      while (start < offsets.length - 1) {
        if (lines.length >= maxLines) {
          truncateLast();
          return true;
        }
        let end = start + 1;
        while (end < offsets.length && widthOf(word.slice(offsets[start], offsets[end])) <= maxWidth) end += 1;
        const fittingEnd = Math.max(start + 1, end - 1);
        const chunk = word.slice(offsets[start], offsets[fittingEnd]);
        if (lines.length >= maxLines - 1 && fittingEnd < offsets.length - 1) {
          lines.push(ellipsize(chunk));
          return true;
        }
        lines.push(chunk);
        start = fittingEnd;
      }
      return false;
    };
    for (let wordIndex = 0; wordIndex < words.length; wordIndex += 1) {
      const word = words[wordIndex];
      const candidate = line ? `${line} ${word}` : word;
      const candidateWidth = widthOf(candidate);
      if (line && candidateWidth > maxWidth) {
        lines.push(line);
        line = '';
        if (lines.length >= maxLines) {
          truncateLast();
          break;
        }
        if (widthOf(word) > maxWidth) {
          if (pushOverlong(word)) break;
        }
        else line = word;
      } else if (!line && candidateWidth > maxWidth) {
        if (pushOverlong(word)) break;
      } else {
        line = candidate;
      }
      if (wordIndex < words.length - 1 && lines.length >= maxLines && !line) {
        truncateLast();
        break;
      }
    }
    if (line && lines.length < maxLines) lines.push(line);
    return lines;
  }

  /**
   * {@link wrapErrorPlaceholderText} measured with the context's current font.
   * A failed part repaints with the same text, so each context keeps its last
   * few results instead of re-measuring every word on every paint.
   */
  function errorPlaceholderCanvasLines(
    ctx: PlaceholderMeasureContext,
    text: string,
    maxWidth: number,
    maxLines: number,
  ): readonly string[] {
    let cache = canvasLineCaches.get(ctx);
    if (cache === undefined) {
      cache = new Map();
      canvasLineCaches.set(ctx, cache);
    }
    const key = `${ctx.font}\u0000${maxWidth}\u0000${maxLines}\u0000${text}`;
    const hit = cache.get(key);
    if (hit !== undefined) {
      cache.delete(key);
      cache.set(key, hit);
      return hit;
    }
    const lines = Object.freeze(wrapErrorPlaceholderText(text, maxWidth, maxLines, (value) => ctx.measureText(value).width));
    cache.set(key, lines);
    if (cache.size > CANVAS_LINE_CACHE_ENTRIES) cache.delete(cache.keys().next().value as string);
    return lines;
  }

  return { errorPlaceholderTitle, errorPlaceholderDetail, wrapErrorPlaceholderText, errorPlaceholderCanvasLines };
}

export type ErrorPlaceholderTextFunctions = ReturnType<typeof errorPlaceholderText>;
