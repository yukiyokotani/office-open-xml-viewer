/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import type { DocxDocumentModel } from '../types.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { DeepReadonly, DocumentLayout, TextPlacement } from '../layout/types.js';
import { collectDocumentCommentRanges, buildCommentThreads } from '../comment-margin-layout.js';
import { generateCommentedDocx, generateTrackedChangesDocx } from './generate.js';

// End-to-end over the REAL parser: the redistributable §17.13.5 / §17.13.4
// fixtures round-trip XML → WASM parse → layout, pinning the final-view
// default, the markup variant's decorations, and the comment threading +
// anchor-range projection the margin overlay consumes.

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 6,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return measureContext(); }
};

function parseRaw(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel;
  } finally {
    archive.free();
  }
}

function layoutOf(model: DocxDocumentModel, showTrackedChanges?: boolean): DeepReadonly<DocumentLayout> {
  const normalized = normalizeInternalDocumentModel(model).document;
  return layoutDocument(
    normalized,
    createLayoutServices(normalized, { measureContext: measureContext() }),
    { currentDateMs: 0, ...(showTrackedChanges === undefined ? {} : { showTrackedChanges }) },
  ) as DeepReadonly<DocumentLayout>;
}

function textPlacements(layout: DeepReadonly<DocumentLayout>): DeepReadonly<TextPlacement>[] {
  return layout.pages.flatMap((page) =>
    page.layers.body.flatMap((node) => node.kind === 'paragraph'
      ? node.lines.flatMap((line) =>
          line.placements.filter((placement) => placement.kind === 'text'))
      : [])) as DeepReadonly<TextPlacement>[];
}

function fullText(layout: DeepReadonly<DocumentLayout>): string {
  return textPlacements(layout).map((placement) => placement.text).join('');
}

beforeAll(async () => {
  const wasm = await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url));
  await init({ module_or_path: wasm });
});

describe('tracked-changes fixture (§17.13.5, real parser)', () => {
  it('lays out the final state by default and the full markup on demand', () => {
    const model = parseRaw(generateTrackedChangesDocx());
    const finalView = layoutOf(model);
    const finalText = fullText(finalView);
    expect(finalText).toContain('Kept inserted moved-in tail');
    expect(finalText).not.toContain('deleted');
    // Markup: the deletion and the move's source come back, decorated.
    const markupView = layoutOf(model, true);
    const markupText = fullText(markupView);
    expect(markupText).toContain('deleted');
    expect(markupText.split('moved-in').length - 1).toBe(2);
    const decorated = textPlacements(markupView);
    const inserted = decorated.find((placement) => placement.text.includes('inserted'))!;
    expect(inserted.decorations.some((d) => d.kind === 'underline')).toBe(true);
    const deleted = decorated.find((placement) => placement.text.includes('deleted'))!;
    expect(deleted.decorations.some((d) => d.kind === 'strikethrough')).toBe(true);
    // Change bars: at least the revision line carries a margin bar.
    expect((markupView.pages[0]!.changeBars ?? []).length).toBeGreaterThan(0);
    expect(finalView.pages[0]!.changeBars).toBeUndefined();
  });
});

describe('commented fixture (§17.13.4 + commentsExtended, real parser)', () => {
  it('parses threading, resolved state, and joinable anchor ranges', () => {
    const model = parseRaw(generateCommentedDocx());
    const comments = model.comments ?? [];
    expect(comments.map((comment) => comment.id)).toEqual(['1', '2', '3']);
    expect(comments[1]!.parentId).toBe('1');
    expect(comments[2]!.resolved).toBe(true);
    const threads = buildCommentThreads(comments);
    expect(threads.map((thread) => thread.root.id)).toEqual(['1']);
    expect(threads[0]!.replies.map((reply) => reply.id)).toEqual(['2']);
    const ranges = collectDocumentCommentRanges(model.body);
    expect(ranges).toContainEqual(
      expect.objectContaining({
        commentId: '1',
        paragraphPath: [0],
        startRunIndex: 1,
        endRunIndex: 2,
      }),
    );
    // The commented run is unchanged text — geometry is identical whether or
    // not the anchors exist (zero-effect marks): the annotated paragraph's
    // full text survives layout untouched.
    const layout = layoutOf(model);
    expect(fullText(layout)).toContain('Before annotated text after.');
  });
});
