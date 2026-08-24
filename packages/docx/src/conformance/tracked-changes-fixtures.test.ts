/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import type { DocxDocumentModel } from '../types.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { DeepReadonly, DocumentLayout, TextPlacement } from '../layout/types.js';
import { collectLayoutSourceCommentRanges, resolveCommentAnchorRuns } from '../comments.js';
import {
  collectLayoutSourceRevisionRanges,
  resolveRevisionAnchorRuns,
} from '../revisions.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { attachDocumentLayoutVariants } from '../layout/document-layout-variants.js';
import { textRunsForSelectedPage } from '../text-run-projection.js';
import { textRunSourceIndexForDocument } from '../layout/text-index.js';
import {
  generateAllStoryCommentsDocx,
  generateCommentedDocx,
  generateTrackedChangesDocx,
} from './generate.js';

// End-to-end over the REAL parser: the redistributable §17.13.5 / §17.13.4
// fixtures round-trip XML → WASM parse → layout, pinning final-state rendering,
// revision metadata, comment threading, and consumer-owned anchor projection.

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

function layoutOf(model: DocxDocumentModel): DeepReadonly<DocumentLayout> {
  const normalized = normalizeInternalDocumentModel(model).document;
  return layoutDocument(
    normalized,
    createLayoutServices(normalized, { measureContext: measureContext() }),
    { currentDateMs: 0 },
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
  it('lays out the accepted final state while retaining revision metadata', () => {
    const model = parseRaw(generateTrackedChangesDocx());
    const finalView = layoutOf(model);
    const finalText = fullText(finalView);
    expect(finalText).toContain('Kept inserted moved-in tail');
    expect(finalText).not.toContain('deleted');
    const revisions = model.body.flatMap((element) => element.type === 'paragraph'
      ? element.runs.flatMap((run) => run.type === 'text' && run.revision
          ? [run.revision.kind]
          : [])
      : []);
    expect(revisions).toEqual(['insertion', 'deletion', 'moveFrom', 'moveTo']);
  });
});

describe('commented fixture (§17.13.4 + commentsExtended, real parser)', () => {
  it('parses threading, resolved state, and joinable anchor ranges', () => {
    const model = parseRaw(generateCommentedDocx());
    const comments = model.comments ?? [];
    expect(comments.map((comment) => comment.id)).toEqual(['1', '2', '3']);
    expect(comments[1]!.parentId).toBe('1');
    expect(comments[2]!.resolved).toBe(true);
    const ranges = collectLayoutSourceCommentRanges(comments, layoutSourceStore(model));
    expect(ranges).toContainEqual(
      expect.objectContaining({
        commentId: '1',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        startRunIndex: 1,
        endRunIndex: 2,
        reference: {
          source: { story: 'body', storyInstance: 'body', path: [0] },
          runIndex: 2,
          affinity: 'following',
        },
      }),
    );
    // The commented run is unchanged text — geometry is identical whether or
    // not the anchors exist (zero-effect marks): the annotated paragraph's
    // full text survives layout untouched.
    const layout = layoutOf(model);
    expect(fullText(layout)).toContain('Before annotated text after.');
  });

  it('projects anchors and run geometry for all six retained stories', () => {
    const raw = parseRaw(generateAllStoryCommentsDocx());
    const normalized = normalizeInternalDocumentModel(raw).document;
    const source = layoutSourceStore(normalized);
    const ranges = collectLayoutSourceCommentRanges(normalized.comments ?? [], source);
    expect(new Set(ranges.map((range) => range.source.story))).toEqual(new Set([
      'body', 'header', 'footer', 'footnote', 'endnote', 'textbox',
    ]));

    const services = createLayoutServices(normalized, { measureContext: measureContext() });
    const variants = attachDocumentLayoutVariants({
      source,
      services,
      defaultCurrentDateMs: 0,
      buildLayout: (options) => layoutDocument(normalized, services, options),
    });
    const pageCount = variants.store.defaultLayout.pages.length;
    const runs = Array.from({ length: pageCount }, (_, pageIndex) =>
      textRunsForSelectedPage(services, pageIndex, {
        currentDate: 0,
        defaultCurrentDateMs: 0,
      })).flat();
    const geometricStories = new Set(runs.flatMap((run) => run.source ? [run.source.story] : []));
    expect(geometricStories).toEqual(new Set([
      'body', 'header', 'footer', 'footnote', 'endnote', 'textbox',
    ]));

    for (const anchor of ranges) {
      const hasCoveredRun = runs.some((run) => run.source
        && run.sourceRunIndex !== undefined
        && run.source.story === anchor.source.story
        && run.source.storyInstance === anchor.source.storyInstance
        && run.source.path.join('.') === anchor.source.path.join('.')
        && run.sourceRunIndex >= anchor.startRunIndex
        && run.sourceRunIndex < anchor.endRunIndex);
      expect(hasCoveredRun, `${anchor.source.story}:${anchor.commentId}`).toBe(true);
    }
  });

  it('keeps the public review UI sample backed by real threads, changes, and geometry', async () => {
    const bytes = await readFile(new URL('../../public/demo/sample-1.docx', import.meta.url));
    const raw = parseRaw(bytes);
    const normalized = normalizeInternalDocumentModel(raw).document;
    expect(normalized.comments).toHaveLength(4);
    expect(normalized.comments?.filter((comment) => comment.parentId)).toHaveLength(1);
    expect(normalized.comments?.filter((comment) => comment.resolved)).toHaveLength(1);
    expect(normalized.revisions).toHaveLength(4);
    expect(normalized.revisions?.map((revision) => revision.id)).toEqual(['102', '103', '100', '101']);

    const source = layoutSourceStore(normalized);
    const anchors = collectLayoutSourceCommentRanges(normalized.comments ?? [], source);
    expect(anchors).toHaveLength(3);
    const services = createLayoutServices(normalized, { measureContext: measureContext() });
    const variants = attachDocumentLayoutVariants({
      source,
      services,
      defaultCurrentDateMs: 0,
      buildLayout: (options) => layoutDocument(normalized, services, options),
    });
    const pageRuns = Array.from({ length: variants.store.defaultLayout.pages.length }, (_, pageIndex) =>
      textRunsForSelectedPage(services, pageIndex, {
        currentDate: 0,
        defaultCurrentDateMs: 0,
      }));
    const revisionAnchors = collectLayoutSourceRevisionRanges(
      normalized.revisions ?? [],
      source,
      textRunSourceIndexForDocument(variants.store.defaultLayout),
    );
    expect(revisionAnchors).toHaveLength(4);
    for (const anchor of anchors) {
      expect(pageRuns.some((runs) => resolveCommentAnchorRuns(anchor, runs).length > 0)).toBe(true);
    }
    const deletionTargets = revisionAnchors.flatMap((anchor) => {
      const revision = normalized.revisions?.[anchor.revisionIndex];
      if (revision?.kind !== 'deletion') return [];
      const target = pageRuns
        .flatMap((runs) => resolveRevisionAnchorRuns(anchor, runs))
        .map((run) => run.text)
        .join('');
      return [{ deleted: revision.text, target }];
    });
    expect(deletionTargets).toEqual([
      { deleted: 'road-noise', target: 'road noise' },
      { deleted: 'city pace', target: 'hurried pace' },
    ]);
  });
});
