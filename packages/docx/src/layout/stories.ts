import { layoutFlowBlocks } from './flow.js';
import { LayoutInvariantError } from './diagnostics.js';
import {
  translateBorder,
  translateCompleteParagraphLayout,
  translateRect,
  translateTableLayout,
  type LayoutTranslation,
} from './retained-geometry-translation.js';
import type {
  BlockLayoutAlgorithms,
  LayoutServices,
  NoteLayout,
  NoteSeparatorLayout,
  StoryLayout,
  StoryBlockInput,
  StoryLayoutInput,
  ParsedUnsupportedTextBoxBlock,
} from './types.js';

const storyBlockAlgorithms = new WeakMap<LayoutServices, BlockLayoutAlgorithms>();
const isUnsupportedTextBoxBlock = (
  block: StoryBlockInput,
): block is ParsedUnsupportedTextBoxBlock =>
  'type' in block && block.type === 'unsupportedTextBoxBlock';

/**
 * Associate document/session-private paragraph and table acquisition with one
 * immutable service view. The stable LayoutServices data contract stays free of
 * renderer handles while layoutStory keeps the plan-mandated two-argument API.
 */
export function attachStoryBlockLayoutAlgorithms(
  services: LayoutServices,
  algorithms: BlockLayoutAlgorithms,
): void {
  if (storyBlockAlgorithms.has(services)) {
    throw new Error('Story block layout algorithms are already attached');
  }
  storyBlockAlgorithms.set(services, Object.freeze({ ...algorithms }));
}

export function translateStoryLayout(
  story: StoryLayout,
  delta: LayoutTranslation,
): StoryLayout {
  return Object.freeze({
    ...story,
    flowBounds: translateRect(story.flowBounds, delta),
    inkBounds: translateRect(story.inkBounds, delta),
    ...(story.clipBounds ? { clipBounds: translateRect(story.clipBounds, delta) } : {}),
    blocks: Object.freeze(story.blocks.map((block) => {
      if (block.kind === 'paragraph') return translateCompleteParagraphLayout(block, delta);
      if (block.kind === 'table') return translateTableLayout(block, delta);
      throw new Error(`Story contains unsupported retained node: ${block.kind}`);
    })),
  });
}

function translateNoteSeparatorLayout(
  occurrence: NoteSeparatorLayout,
  delta: LayoutTranslation,
): NoteSeparatorLayout {
  return Object.freeze({
    ...occurrence,
    flowBounds: translateRect(occurrence.flowBounds, delta),
    ...(occurrence.paragraph
      ? { paragraph: translateCompleteParagraphLayout(occurrence.paragraph, delta) } : {}),
    ...(occurrence.followingParagraphs ? {
      followingParagraphs: Object.freeze(occurrence.followingParagraphs.map((block) =>
        translateCompleteParagraphLayout(block, delta))),
    } : {}),
    ...(occurrence.rule ? {
      rule: Object.freeze({ ...occurrence.rule, segment: translateBorder(occurrence.rule.segment, delta) }),
    } : {}),
  });
}

/** Translate a note and every geometry it owns, including its native
 * leading separator and trailing continuation notice occurrences. */
export function translateNoteLayout(
  note: NoteLayout,
  delta: LayoutTranslation,
): NoteLayout {
  return Object.freeze({
    ...note,
    flowBounds: translateRect(note.flowBounds, delta),
    inkBounds: translateRect(note.inkBounds, delta),
    ...(note.clipBounds ? { clipBounds: translateRect(note.clipBounds, delta) } : {}),
    separator: Object.freeze(note.separator.map((segment) => translateBorder(segment, delta))),
    ...(note.leading ? { leading: translateNoteSeparatorLayout(note.leading, delta) } : {}),
    story: translateStoryLayout(note.story, delta),
    ...(note.trailing ? { trailing: translateNoteSeparatorLayout(note.trailing, delta) } : {}),
  });
}

/** Lay out one OOXML story through the same paragraph/table flow dispatcher as body content. */
export function layoutStory(
  input: StoryLayoutInput,
  services: LayoutServices,
): StoryLayout {
  for (const block of input.blocks) {
    if (isUnsupportedTextBoxBlock(block)) continue;
    if (
      block.source.story !== input.source.story
      || block.source.storyInstance !== input.source.storyInstance
    ) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `Story block ${block.source.story}:${block.source.storyInstance} is not owned by `
          + `${input.source.story}:${input.source.storyInstance}`,
      );
    }
  }
  const algorithms = storyBlockAlgorithms.get(services);
  if (!algorithms) {
    throw new Error('Story block layout algorithms are not attached to the supplied services');
  }
  const unsupported = input.blocks.filter(
    isUnsupportedTextBoxBlock,
  );
  const supported = input.blocks.filter(
    (block): block is Exclude<StoryBlockInput, ParsedUnsupportedTextBoxBlock> =>
      !isUnsupportedTextBoxBlock(block),
  );
  const flow = layoutFlowBlocks({
    blocks: supported,
    container: input.container,
    cursor: {
      xPt: input.container.bounds.xPt,
      yPt: input.container.bounds.yPt,
    },
    source: input.source,
  }, services, algorithms);
  return Object.freeze({
    story: input.source.story,
    flowBounds: flow.flowBounds,
    inkBounds: flow.inkBounds,
    ...(flow.clipBounds ? { clipBounds: flow.clipBounds } : {}),
    blocks: Object.freeze([...flow.blocks]),
    advancePt: flow.advancePt,
    diagnostics: Object.freeze(unsupported.map((block) => Object.freeze({
      code: 'UNSUPPORTED_FEATURE' as const,
      severity: 'warning' as const,
      source: Object.freeze({
        story: input.source.story,
        storyInstance: input.source.storyInstance,
        path: Object.freeze([...block.sourcePath]),
      }),
      message: `Unsupported text-box block ${block.qName}`,
    }))),
  });
}
