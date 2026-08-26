import type {
  Presentation,
  SlideElement,
  TextBody,
} from '@maxgent/ooxml/pptx';

import {
  getElementSources,
  getSlideMutationId,
  resolveElementRef,
  type ResolvedElementRef,
} from '../adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../domain/element-origin';
import type {
  ElementMutation,
  ElementRef,
  Mutation,
  MutationCommandContext,
  MutationTarget,
  SlideRef,
} from '../domain/mutation';
import { MutationExecutionError } from '../engine/errors';
import { OfficeCliTranslatorError } from '../transport/officecli/errors';
import type { OfficeCliProps } from '../transport/officecli/types';

export type ResolvedMutationTarget = ResolvedElementRef;

export function resolveMutationTarget(
  presentation: Presentation,
  mutation: ElementMutation,
): ResolvedMutationTarget {
  const slide = presentation.slides.find(
    (candidate) => getSlideMutationId(candidate) === mutation.target.slideId,
  );
  if (!slide) {
    throw new MutationExecutionError(
      'slide.notFound',
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw new MutationExecutionError(
      'element.unsupportedOrigin',
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  if (!getElementSources(slide)) {
    throw new MutationExecutionError(
      'element.metadataUnavailable',
      mutation,
      `Slide ${mutation.target.slideId} has no complete element source metadata`,
    );
  }

  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    throw new MutationExecutionError(
      'element.notFound',
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }
  if (resolved.source.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw new MutationExecutionError(
      'element.unsupportedOrigin',
      mutation,
      `Editing ${resolved.source.origin} elements is not supported`,
    );
  }
  return resolved;
}

export function resolveStableElementPath(
  presentation: Presentation,
  mutation: ElementMutation,
  context: MutationCommandContext,
  expectedType?: SlideElement['type'],
): string {
  const slide = presentation.slides.find(
    (candidate) => getSlideMutationId(candidate) === mutation.target.slideId,
  );
  if (!slide) {
    throw officeCliError(
      'target.slideNotFound',
      context,
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  if (!getElementSources(slide)) {
    throw officeCliError(
      'target.metadataUnavailable',
      context,
      mutation,
      `Slide ${mutation.target.slideId} has no complete element source metadata`,
    );
  }

  const resolved = resolveElementRef(presentation, mutation.target);
  if (!resolved) {
    throw officeCliError(
      'target.elementNotFound',
      context,
      mutation,
      `Cannot resolve ${mutation.target.slideId}/${mutation.target.elementId}`,
    );
  }

  if (expectedType && resolved.element.type !== expectedType) {
    throw officeCliError(
      'target.unsupportedElement',
      context,
      mutation,
      `Expected ${expectedType}, got ${resolved.element.type}`,
    );
  }
  if (resolved.source.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${resolved.source.origin} elements is not supported`,
    );
  }
  const segment = stableElementPathSegment(resolved.element);
  if (!segment) {
    throw officeCliError(
      'target.unsupportedElement',
      context,
      mutation,
      `OfficeCLI cannot address ${resolved.element.type} elements by stable id`,
    );
  }
  if (!resolved.element.id || !/^\d+$/.test(resolved.element.id)) {
    throw officeCliError(
      'target.unstableElementId',
      context,
      mutation,
      `Element ${mutation.target.elementId} has no stable numeric OOXML id`,
    );
  }

  return `/slide[${resolved.slideIndex + 1}]/${segment}[@id=${resolved.element.id}]`;
}

/** 0-based paragraphIndex → OfficeCLI 1-based `/p[N]` under the stable shape path. */
export function resolveStableParagraphPath(
  presentation: Presentation,
  mutation: ElementMutation,
  context: MutationCommandContext,
  paragraphIndex: number,
  /**
   * 当同一次 mutation 先改写了 text（段落数变化）时，用替换后的段落数做边界检查，
   * 而不是修改前 presentation 上的旧结构。
   */
  paragraphCountOverride?: number,
): string {
  if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) {
    throw officeCliError(
      'value.invalidText',
      context,
      mutation,
      `Invalid paragraphIndex ${paragraphIndex}`,
    );
  }
  const shapePath = resolveStableElementPath(presentation, mutation, context, 'shape');
  const resolved = resolveElementRef(presentation, mutation.target);
  const paragraphCount = paragraphCountOverride ?? (
    resolved?.element.type === 'shape'
      ? resolved.element.textBody?.paragraphs.length ?? 0
      : 0
  );
  if (paragraphIndex >= paragraphCount) {
    throw officeCliError(
      'value.invalidText',
      context,
      mutation,
      `paragraphIndex ${paragraphIndex} is out of range for ${paragraphCount} paragraphs`,
    );
  }
  return `${shapePath}/p[${paragraphIndex + 1}]`;
}

function stableElementPathSegment(
  element: SlideElement,
): 'shape' | 'picture' | 'table' | 'chart' | undefined {
  switch (element.type) {
    case 'shape': return 'shape';
    case 'picture': return 'picture';
    case 'table': return 'table';
    case 'chart': return 'chart';
    // OfficeCLI 1.0.139 addresses video/audio by ordinal (`/video[M]`,
    // `/audio[M]`) and does not expose a cNvPr `@id` selector.
    case 'media': return undefined;
  }
}

export function resolveStableSlidePath(
  presentation: Presentation,
  mutation: ElementMutation,
  context: MutationCommandContext,
): string {
  const slideIndex = presentation.slides.findIndex(
    (slide) => getSlideMutationId(slide) === mutation.target.slideId,
  );
  if (slideIndex < 0) {
    throw officeCliError(
      'target.slideNotFound',
      context,
      mutation,
      `Cannot resolve slide ${mutation.target.slideId}`,
    );
  }
  if (mutation.target.origin !== ELEMENT_ORIGINS.SLIDE) {
    throw officeCliError(
      'target.unsupportedOrigin',
      context,
      mutation,
      `Editing ${mutation.target.origin} elements is not supported`,
    );
  }
  return `/slide[${slideIndex + 1}]`;
}

export function officeCliError(
  code: OfficeCliTranslatorError['code'],
  context: MutationCommandContext,
  mutation: Mutation,
  message: string,
): OfficeCliTranslatorError {
  return new OfficeCliTranslatorError(
    code,
    context.commandId,
    context.mutationIndex,
    mutation,
    message,
  );
}

export function freezeProps(props: Record<string, string>): OfficeCliProps {
  return Object.freeze(props);
}

export function freezeTarget(target: ElementRef): ElementRef;
export function freezeTarget(target: SlideRef): SlideRef;
export function freezeTarget(target: MutationTarget): MutationTarget {
  return Object.freeze({ ...target });
}

export function plainTextOf(textBody: TextBody): string | undefined {
  const paragraphs: string[] = [];
  for (const paragraph of textBody.paragraphs) {
    let text = '';
    for (const run of paragraph.runs) {
      if (run.type === 'math') return undefined;
      text += run.type === 'break' ? '\n' : run.text;
    }
    paragraphs.push(text);
  }
  return paragraphs.join('\n');
}
