import type { Presentation } from '@maxgent/ooxml/pptx';

import { replaceResolvedElement } from '../../adapters/pptx-json-adapter';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import type { MutationExecutionResult } from '../../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../../transport/officecli/constants';
import type { OfficeCliCommand } from '../../transport/officecli/types';
import { AddElementMutation } from '../add-element';
import {
  freezeTarget,
  resolveMutationTarget,
  resolveStableElementPath,
} from '../mutation-utils';

export interface RemoveElementMutationParams {
  readonly target: ElementRef;
}

export class RemoveElementMutation extends Mutation {
  readonly type = MUTATION_TYPES.REMOVE_ELEMENT;
  readonly target: ElementRef;

  constructor({ target }: RemoveElementMutationParams) {
    super();
    this.target = freezeTarget(target);
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    return {
      presentation: replaceResolvedElement(presentation, resolved, null),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): AddElementMutation | undefined {
    const resolved = resolveMutationTarget(presentation, this);
    if (resolved.element.type !== 'shape') return undefined;
    return new AddElementMutation({
      target: this.target,
      element: resolved.element,
      presentationElementIndex: resolved.presentationElementIndex,
    });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.REMOVE,
      path: resolveStableElementPath(presentation, this, context),
    });
  }
}
