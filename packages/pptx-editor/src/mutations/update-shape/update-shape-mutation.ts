import type { Fill, Presentation, ShapeElement, Stroke } from '@maxgent/ooxml/pptx';

import { replaceResolvedElement } from '../../adapters/pptx-json-adapter';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import { createUnchangedResult } from '../../engine/mutation-engine-utils';
import { MutationExecutionError } from '../../engine/errors';
import type { MutationExecutionResult } from '../../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../../transport/officecli/constants';
import type { OfficeCliCommand } from '../../transport/officecli/types';
import {
  freezeProps,
  freezeTarget,
  officeCliError,
  resolveMutationTarget,
  resolveStableElementPath,
} from '../mutation-utils';
import { applyShapeFillProps, applyShapeStrokeProps } from '../shape-officecli';
import type { ShapePatch } from './interface';

export interface UpdateShapeMutationParams {
  readonly target: ElementRef;
  readonly value: ShapePatch;
}

export class UpdateShapeMutation extends Mutation {
  readonly type = MUTATION_TYPES.UPDATE_SHAPE;
  readonly target: ElementRef;
  readonly value: ShapePatch;

  constructor({ target, value }: UpdateShapeMutationParams) {
    super();
    if (Object.keys(value).length === 0) {
      throw new TypeError('UpdateShapeMutation requires at least one shape property');
    }
    this.target = freezeTarget(target);
    this.value = freezeShapePatch(value);
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    const shape = requireShape(resolved.element, this);
    if (!shapePatchChanges(shape, this.value)) return createUnchangedResult(presentation);

    return {
      presentation: replaceResolvedElement(
        presentation,
        resolved,
        { ...shape, ...this.value },
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): UpdateShapeMutation {
    const { element } = resolveMutationTarget(presentation, this);
    const shape = requireShape(element, this);
    const value: Record<string, unknown> = {};
    for (const key of Object.keys(this.value) as Array<keyof ShapePatch>) {
      value[key] = shape[key];
    }
    return new UpdateShapeMutation({
      target: this.target,
      value: value as ShapePatch,
    });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    const { element } = resolveMutationTarget(presentation, this);
    requireShape(element, this);
    const props: Record<string, string> = {};
    appendTransformProps(props, this.value, context, this);
    if ('fill' in this.value) {
      applyShapeFillProps(
        props,
        this.value.fill ?? { fillType: 'none' },
        context,
        this,
      );
    }
    if ('stroke' in this.value) {
      if (this.value.stroke == null) props.line = 'none';
      else applyShapeStrokeProps(props, this.value.stroke, context, this);
    }

    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.SET,
      path: resolveStableElementPath(presentation, this, context, 'shape'),
      props: freezeProps(props),
    });
  }
}

function requireShape(
  element: { type: string },
  mutation: UpdateShapeMutation,
): ShapeElement {
  if (element.type !== 'shape') {
    throw new MutationExecutionError(
      'element.unsupportedElement',
      mutation,
      `Element ${mutation.target.elementId} is not a shape`,
    );
  }
  return element as ShapeElement;
}

function appendTransformProps(
  props: Record<string, string>,
  value: ShapePatch,
  context: MutationCommandContext,
  mutation: UpdateShapeMutation,
): void {
  for (const key of ['x', 'y', 'width', 'height'] as const) {
    if (!(key in value)) continue;
    const coordinate = value[key] as number;
    if (!Number.isSafeInteger(coordinate) || ((key === 'width' || key === 'height') && coordinate < 0)) {
      throw officeCliError(
        'value.invalidTransform',
        context,
        mutation,
        'OfficeCLI transform requires safe-integer EMUs and non-negative dimensions',
      );
    }
    props[key] = `${coordinate}emu`;
  }
  if ('rotation' in value) {
    if (!Number.isFinite(value.rotation)) {
      throw officeCliError(
        'value.invalidTransform',
        context,
        mutation,
        'OfficeCLI transform requires a finite rotation',
      );
    }
    props.rotation = String(value.rotation);
  }
  if ('flipH' in value) props.flipH = String(value.flipH);
  if ('flipV' in value) props.flipV = String(value.flipV);
}

function shapePatchChanges(shape: ShapeElement, value: ShapePatch): boolean {
  return (Object.keys(value) as Array<keyof ShapePatch>)
    .some((key) => JSON.stringify(shape[key]) !== JSON.stringify(value[key]));
}

function freezeShapePatch(value: ShapePatch): ShapePatch {
  return Object.freeze({
    ...value,
    ...('fill' in value ? { fill: cloneFill(value.fill) } : {}),
    ...('stroke' in value ? { stroke: cloneStroke(value.stroke) } : {}),
  });
}

function cloneFill(fill: Fill | null | undefined): Fill | null | undefined {
  if (fill == null) return fill;
  if (fill.fillType === 'gradient') {
    return Object.freeze({
      ...fill,
      stops: Object.freeze(fill.stops.map((stop) => Object.freeze({ ...stop }))),
      fillToRect: fill.fillToRect && Object.freeze({ ...fill.fillToRect }),
      tileRect: fill.tileRect && Object.freeze({ ...fill.tileRect }),
    }) as Fill;
  }
  if (fill.fillType === 'image') {
    return Object.freeze({
      ...fill,
      fillRect: fill.fillRect && Object.freeze({ ...fill.fillRect }),
      tile: fill.tile && Object.freeze({ ...fill.tile }),
      duotone: fill.duotone && Object.freeze({ ...fill.duotone }),
    });
  }
  return Object.freeze({ ...fill });
}

function cloneStroke(stroke: Stroke | null | undefined): Stroke | null | undefined {
  if (stroke == null) return stroke;
  return Object.freeze({
    ...stroke,
    ...(stroke.fill ? { fill: cloneFill(stroke.fill) as Stroke['fill'] } : {}),
    customDash: stroke.customDash
      && Object.freeze(stroke.customDash.map((item) => Object.freeze({ ...item }))),
    headEnd: stroke.headEnd && Object.freeze({ ...stroke.headEnd }),
    tailEnd: stroke.tailEnd && Object.freeze({ ...stroke.tailEnd }),
  });
}
