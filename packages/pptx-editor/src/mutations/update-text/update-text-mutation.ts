import type { Presentation, ShapeElement, TextBody } from '@maxgent/ooxml/pptx';

import {
  replaceResolvedElement,
  replaceTextBodyPlainText,
} from '../../adapters/pptx-json-adapter';
import type { NonEmptyReadonlyArray } from '../../domain/command';
import { ELEMENT_ORIGINS } from '../../domain/element-origin';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../../domain/mutation';
import { MUTATION_TYPES } from '../../domain/mutation-types';
import { textNotEditable } from '../../engine/mutation-engine-utils';
import type { MutationExecutionResult } from '../../engine/types';
import { OFFICECLI_COMMAND_TYPES } from '../../transport/officecli/constants';
import type { OfficeCliCommand } from '../../transport/officecli/types';
import {
  freezeProps,
  freezeTarget,
  officeCliError,
  plainTextOf,
  resolveMutationTarget,
  resolveStableParagraphPath,
  resolveStableElementPath,
} from '../mutation-utils';
import {
  applyTextStyleEdit,
  applyTextStylePatch,
  canInvertTextBodyPlainTextReplacement,
  captureInverseShapeStylePatch,
  captureInverseTextStyleEdits,
  captureTextStylePatchAtScope,
  formatOfficeCliRange,
  freezeStyle,
  freezeTextStyleEdit,
  hasNullClearableStyleKeys,
  hasStyleKeys,
  materializeShapeStyleForOfficeCli,
  materializeTextStyleEditForOfficeCli,
  type TextStyleEdit,
  type TextStylePatch,
} from './text-editing';

export type {
  TextSpan,
  TextScope,
  TextStyleEdit,
  TextStylePatch,
} from './text-editing';

export interface UpdateTextMutationParams {
  readonly target: ElementRef;
  /**
   * 整框纯文本替换（`\n` → 段落）。与 `style` / `edits` 至少提供一类；
   * 与 `edits` 互斥。
   */
  readonly value?: string;
  /**
   * 整框统一样式。与 `value` / `edits` 至少提供一类；
   * 通常与 `edits` 互斥，但 inverse 可附带仅含 `verticalAlign` 的顶层 style。
   */
  readonly style?: TextStylePatch;
  /**
   * 增量段落/选区编辑。每条可含 `text`（整段文案）和/或 `style`。
   * 与顶层 `value` 互斥；与顶层 `style` 互斥，除非 `style` 仅含 `verticalAlign`。
   * 段落 `text` 与 span 编辑不能出现在同一个 mutation 中。
   * 字符坐标相对修改前的 run 拼接纯文本（与 OfficeCLI `range` 一致）。
   */
  readonly edits?: readonly TextStyleEdit[];
}

export class UpdateTextMutation extends Mutation {
  readonly type = MUTATION_TYPES.UPDATE_TEXT;
  readonly target: ElementRef;
  readonly value: string | undefined;
  readonly style: TextStylePatch | undefined;
  readonly edits: readonly TextStyleEdit[] | undefined;

  constructor({ target, value, style, edits }: UpdateTextMutationParams) {
    super();
    const hasEdits = edits !== undefined && edits.length > 0;
    const hasStyle = style !== undefined && hasStyleKeys(style);
    if (hasEdits && value !== undefined) {
      throw new TypeError(
        'UpdateTextMutation edits cannot be combined with top-level value',
      );
    }
    if (hasEdits && hasStyle && !isVerticalAlignOnlyStyle(style)) {
      throw new TypeError(
        'UpdateTextMutation edits can only be combined with a verticalAlign-only top-level style',
      );
    }
    if (!hasEdits && value === undefined && !hasStyle) {
      throw new TypeError(
        'UpdateTextMutation requires value, a non-empty style patch, and/or non-empty edits',
      );
    }
    if (edits !== undefined && edits.length === 0) {
      throw new TypeError('UpdateTextMutation edits must be non-empty when provided');
    }

    const frozenEdits = edits === undefined
      ? undefined
      : Object.freeze(edits.map(freezeTextStyleEdit));
    if (
      frozenEdits?.some((edit) => edit.text !== undefined)
      && frozenEdits.some((edit) => (
        edit.text === undefined
        && (edit.scope.kind === 'spans' || edit.scope.spans !== undefined)
      ))
    ) {
      throw new TypeError(
        'UpdateTextMutation paragraph text replacement cannot be combined with span edits',
      );
    }

    this.target = freezeTarget(target);
    this.value = value;
    this.style = style === undefined ? undefined : freezeStyle(style);
    this.edits = frozenEdits;
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    const resolved = resolveMutationTarget(presentation, this);
    if (resolved.element.type !== 'shape' || !resolved.element.textBody) {
      throw textNotEditable(this);
    }

    let textBody: TextBody = resolved.element.textBody;
    if (this.edits) {
      for (const edit of this.edits) {
        textBody = applyTextStyleEdit(textBody, edit);
      }
      if (this.style) {
        textBody = applyTextStylePatch(textBody, this.style);
      }
    } else {
      if (this.value !== undefined) {
        const replaced = replaceTextBodyPlainText(textBody, this.value);
        if (!replaced) throw textNotEditable(this);
        textBody = replaced;
      }
      if (this.style) {
        textBody = applyTextStylePatch(textBody, this.style);
      }
    }

    return {
      presentation: replaceResolvedElement(
        presentation,
        resolved,
        { ...resolved.element, textBody } satisfies ShapeElement,
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(presentation: Presentation): UpdateTextMutation | undefined {
    const { element } = resolveMutationTarget(presentation, this);
    if (element.type !== 'shape' || !element.textBody) {
      throw textNotEditable(this);
    }

    if (this.edits) {
      const inverseEdits = captureInverseTextStyleEdits(element.textBody, this.edits);
      if (!inverseEdits || inverseEdits.length === 0) return undefined;
      if (this.style) {
        const boxStyle = captureTextStylePatchAtScope(
          element.textBody,
          { kind: 'shape' },
          this.style,
        );
        if (!boxStyle) return undefined;
        return new UpdateTextMutation({
          target: this.target,
          edits: inverseEdits,
          style: boxStyle,
        });
      }
      return new UpdateTextMutation({
        target: this.target,
        edits: inverseEdits,
      });
    }

    if (
      this.value !== undefined
      && !canInvertTextBodyPlainTextReplacement(element.textBody)
    ) return undefined;

    const value = this.value === undefined
      ? undefined
      : plainTextOf(element.textBody);
    if (this.value !== undefined && value === undefined) return undefined;

    if (this.style !== undefined && this.value === undefined) {
      const inverseStyle = captureInverseShapeStylePatch(element.textBody, this.style);
      if (!inverseStyle) return undefined;
      return new UpdateTextMutation({
        target: this.target,
        ...inverseStyle,
      });
    }

    const style = this.style === undefined
      ? undefined
      : captureTextStylePatchAtScope(element.textBody, { kind: 'shape' }, this.style);
    if (this.style !== undefined && style === undefined) return undefined;

    return new UpdateTextMutation({
      target: this.target,
      value,
      style,
    });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand | NonEmptyReadonlyArray<OfficeCliCommand> {
    if (this.target.origin !== ELEMENT_ORIGINS.SLIDE) {
      throw officeCliError(
        'target.unsupportedOrigin',
        context,
        this,
        `OfficeCLI translation does not support ${this.target.origin} elements`,
      );
    }

    const { element } = resolveMutationTarget(presentation, this);
    if (element.type !== 'shape' || !element.textBody) {
      throw textNotEditable(this);
    }
    const inheritance = {
      shapeDefaultTextColor: element.defaultTextColor,
      presentationDefaultTextColor: presentation.defaultTextColor,
      presentationMinorFont: presentation.minorFont,
      presentationMajorFont: presentation.majorFont,
    };

    if (this.edits) {
      const commands: OfficeCliCommand[] = [];
      for (const edit of this.edits) {
        let materialized: readonly TextStyleEdit[];
        try {
          materialized = materializeTextStyleEditForOfficeCli(
            element.textBody,
            edit,
            inheritance,
          );
        } catch (error) {
          throw officeCliError(
            'value.unsupportedFidelity',
            context,
            this,
            error instanceof Error ? error.message : String(error),
          );
        }
        for (const piece of materialized) {
          commands.push(this.#editToOfficeCli(presentation, context, piece));
        }
      }
      if (this.style) {
        commands.push(Object.freeze({
          command: OFFICECLI_COMMAND_TYPES.SET,
          path: resolveStableElementPath(presentation, this, context, 'shape'),
          props: freezeProps(styleToOfficeCliProps(this.style, context, this)),
        }));
      }
      const [first, ...rest] = commands;
      return Object.freeze([first, ...rest]) as NonEmptyReadonlyArray<OfficeCliCommand>;
    }

    const props: Record<string, string> = {};
    if (this.value !== undefined) props.text = this.value;

    if (this.style) {
      // value 先改段落结构；null style 的 resolve-then-set 必须基于替换后的 textBody，
      // 否则会按旧 p[N] 拆命令（例如两段收成一段后仍生成 /p[2]）。
      let textBodyForStyle = element.textBody;
      if (this.value !== undefined) {
        const replaced = replaceTextBodyPlainText(element.textBody, this.value);
        if (!replaced) throw textNotEditable(this);
        textBodyForStyle = replaced;
      }

      let materialized;
      try {
        materialized = materializeShapeStyleForOfficeCli(
          textBodyForStyle,
          this.style,
          inheritance,
        );
      } catch (error) {
        throw officeCliError(
          'value.unsupportedFidelity',
          context,
          this,
          error instanceof Error ? error.message : String(error),
        );
      }

      if ('edits' in materialized) {
        const commands: OfficeCliCommand[] = [];
        if (Object.keys(props).length > 0) {
          commands.push(Object.freeze({
            command: OFFICECLI_COMMAND_TYPES.SET,
            path: resolveStableElementPath(presentation, this, context, 'shape'),
            props: freezeProps(props),
          }));
        }
        const paragraphCountOverride = this.value !== undefined
          ? textBodyForStyle.paragraphs.length
          : undefined;
        for (const edit of materialized.edits) {
          commands.push(this.#editToOfficeCli(
            presentation,
            context,
            edit,
            paragraphCountOverride,
          ));
        }
        if (commands.length === 0) {
          throw officeCliError(
            'value.invalidText',
            context,
            this,
            'UpdateTextMutation produced an empty OfficeCLI property set',
          );
        }
        const [first, ...rest] = commands;
        return Object.freeze([first, ...rest]) as NonEmptyReadonlyArray<OfficeCliCommand>;
      }

      Object.assign(props, styleToOfficeCliProps(materialized.style, context, this));
    }

    if (Object.keys(props).length === 0) {
      throw officeCliError(
        'value.invalidText',
        context,
        this,
        'UpdateTextMutation produced an empty OfficeCLI property set',
      );
    }

    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.SET,
      path: resolveStableElementPath(presentation, this, context, 'shape'),
      props: freezeProps(props),
    });
  }

  #editToOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
    edit: TextStyleEdit,
    paragraphCountOverride?: number,
  ): OfficeCliCommand {
    const styleProps = edit.style
      ? styleToOfficeCliProps(edit.style, context, this)
      : {};
    const props: Record<string, string> = { ...styleProps };
    if (edit.text !== undefined) {
      props.text = edit.text;
    }
    if (Object.keys(props).length === 0) {
      throw officeCliError(
        'value.invalidText',
        context,
        this,
        'TextStyleEdit produced an empty OfficeCLI property set',
      );
    }

    if (edit.scope.kind === 'spans') {
      if (edit.text !== undefined) {
        throw officeCliError(
          'value.unsupportedFidelity',
          context,
          this,
          'TextStyleEdit text is only supported on paragraph scope',
        );
      }
      return Object.freeze({
        command: OFFICECLI_COMMAND_TYPES.SET,
        path: resolveStableElementPath(presentation, this, context, 'shape'),
        props: freezeProps({
          range: formatOfficeCliRange(edit.scope.spans),
          ...styleProps,
        }),
      });
    }

    const path = resolveStableParagraphPath(
      presentation,
      this,
      context,
      edit.scope.paragraphIndex,
      paragraphCountOverride,
    );
    if (edit.scope.spans) {
      if (edit.text !== undefined) {
        throw officeCliError(
          'value.unsupportedFidelity',
          context,
          this,
          'TextStyleEdit text cannot be combined with paragraph spans',
        );
      }
      props.range = formatOfficeCliRange(edit.scope.spans);
    }
    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.SET,
      path,
      props: freezeProps(props),
    });
  }
}

const HEX_COLOR = /^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/;

function styleToOfficeCliProps(
  style: TextStylePatch,
  context: MutationCommandContext,
  mutation: UpdateTextMutation,
): Record<string, string> {
  if (hasNullClearableStyleKeys(style)) {
    throw officeCliError(
      'value.unsupportedFidelity',
      context,
      mutation,
      'Clear-to-inherit style keys must be resolve-then-set before OfficeCLI translation',
    );
  }

  const props: Record<string, string> = {};

  if ('bold' in style && style.bold !== undefined) {
    props.bold = String(style.bold);
  }
  if ('italic' in style && style.italic !== undefined) {
    props.italic = String(style.italic);
  }
  if (style.underline !== undefined) {
    props.underline = style.underline === false
      ? 'none'
      : style.underline === 'single'
        ? 'single'
        : 'double';
  }
  if (style.strikethrough !== undefined) {
    props.strike = style.strikethrough === false
      ? 'none'
      : style.strikethrough === 'single'
        ? 'single'
        : 'double';
  }
  if ('fontSize' in style) {
    if (!Number.isFinite(style.fontSize) || (style.fontSize as number) <= 0) {
      throw officeCliError(
        'value.invalidText',
        context,
        mutation,
        `Invalid fontSize ${String(style.fontSize)}`,
      );
    }
    props.size = `${style.fontSize}pt`;
  }
  if ('color' in style) {
    if (typeof style.color !== 'string' || !HEX_COLOR.test(style.color)) {
      throw officeCliError(
        'value.invalidText',
        context,
        mutation,
        `Text color ${String(style.color)} is not a plain hex color`,
      );
    }
    props.color = style.color.slice(0, 6);
  }
  if ('fontFamily' in style) {
    props.font = style.fontFamily as string;
  }
  if ('fontFamilyEa' in style) {
    props['font.ea'] = style.fontFamilyEa as string;
  }
  if (style.caps !== undefined) props.cap = style.caps;
  if ('letterSpacing' in style) {
    if (!Number.isFinite(style.letterSpacing)) {
      throw officeCliError(
        'value.invalidText',
        context,
        mutation,
        `Invalid letterSpacing ${String(style.letterSpacing)}`,
      );
    }
    // letterSpacing 为 pt（与 parser/renderer 一致）。OfficeCLI spacing 也按 pt
    // 解释并写入 rPr@spc（×100）；勿再 /100，否则乐观模型与重解析会差 100 倍。
    props.spacing = String(style.letterSpacing as number);
  }
  if ('highlight' in style) {
    if (style.highlight == null) {
      props.highlight = 'none';
    } else if (!HEX_COLOR.test(style.highlight)) {
      throw officeCliError(
        'value.invalidText',
        context,
        mutation,
        `Highlight color ${style.highlight} is not a plain hex color`,
      );
    } else {
      props.highlight = style.highlight.slice(0, 6);
    }
  }
  if (style.align !== undefined) {
    props.align = (
      {
        l: 'left',
        ctr: 'center',
        r: 'right',
        just: 'justify',
      } as const
    )[style.align];
  }
  if (style.verticalAlign !== undefined) {
    props.valign = (
      {
        t: 'top',
        ctr: 'middle',
        b: 'bottom',
      } as const
    )[style.verticalAlign];
  }

  return props;
}

function isVerticalAlignOnlyStyle(style: TextStylePatch | undefined): boolean {
  if (!style) return false;
  const keys = Object.keys(style);
  return keys.length === 1 && keys[0] === 'verticalAlign';
}
