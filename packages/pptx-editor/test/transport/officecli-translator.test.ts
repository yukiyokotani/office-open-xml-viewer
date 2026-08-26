import { describe, expect, it } from 'vitest';

import type {
  ChartElement,
  MediaElement,
  PictureElement,
  Presentation,
  TableElement,
} from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import { applyCommand } from '../../src/engine/mutation-engine';
import { createUndoRedoEntry } from '../../src/history/command-inverter';
import { AddElementMutation } from '../../src/mutations/add-element';
import { RemoveElementMutation } from '../../src/mutations/remove-element';
import { UpdateShapeMutation } from '../../src/mutations/update-shape';
import { UpdateTextMutation } from '../../src/mutations/update-text';
import {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_VERSION,
} from '../../src/transport/officecli/constants';
import { OfficeCliTranslatorError } from '../../src/transport/officecli/errors';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import { deck, plainShape, shape } from '../fixtures/presentation';

describe('toOfficeCliBatch', () => {
  it('translates a complete transform to explicit OfficeCLI values', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'transform-1',
      mutations: [new UpdateShapeMutation({
        target: ref,
        value: {
          x: 914400,
          y: 457200,
          width: 1828800,
          height: 914400,
          rotation: 45,
          flipH: true,
          flipV: false,
        },
      })],
    });

    expect(batch).toEqual({
      schemaVersion: OFFICECLI_BATCH_SCHEMA_VERSION,
      officecliVersion: OFFICECLI_VERSION,
      commandId: 'transform-1',
      commands: [{
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: {
          x: '914400emu',
          y: '457200emu',
          width: '1828800emu',
          height: '914400emu',
          rotation: '45',
          flipH: 'true',
          flipV: 'false',
        },
      }],
    });
  });

  it('translates AddElement to an add command carrying a 1-based zorder instead of a top-level index', () => {
    const existing = shape('7', 'kept');
    const restored = shape('9', 'restored');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: restored,
        presentationElementIndex: 1,
      })],
    });

    expect(batch.commands).toEqual([{
      command: 'add',
      parent: '/slide[1]',
      type: 'shape',
      props: {
        id: '9',
        zorder: '2',
        preset: 'rect',
        x: '0emu',
        y: '0emu',
        width: '10emu',
        height: '10emu',
        rotation: '0',
        flipH: 'false',
        flipV: 'false',
        text: 'restored',
      },
    }]);
    expect(batch.commands[0]).not.toHaveProperty('index');
  });

  it('translates fill, outline, and shadow fidelity into the officecli styling grammar', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };
    const styled = shape('9', 'styled', {
      fill: { fillType: 'solid', color: 'FF000080' },
      stroke: {
        color: '0000FF40',
        width: 19050,
        dashStyle: 'lgDashDot',
        lineCap: 'butt',
        cmpd: 'dbl',
        headEnd: { type: 'triangle', w: 'med', len: 'med' },
        tailEnd: { type: 'none', w: 'med', len: 'med' },
      },
      shadow: { color: '808080', alpha: 0.4, blur: 50800, dist: 38100, dir: 45 },
    });

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-styled-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: styled,
        presentationElementIndex: 1,
      })],
    });

    expect(batch.commands[0]).toMatchObject({
      props: expect.objectContaining({
        fill: 'FF0000',
        opacity: '0.501961',
        line: '0000FF:1.5',
        lineOpacity: '0.25098',
        lineDash: 'lgDashDot',
        lineCap: 'flat',
        cmpd: 'dbl',
        headEnd: 'triangle',
        shadow: '808080-4-45-3-40',
      }),
    });
    expect((batch.commands[0] as { props: Record<string, string> }).props)
      .not.toHaveProperty('tailEnd');
  });

  it('translates two-stop linear gradients and pattern fills', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    const gradientBatch = toOfficeCliBatch(presentation, {
      id: 'add-gradient-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'grad', {
          fill: {
            fillType: 'gradient',
            gradType: 'linear',
            angle: 45,
            stops: [
              { position: 0, color: 'FF0000' },
              { position: 1, color: '0000FF' },
            ],
          },
        }),
        presentationElementIndex: 1,
      })],
    });
    expect(gradientBatch.commands[0]).toMatchObject({
      props: expect.objectContaining({ gradient: 'LINEAR;FF0000;0000FF;45' }),
    });

    const patternBatch = toOfficeCliBatch(presentation, {
      id: 'add-pattern-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'pat', {
          fill: { fillType: 'pattern', preset: 'diagBrick', fg: 'FF0000', bg: 'FFFFFF' },
        }),
        presentationElementIndex: 1,
      })],
    });
    expect(patternBatch.commands[0]).toMatchObject({
      props: expect.objectContaining({ pattern: 'diagBrick:FF0000:FFFFFF' }),
    });
  });

  it('degrades custGeom and adjust values instead of rejecting the restore', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };
    const mutation = new AddElementMutation({
      target: ref,
      element: shape('9', 'degraded', {
        geometry: 'custGeom',
        custGeom: [[]],
        adj: 16667,
        adj2: 40000,
      }),
      presentationElementIndex: 1,
    });

    // The snapshot itself is sanitized so the optimistic apply() result
    // matches what OfficeCLI will actually rebuild.
    expect(mutation.element).toMatchObject({
      geometry: 'rect',
      custGeom: null,
      adj: null,
      adj2: null,
    });

    const batch = toOfficeCliBatch(presentation, {
      id: 'add-degraded-1',
      mutations: [mutation],
    });
    const command = batch.commands[0] as { props: Record<string, string> };
    expect(command.props.preset).toBe('rect');
    expect(command.props).not.toHaveProperty('adj');
  });

  it.each([
    ['image fill(媒体字节不可达)', {
      fill: { fillType: 'image', imagePath: 'ppt/media/image1.png', mimeType: 'image/png' } as never,
    }],
    ['多停 gradient', {
      fill: {
        fillType: 'gradient',
        gradType: 'linear',
        angle: 45,
        stops: [
          { position: 0, color: 'FF0000' },
          { position: 0.5, color: '00FF00' },
          { position: 1, color: '0000FF' },
        ],
      } as never,
    }],
  ])('rejects restores that cannot round-trip faithfully: %s', (_label, overrides) => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const ref = {
      ...createElementRef(presentation.slides[0], existing, 0),
      elementId: '9',
    };

    expect(() => toOfficeCliBatch(presentation, {
      id: 'add-guarded-1',
      mutations: [new AddElementMutation({
        target: ref,
        element: shape('9', 'guarded', overrides),
        presentationElementIndex: 1,
      })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'value.unsupportedFidelity',
    }));
  });

  it('translates text style patches into OfficeCLI set props', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-style-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        value: 'after',
        style: {
          bold: true,
          italic: false,
          underline: 'double',
          strikethrough: 'single',
          fontSize: 18,
          color: 'FF0000AA',
          fontFamily: 'Arial',
          fontFamilyEa: '微软雅黑',
          caps: 'small',
          letterSpacing: 1.2,
          highlight: 'FFFF00',
          align: 'ctr',
          verticalAlign: 'b',
        },
      })],
    });

    expect(batch.commands[0]).toEqual({
      command: 'set',
      path: '/slide[1]/shape[@id=7]',
      props: {
        text: 'after',
        bold: 'true',
        italic: 'false',
        underline: 'double',
        strike: 'single',
        size: '18pt',
        color: 'FF0000',
        font: 'Arial',
        'font.ea': '微软雅黑',
        cap: 'small',
        spacing: '1.2',
        highlight: 'FFFF00',
        align: 'center',
        valign: 'bottom',
      },
    });
  });

  it('resolve-then-sets clear-to-inherit bold using paragraph/body defaults', () => {
    const target = shape('7', 'before');
    target.textBody!.paragraphs[0].defBold = true;
    const run = target.textBody!.paragraphs[0].runs[0];
    if (run.type === 'text') run.bold = false;
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-style-clear-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        style: { bold: null },
      })],
    });

    expect(batch.commands[0]).toEqual({
      command: 'set',
      path: '/slide[1]/shape[@id=7]',
      props: { bold: 'true' },
    });
  });

  it('splits clear-to-inherit whole-shape style when paragraph inheritance differs', () => {
    const first = shape('7', 'Hello');
    const second = shape('8', 'World');
    first.textBody!.paragraphs[0].defBold = true;
    second.textBody!.paragraphs[0].defBold = false;
    const target = shape('7', 'Hello');
    target.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-style-clear-split-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        style: { bold: null },
      })],
    });

    expect(batch.commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[1]',
        props: { bold: 'true' },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[2]',
        props: { bold: 'false' },
      },
    ]);
  });

  it('resolves value+null style against post-replace paragraphs, not pre-replace paths', () => {
    const first = shape('7', 'Hello');
    const second = shape('8', 'World');
    first.textBody!.paragraphs[0].defBold = true;
    second.textBody!.paragraphs[0].defBold = false;
    const target = shape('7', 'Hello');
    target.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-value-style-clear-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        value: 'one line',
        style: { bold: null },
      })],
    });

    // 替换后只剩一段（继承自原 p0），不应再生成旧的 /p[2]。
    expect(batch.commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: { text: 'one line', bold: 'true' },
      },
    ]);
  });

  it('emits post-replace paragraph style commands when value expands lines with mixed defs', () => {
    const first = shape('7', 'Hello');
    const second = shape('8', 'World');
    first.textBody!.paragraphs[0].defBold = true;
    second.textBody!.paragraphs[0].defBold = false;
    const target = shape('7', 'Hello');
    target.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    // 替换后 3 段：继承模板为 p0/p1/p1 → bold true/false/false，路径对应当前结构。
    const batch = toOfficeCliBatch(presentation, {
      id: 'text-value-style-expand-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        value: 'a\nb\nc',
        style: { bold: null },
      })],
    });

    expect(batch.commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: { text: 'a\nb\nc' },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[1]',
        props: { bold: 'true' },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[2]',
        props: { bold: 'false' },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[3]',
        props: { bold: 'false' },
      },
    ]);
  });

  it('rejects clear-to-inherit fontFamily when no inherited font is available', () => {
    const target = shape('7', 'before');
    target.textBody!.paragraphs[0].defFontFamily = null;
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'text-style-clear-font-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        style: { fontFamily: null },
      })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'value.unsupportedFidelity',
    }));
  });

  it('expands multi-span style edits into multiple OfficeCLI set commands with range', () => {
    const target = shape('7', 'Hello World');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-edits-1',
      mutations: [new UpdateTextMutation({
        target: ref,
        edits: [
          {
            scope: { kind: 'spans', spans: [{ start: 0, end: 5 }] },
            style: { bold: true, color: 'FF0000' },
          },
          {
            scope: { kind: 'paragraph', paragraphIndex: 0, spans: [{ start: 6, end: 11 }] },
            style: { italic: true, color: '0000FF' },
          },
        ],
      })],
    });

    expect(batch.commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: {
          range: '0:5',
          bold: 'true',
          color: 'FF0000',
        },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[1]',
        props: {
          range: '6:11',
          italic: 'true',
          color: '0000FF',
        },
      },
    ]);
  });

  it('translates paragraph text+style edits to set on /p[N]', () => {
    const first = shape('7', 'Title');
    const second = shape('8', 'Body');
    first.textBody = {
      ...first.textBody!,
      paragraphs: [
        first.textBody!.paragraphs[0],
        second.textBody!.paragraphs[0],
      ],
    };
    const presentation = deck([first]);
    const ref = createElementRef(presentation.slides[0], first, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'paragraph-text-edits',
      mutations: [new UpdateTextMutation({
        target: ref,
        edits: [
          {
            scope: { kind: 'paragraph', paragraphIndex: 0 },
            text: '新标题',
            style: { bold: true, fontSize: 24 },
          },
          {
            scope: { kind: 'paragraph', paragraphIndex: 1 },
            text: '新正文',
          },
        ],
      })],
    });

    expect(batch.commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[1]',
        props: {
          text: '新标题',
          bold: 'true',
          size: '24pt',
        },
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]/p[2]',
        props: {
          text: '新正文',
        },
      },
    ]);
  });

  it('preserves mutation order in a native OfficeCLI batch command array', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const command: Command = {
      id: 'compound-1',
      mutations: [
        new UpdateTextMutation({ target: ref, value: 'after' }),
        new RemoveElementMutation({ target: ref }),
      ],
    };

    expect(toOfficeCliBatch(presentation, command).commands).toEqual([
      {
        command: 'set',
        path: '/slide[1]/shape[@id=7]',
        props: { text: 'after' },
      },
      {
        command: 'remove',
        path: '/slide[1]/shape[@id=7]',
      },
    ]);
  });

  it('advances presentation between mutations so consecutive Adds see updated indexes', () => {
    const existing = shape('7', 'kept');
    const presentation = deck([existing]);
    const first = shape('8', 'first-add');
    const second = shape('9', 'second-add');

    const batch = toOfficeCliBatch(presentation, {
      id: 'multi-add-1',
      mutations: [
        new AddElementMutation({
          target: {
            origin: 'slide',
            slideId: 'ppt/slides/slide1.xml',
            elementId: '8',
          },
          element: first,
          presentationElementIndex: 1,
        }),
        new AddElementMutation({
          target: {
            origin: 'slide',
            slideId: 'ppt/slides/slide1.xml',
            elementId: '9',
          },
          element: second,
          presentationElementIndex: 2,
        }),
      ],
    });

    expect(batch.commands).toEqual([
      {
        command: 'add',
        parent: '/slide[1]',
        type: 'shape',
        props: expect.objectContaining({ id: '8', zorder: '2' }),
      },
      {
        command: 'add',
        parent: '/slide[1]',
        type: 'shape',
        props: expect.objectContaining({ id: '9', zorder: '3' }),
      },
    ]);
  });

  it('advances presentation for Remove→Add so the restore uses post-remove indexes', () => {
    const keep = shape('7', 'keep');
    const victim = shape('8', 'victim');
    const presentation = deck([keep, victim]);
    const victimRef = createElementRef(presentation.slides[0], victim, 1);
    const remove = new RemoveElementMutation({ target: victimRef });
    const restore = remove.inverse(presentation);
    if (!restore) throw new TypeError('A removed shape must remain restorable');

    const batch = toOfficeCliBatch(presentation, {
      id: 'remove-add-1',
      mutations: [remove, restore],
    });

    expect(batch.commands).toEqual([
      {
        command: 'remove',
        path: '/slide[1]/shape[@id=8]',
      },
      {
        command: 'add',
        parent: '/slide[1]',
        type: 'shape',
        props: expect.objectContaining({
          id: '8',
          zorder: '2',
          text: 'victim',
        }),
      },
    ]);
  });

  it('advances presentation when translating a compound undo (Add then text restore)', () => {
    const keep = shape('7', 'keep');
    const victim = plainShape('8', 'before');
    const presentation = deck([keep, victim]);
    const victimRef = createElementRef(presentation.slides[0], victim, 1);
    const forward: Command = {
      id: 'edit-1',
      mutations: [
        new UpdateTextMutation({ target: victimRef, value: 'after' }),
        new RemoveElementMutation({ target: victimRef }),
      ],
    };
    const entry = createUndoRedoEntry(presentation, forward);
    expect(entry).toBeDefined();
    const afterForward = applyCommand(presentation, forward).presentation;

    const batch = toOfficeCliBatch(afterForward, {
      id: 'undo-1',
      mutations: entry!.inverseMutations,
    });

    expect(batch.commands).toEqual([
      {
        command: 'add',
        parent: '/slide[1]',
        type: 'shape',
        props: expect.objectContaining({
          id: '8',
          zorder: '2',
          text: 'after',
        }),
      },
      {
        command: 'set',
        path: '/slide[1]/shape[@id=8]',
        props: { text: 'before' },
      },
    ]);
  });

  it('uses presentation order rather than the slide part filename', () => {
    const target = shape('7', 'before');
    const firstSlide = deck([]).slides[0];
    const targetSlide = {
      ...deck([target]).slides[0],
      index: 1,
      slideNumber: 2,
      partName: 'ppt/slides/slide42.xml',
    };
    const presentation: Presentation = {
      ...deck([]),
      slides: [firstSlide, targetSlide],
    };
    const ref = createElementRef(targetSlide, target, 0);

    const batch = toOfficeCliBatch(presentation, {
      id: 'text-1',
      mutations: [new UpdateTextMutation({ target: ref, value: 'after' })],
    });

    const command = batch.commands[0];
    expect(command.command).toBe('set');
    if (command.command !== 'set') throw new TypeError('Expected an OfficeCLI set command');
    expect(command.path).toBe('/slide[2]/shape[@id=7]');
  });

  it('rejects positional element references that cannot form a stable OfficeCLI path', () => {
    const target = shape(undefined, 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'text-1',
      mutations: [new UpdateTextMutation({ target: ref, value: 'after' })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'target.unstableElementId',
      commandId: 'text-1',
      mutationIndex: 0,
    }));
  });

  it.each([
    ['picture', picture('7'), '/slide[1]/picture[@id=7]'],
    ['table', table('10'), '/slide[1]/table[@id=10]'],
    ['chart', chart('11'), '/slide[1]/chart[@id=11]'],
  ] as const)(
    'translates removal of a %s from its frontend element type',
    (_name, target, path) => {
      const presentation = deck([target]);
      const ref = createElementRef(presentation.slides[0], target, 0);

      expect(toOfficeCliBatch(presentation, {
        id: 'remove-1',
        mutations: [new RemoveElementMutation({ target: ref })],
      }).commands).toEqual([{ command: 'remove', path }]);
    },
  );

  it('rejects removal of media; OfficeCLI has no stable @id selector for video/audio', () => {
    const target = media('12');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'target.unsupportedElement',
    }));
  });

  it('rejects transforms that cannot round-trip as exact EMUs', () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);

    expect(() => toOfficeCliBatch(presentation, {
      id: 'transform-1',
      mutations: [new UpdateShapeMutation({
        target: ref,
        value: {
          x: 0.5,
          y: 0,
          width: 10,
          height: 10,
          rotation: 0,
          flipH: false,
          flipV: false,
        },
      })],
    })).toThrowError(expect.objectContaining<Partial<OfficeCliTranslatorError>>({
      code: 'value.invalidTransform',
    }));
  });
});

function frame(id: string) {
  return {
    id,
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    rotation: 0,
    flipH: false,
    flipV: false,
  } as const;
}

function picture(id: string): PictureElement {
  return {
    type: 'picture',
    ...frame(id),
    imagePath: 'ppt/media/image1.png',
    mimeType: 'image/png',
    stroke: null,
  };
}

function table(id: string): TableElement {
  return {
    type: 'table',
    ...frame(id),
    cols: [],
    rows: [],
  };
}

function chart(id: string): ChartElement {
  return {
    type: 'chart',
    ...frame(id),
    chart: {} as ChartElement['chart'],
  };
}

function media(id: string): MediaElement {
  return {
    type: 'media',
    ...frame(id),
    mediaKind: 'video',
    posterPath: '',
    posterMimeType: '',
    mediaPath: 'ppt/media/video1.mp4',
    mimeType: 'video/mp4',
  };
}
