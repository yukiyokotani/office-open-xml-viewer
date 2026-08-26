import { describe, expect, it, vi } from 'vitest';

import type { PictureElement, ShapeElement } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import type { ElementRef } from '../../src/domain/mutation';
import { UndoRedoStackError } from '../../src/history/errors';
import type { UndoRedoCommandIdFactory } from '../../src/history/types';
import { UndoRedoStack } from '../../src/history/undo-redo-stack';
import { RemoveElementMutation } from '../../src/mutations/remove-element';
import { RemoveSlideMutation } from '../../src/mutations/remove-slide';
import { UpdateShapeMutation } from '../../src/mutations/update-shape';
import { UpdateTextMutation } from '../../src/mutations/update-text';
import { PptxEditorStore } from '../../src/store/editor-store';
import {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from '../../src/submission/constants';
import { SerialOfficeCliSubmitter } from '../../src/submission/serial-officecli-submitter';
import type { OfficeCliBatchSendResult } from '../../src/submission/types';
import type { OfficeCliBatch } from '../../src/transport/officecli/types';
import { deck, plainShape, shape } from '../fixtures/presentation';

describe('UndoRedoStack', () => {
  it('records commands optimistically and submits undo and redo serially', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const history = createHistory(store, sendBatch);

    const edit = history.submit({
      id: 'edit-1',
      mutations: [
        new UpdateShapeMutation({
          target: ref,
          value: {
            x: 100,
            y: 0,
            width: 10,
            height: 10,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
        }),
        new UpdateTextMutation({ target: ref, value: 'after' }),
      ],
    });
    expect(textOf(store)).toBe('after');
    expect(shapeOf(store).x).toBe(100);
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      pendingSubmissions: 1,
      canUndo: true,
    });

    await edit.settled;
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      pendingSubmissions: 0,
      canUndo: true,
    });

    const undo = history.undo();
    expect(textOf(store)).toBe('before');
    expect(shapeOf(store).x).toBe(0);
    await undo.settled;
    expect(textOfBase(store)).toBe('before');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
      canRedo: true,
    });

    const redo = history.redo();
    expect(textOf(store)).toBe('after');
    expect(shapeOf(store).x).toBe(100);
    await redo.settled;
    expect(textOfBase(store)).toBe('after');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });
    expect(sendBatch.mock.calls.map(([batch]) => batch.commandId)).toEqual([
      'edit-1',
      'undo-1',
      'redo-2',
    ]);
  });

  it('redoes an optimistic undo before any submission settles', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => gate.promise)
      .mockResolvedValue(confirmedSendResult());
    const history = createHistory(store, sendBatch);

    const edit = history.submit(updateTextCommand('edit-1', ref, 'after'));
    const undo = history.undo();
    const redo = history.redo();

    expect(textOf(store)).toBe('after');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      pendingSubmissions: 3,
      canUndo: true,
    });
    expect(sendBatch).toHaveBeenCalledTimes(1);

    gate.resolve(confirmedSendResult());
    await Promise.all([edit.settled, undo.settled, redo.settled]);

    expect(textOfBase(store)).toBe('after');
    expect(sendBatch.mock.calls.map(([batch]) => batch.commandId)).toEqual([
      'edit-1',
      'undo-1',
      'redo-2',
    ]);
  });

  it('reverses inverse mutations so a compound command returns to its original state', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const history = createHistory(store, sendBatch);
    const command: Command = {
      id: 'compound-1',
      mutations: [
        new UpdateTextMutation({ target: ref, value: 'one' }),
        new UpdateTextMutation({ target: ref, value: 'two' }),
      ],
    };

    await history.submit(command).settled;
    await history.undo().settled;

    expect(textOfBase(store)).toBe('before');
    expect(sendBatch.mock.calls[1][0].commands).toEqual([
      expect.objectContaining({ props: { text: 'one' } }),
      expect.objectContaining({ props: { text: 'before' } }),
    ]);
  });

  it('undoes pending commands immediately and drops the optimistic tail after rejection', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const rejection = new Error('backend rejected command');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockImplementationOnce(() => gate.promise);
    const history = createHistory(store, sendBatch);
    const listener = vi.fn();
    store.subscribe(listener);
    const first = history.submit(updateTextCommand('edit-1', ref, 'one'));
    const second = history.submit(updateTextCommand('edit-2', ref, 'two'));
    const undo = history.undo();

    expect(textOf(store)).toBe('one');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 1,
      pendingSubmissions: 3,
      canUndo: true,
      canRedo: true,
    });
    gate.resolve({
      status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED,
      cause: rejection,
    });

    await expect(first.settled).resolves.toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.REJECTED,
    });
    await expect(second.settled).resolves.toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.INVALIDATED,
    });
    await expect(undo.settled).resolves.toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.INVALIDATED,
    });
    expect(textOf(store)).toBe('before');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 0,
      pendingSubmissions: 0,
      canUndo: false,
    });
    expect(listener).toHaveBeenCalledWith(expect.objectContaining({
      reason: 'command.rejected',
      commandId: 'edit-1',
      invalidatedCommandIds: ['edit-2', 'undo-1'],
    }));
    expect(sendBatch).toHaveBeenCalledTimes(1);
  });

  it('restores confirmed history when an optimistic undo is rejected', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const rejection = new Error('backend rejected undo');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce(confirmedSendResult())
      .mockResolvedValueOnce({
        status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED,
        cause: rejection,
      });
    const history = createHistory(store, sendBatch);

    await history.submit(updateTextCommand('edit-1', ref, 'after')).settled;
    const undo = history.undo();

    expect(textOf(store)).toBe('before');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
      canRedo: true,
    });
    await expect(undo.settled).resolves.toMatchObject({
      status: COMMAND_SUBMISSION_STATUSES.REJECTED,
    });
    expect(textOf(store)).toBe('after');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });
  });

  it('keeps an entry redoable when an optimistic redo is rejected', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce(confirmedSendResult())
      .mockResolvedValueOnce(confirmedSendResult())
      .mockResolvedValueOnce({
        status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED,
        cause: new Error('backend rejected redo'),
      });
    const history = createHistory(store, sendBatch);

    await history.submit(updateTextCommand('edit-1', ref, 'after')).settled;
    await history.undo().settled;
    const redo = history.redo();

    expect(textOf(store)).toBe('after');
    await redo.settled;
    expect(textOf(store)).toBe('before');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
      canRedo: true,
    });
  });

  it('temporarily blocks history while a non-undoable command is pending', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const gate = deferred<OfficeCliBatchSendResult>();
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce(confirmedSendResult())
      .mockImplementationOnce(() => gate.promise)
      .mockResolvedValueOnce(confirmedSendResult());
    const history = createHistory(store, sendBatch);

    await history.submit(updateTextCommand('edit-1', ref, 'after')).settled;
    const remove = history.submit({
      id: 'remove-slide-1',
      mutations: [new RemoveSlideMutation({
        target: { slideId: 'ppt/slides/slide1.xml' },
      })],
    });

    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: false,
      canRedo: false,
    });
    expect(() => history.undo()).toThrowError(
      expect.objectContaining<Partial<UndoRedoStackError>>({ code: 'undoRedo.busy' }),
    );

    gate.resolve({
      status: OFFICECLI_BATCH_SEND_STATUSES.REJECTED,
      cause: new Error('backend rejected removal'),
    });
    await remove.settled;

    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });

    await history.submit({
      id: 'remove-slide-2',
      mutations: [new RemoveSlideMutation({
        target: { slideId: 'ppt/slides/slide1.xml' },
      })],
    }).settled;
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 0,
      canUndo: false,
    });
  });

  it('restores a removed element at its original position through undo', async () => {
    const target = shape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const history = createHistory(store, sendBatch);

    await history.submit({
      id: 'remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    }).settled;

    expect(store.getSnapshot().basePresentation.slides[0].elements).toEqual([]);
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      canUndo: true,
    });

    await history.undo().settled;

    expect(textOfBase(store)).toBe('before');
    expect(sendBatch.mock.calls[1][0].commands).toEqual([{
      command: 'add',
      parent: '/slide[1]',
      type: 'shape',
      props: expect.objectContaining({
        id: '7',
        zorder: '1',
        text: 'before',
        x: '0emu',
        y: '0emu',
      }),
    }]);
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
      canRedo: true,
    });

    await history.redo().settled;
    expect(store.getSnapshot().basePresentation.slides[0].elements).toEqual([]);
  });

  it('treats removal of a non-restorable picture as a non-undoable history barrier', async () => {
    const target: PictureElement = {
      type: 'picture',
      id: '7',
      x: 0,
      y: 0,
      width: 10,
      height: 10,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath: 'ppt/media/image1.png',
      mimeType: 'image/png',
      stroke: null,
    };
    const presentation = deck([target]);
    const store = new PptxEditorStore(presentation);
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue(confirmedSendResult());
    const history = createHistory(store, sendBatch);

    await history.submit({
      id: 'remove-picture-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    }).settled;

    expect(store.getSnapshot().basePresentation.slides[0].elements).toEqual([]);
    expect(sendBatch.mock.calls[0][0].commands).toEqual([{
      command: 'remove',
      path: '/slide[1]/picture[@id=7]',
    }]);
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 0,
      canUndo: false,
    });
  });

  it('clears confirmed history after recovering a halted store from authoritative state', async () => {
    const target = plainShape('7', 'before');
    const store = new PptxEditorStore(deck([target]));
    const ref = createElementRef(store.getSnapshot().presentation.slides[0], target, 0);
    const unknownCause = new Error('request timed out');
    const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValueOnce(confirmedSendResult())
      .mockResolvedValueOnce({
        status: OFFICECLI_BATCH_SEND_STATUSES.UNKNOWN,
        cause: unknownCause,
      });
    const history = createHistory(store, sendBatch);

    await history.submit(updateTextCommand('edit-1', ref, 'confirmed')).settled;
    expect(history.getSnapshot().undoDepth).toBe(1);
    await expect(history.submit(updateTextCommand('edit-2', ref, 'unknown')).settled)
      .resolves.toMatchObject({ status: COMMAND_SUBMISSION_STATUSES.HALTED });
    expect(history.getSnapshot().canUndo).toBe(false);

    history.resync(deck([shape('7', 'server')]));

    expect(textOfBase(store)).toBe('server');
    expect(history.getSnapshot()).toMatchObject({
      undoDepth: 0,
      redoDepth: 0,
      pendingSubmissions: 0,
      canUndo: false,
      canRedo: false,
    });
  });
});

function createHistory(
  store: PptxEditorStore,
  sendBatch: (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>,
): UndoRedoStack {
  let nextId = 0;
  const createCommandId: UndoRedoCommandIdFactory = ({ direction }) => {
    nextId += 1;
    return `${direction}-${nextId}`;
  };
  return new UndoRedoStack(
    new SerialOfficeCliSubmitter(store, sendBatch),
    createCommandId,
  );
}

function updateTextCommand(id: string, target: ElementRef, value: string): Command {
  return {
    id,
    mutations: [new UpdateTextMutation({ target, value })],
  };
}

function textOf(store: PptxEditorStore): string | undefined {
  return shapeText(store.getSnapshot().presentation.slides[0].elements[0] as ShapeElement);
}

function textOfBase(store: PptxEditorStore): string | undefined {
  return shapeText(store.getSnapshot().basePresentation.slides[0].elements[0] as ShapeElement);
}

function shapeOf(store: PptxEditorStore): ShapeElement {
  return store.getSnapshot().presentation.slides[0].elements[0] as ShapeElement;
}

function shapeText(target: ShapeElement): string | undefined {
  const run = target.textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : undefined;
}

function confirmedSendResult(): OfficeCliBatchSendResult {
  return { status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED };
}

function deferred<T>(): {
  readonly promise: Promise<T>;
  readonly resolve: (value: T) => void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}
