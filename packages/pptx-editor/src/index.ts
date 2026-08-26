export { MUTATION_TYPES } from './domain/mutation-types';
export type { MutationType } from './domain/mutation-types';
export { ELEMENT_ORIGINS } from './domain/element-origin';
export type { ElementOrigin } from './domain/element-origin';

export { Mutation } from './domain/mutation';
export type {
  ElementMutation,
  ElementRef,
  MutationCommandContext,
  MutationTarget,
  SlideRef,
} from './domain/mutation';

export { InsertSlideMutation } from './mutations/insert-slide';
export type { InsertSlideMutationParams } from './mutations/insert-slide';
export { RemoveSlideMutation } from './mutations/remove-slide';
export type { RemoveSlideMutationParams } from './mutations/remove-slide';

export { AddElementMutation } from './mutations/add-element';
export type {
  AddElementMutationParams,
} from './mutations/add-element';
export { RemoveElementMutation } from './mutations/remove-element';
export type {
  RemoveElementMutationParams,
} from './mutations/remove-element';
export {
  formatOfficeCliRange,
  paragraphRunPlainText,
  runPlainText,
  UpdateTextMutation,
} from './mutations/update-text';
export type {
  TextScope,
  TextSpan,
  TextStyleEdit,
  TextStylePatch,
  UpdateTextMutationParams,
} from './mutations/update-text';
export { UpdateShapeMutation } from './mutations/update-shape';
export type {
  ShapePatch,
  UpdateShapeMutationParams,
} from './mutations/update-shape';

export type { Command, NonEmptyReadonlyArray } from './domain/command';

export {
  POSITIONAL_ELEMENT_ID_PREFIX,
  createElementRef,
  deriveSlideTreeIndex,
  getElementMutationId,
  getSlideMutationId,
  isSlideRegionInsertIndex,
} from './adapters/pptx-json-adapter';
export type { ResolvedElementRef } from './adapters/pptx-json-adapter';

export {
  CommandExecutionError,
  MutationExecutionError,
  applyCommand,
  applyMutation,
} from './engine/mutation-engine';
export type {
  CommandExecutionResult,
  MutationExecutionErrorCode,
  MutationExecutionResult,
} from './engine/mutation-engine';

export { PptxEditorStore } from './store/editor-store';
export { EditorStoreError } from './store/errors';
export type { EditorStoreErrorCode } from './store/errors';
export { EDITOR_STORE_CHANGE_REASONS } from './store/types';
export type {
  EditorStoreChange,
  EditorStoreChangeReason,
  EditorStoreListener,
  EditorStoreSnapshot,
} from './store/types';
export {
  EDITOR_SYNC_STATUSES,
  READY_EDITOR_SYNC_STATE,
} from './store/sync-state';
export type {
  EditorSyncState,
  HaltedEditorSyncState,
  ReadyEditorSyncState,
} from './store/sync-state';

export {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
  OFFICECLI_VERSION,
} from './transport/officecli/constants';
export { OfficeCliTranslatorError } from './transport/officecli/errors';
export type {
  OfficeCliTranslatorErrorCode,
} from './transport/officecli/errors';
export { toOfficeCliBatch } from './transport/officecli/officecli-translator';
export type {
  OfficeCliAddCommand,
  OfficeCliAddShapeCommand,
  OfficeCliAddSlideCommand,
  OfficeCliBatch,
  OfficeCliCommand,
  OfficeCliCommandType,
  OfficeCliProps,
  OfficeCliRemoveCommand,
  OfficeCliSetCommand,
} from './transport/officecli/types';

export { UNDO_REDO_DIRECTIONS } from './history/constants';
export { UndoRedoStackError } from './history/errors';
export type { UndoRedoStackErrorCode } from './history/errors';
export { UndoRedoStack } from './history/undo-redo-stack';
export type {
  UndoRedoCommandIdContext,
  UndoRedoCommandIdFactory,
  UndoRedoDirection,
  UndoRedoStackListener,
  UndoRedoStackSnapshot,
} from './history/types';

export {
  COMMAND_SUBMISSION_STATUSES,
  OFFICECLI_BATCH_SEND_STATUSES,
} from './submission/constants';
export { CommandSubmitterError } from './submission/errors';
export type {
  CommandSubmitterErrorCode,
} from './submission/errors';
export { SerialOfficeCliSubmitter } from './submission/serial-officecli-submitter';
export type {
  CommandSubmission,
  CommandSubmissionResult,
  CommandSubmissionStatus,
  ConfirmedOfficeCliBatchSendResult,
  ConfirmedCommandSubmissionResult,
  HaltedCommandSubmissionResult,
  InvalidatedCommandSubmissionResult,
  OfficeCliBatchSendResult,
  OfficeCliBatchSender,
  RejectedOfficeCliBatchSendResult,
  RejectedCommandSubmissionResult,
  UnknownOfficeCliBatchSendResult,
} from './submission/types';

export { EDITOR_SESSION_CHANGE_REASONS } from './session/constants';
export { PptxEditorSessionError } from './session/errors';
export type { PptxEditorSessionErrorCode } from './session/errors';
export { PptxEditorSession } from './session/pptx-editor-session';
export type {
  PptxEditorSessionChange,
  PptxEditorSessionChangeReason,
  PptxEditorSessionListener,
  PptxEditorSessionListenerErrorHandler,
  PptxEditorSessionOptions,
  PptxEditorSessionSnapshot,
  PptxEditorSessionSubmission,
} from './session/types';

export { PptxEditorViewBindingError } from './rendering/errors';
export type {
  PptxEditorViewBindingErrorCode,
} from './rendering/errors';
export { PptxEditorViewBinding } from './rendering/pptx-editor-view-binding';
export { PptxEditorViewerHost } from './rendering/pptx-editor-viewer-host';
export type {
  PptxEditorBorrowedViewer,
  PptxEditorLoadedPresentation,
} from './rendering/pptx-editor-viewer-host';
export type {
  PptxEditorViewBindingOptions,
  PptxEditorViewErrorHandler,
  PptxEditorViewHost,
} from './rendering/types';

export { EDITOR_SELECTION_CHANGE_REASONS } from './interaction/constants';
export { PptxEditorSelectionControllerError } from './interaction/errors';
export type {
  PptxEditorSelectionControllerErrorCode,
} from './interaction/errors';
export {
  clientPointToSlidePoint,
  hitTestSlideElement,
  hitTestSlideShape,
  resolveElementSelection,
  resolveShapeSelection,
} from './interaction/hit-test';
export { PptxEditorSelectionController } from './interaction/pptx-editor-selection-controller';
export type {
  ClientPoint,
  ElementHitTestOptions,
  PptxEditorInteractionHost,
  PptxEditorElementSelection,
  PptxEditorSelectableElement,
  PptxEditorSelectionChange,
  PptxEditorSelectionChangeReason,
  PptxEditorSelectionControllerOptions,
  PptxEditorSelectionListener,
  PptxEditorSelectionListenerErrorHandler,
  PptxEditorSelectionSnapshot,
  PptxEditorShapeSelection,
  ShapeHitTestOptions,
  SlidePoint,
} from './interaction/types';
