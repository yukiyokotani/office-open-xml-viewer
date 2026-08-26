import type { ElementRef } from '../domain/mutation';
import type { PptxEditorSessionChange } from '../session/types';
import { EDITOR_SELECTION_CHANGE_REASONS } from './constants';
import { PptxEditorSelectionControllerError } from './errors';
import {
  clientPointToSlidePoint,
  hitTestSlideElement,
  resolveElementSelection,
} from './hit-test';
import type {
  PptxEditorElementSelection,
  PptxEditorSelectionChange,
  PptxEditorSelectionControllerOptions,
  PptxEditorSelectionListener,
  PptxEditorSelectionListenerErrorHandler,
  PptxEditorSelectionSnapshot,
} from './types';

const DEFAULT_HIT_SLOP_PX = 4;

export class PptxEditorSelectionController {
  readonly #session: PptxEditorSelectionControllerOptions['session'];
  readonly #host: PptxEditorSelectionControllerOptions['host'];
  readonly #hitSlopPx: number;
  readonly #onListenerError: PptxEditorSelectionListenerErrorHandler;
  readonly #listeners = new Set<PptxEditorSelectionListener>();
  readonly #unsubscribeSession: () => void;
  #snapshot: PptxEditorSelectionSnapshot = Object.freeze({ selection: null });
  #unavailableSelectionTarget: ElementRef | undefined;
  #disposed = false;

  constructor(options: PptxEditorSelectionControllerOptions) {
    this.#session = options.session;
    this.#host = options.host;
    this.#hitSlopPx = normalizeHitSlopPx(options.hitSlopPx);
    this.#onListenerError = options.onListenerError ?? reportListenerError;
    this.#unsubscribeSession = this.#session.subscribe((change) => {
      this.#handleSessionChange(change);
    });
    this.#host.canvasElement.addEventListener('pointerdown', this.#handlePointerDown);
  }

  getSnapshot(): PptxEditorSelectionSnapshot {
    this.#assertActive();
    this.#clearSelectionFromAnotherSlide(this.#host.slideIndex);
    return this.#snapshot;
  }

  subscribe(listener: PptxEditorSelectionListener): () => void {
    this.#assertActive();
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  selectAtClientPoint(clientX: number, clientY: number): PptxEditorElementSelection | null {
    this.#assertActive();
    const slideIndex = this.#host.slideIndex;
    this.#clearSelectionFromAnotherSlide(slideIndex);
    const presentation = this.#session.getSnapshot().presentation;
    const point = clientPointToSlidePoint(
      this.#host.canvasElement,
      presentation,
      { clientX, clientY },
    );
    const selection = point
      ? hitTestSlideElement(
        presentation,
        slideIndex,
        point,
        { hitSlop: this.#hitSlopInSlideUnits(presentation) },
      )
      : undefined;
    if (selection) {
      this.#setSelection(selection, EDITOR_SELECTION_CHANGE_REASONS.SELECTED);
    } else {
      this.clear();
    }
    return this.#snapshot.selection;
  }

  #clearSelectionFromAnotherSlide(slideIndex: number): void {
    const current = this.#snapshot.selection;
    if (current && current.slideIndex !== slideIndex) this.clear();
  }

  select(target: ElementRef): PptxEditorElementSelection {
    this.#assertActive();
    const selection = resolveElementSelection(
      this.#session.getSnapshot().presentation,
      target,
    );
    if (!selection) {
      throw new PptxEditorSelectionControllerError(
        'selection.targetUnavailable',
        `Cannot select unavailable slide element ${target.slideId}/${target.elementId}`,
      );
    }
    this.#setSelection(selection, EDITOR_SELECTION_CHANGE_REASONS.SELECTED);
    return selection;
  }

  clear(): void {
    this.#assertActive();
    this.#unavailableSelectionTarget = undefined;
    if (!this.#snapshot.selection) return;
    this.#snapshot = Object.freeze({ selection: null });
    this.#publish(EDITOR_SELECTION_CHANGE_REASONS.CLEARED);
  }

  dispose(): void {
    if (this.#disposed) return;
    this.#disposed = true;
    this.#host.canvasElement.removeEventListener('pointerdown', this.#handlePointerDown);
    this.#unsubscribeSession();
    this.#listeners.clear();
    this.#snapshot = Object.freeze({ selection: null });
    this.#unavailableSelectionTarget = undefined;
  }

  readonly #handlePointerDown = (event: PointerEvent): void => {
    if (this.#disposed || event.button !== 0 || event.isPrimary === false) return;
    this.selectAtClientPoint(event.clientX, event.clientY);
  };

  #handleSessionChange(change: PptxEditorSessionChange): void {
    if (this.#disposed) return;
    const current = this.#snapshot.selection;
    if (!current) {
      if (!this.#unavailableSelectionTarget) return;
      const restored = resolveElementSelection(
        change.snapshot.presentation,
        this.#unavailableSelectionTarget,
      );
      if (restored) {
        this.#setSelection(restored, EDITOR_SELECTION_CHANGE_REASONS.UPDATED);
      } else if (!change.snapshot.isSubmitting) {
        this.#unavailableSelectionTarget = undefined;
      }
      return;
    }
    const next = resolveElementSelection(change.snapshot.presentation, current.target);
    if (!next) {
      this.#unavailableSelectionTarget = change.snapshot.isSubmitting
        ? current.target
        : undefined;
      this.#snapshot = Object.freeze({ selection: null });
      this.#publish(EDITOR_SELECTION_CHANGE_REASONS.CLEARED);
      return;
    }
    if (
      next.element === current.element
      && next.slideIndex === current.slideIndex
      && next.presentationElementIndex === current.presentationElementIndex
    ) {
      return;
    }
    this.#setSelection(next, EDITOR_SELECTION_CHANGE_REASONS.UPDATED);
  }

  #setSelection(
    selection: PptxEditorElementSelection,
    reason: PptxEditorSelectionChange['reason'],
  ): void {
    this.#unavailableSelectionTarget = undefined;
    const current = this.#snapshot.selection;
    if (
      current
      && current.target.origin === selection.target.origin
      && current.target.slideId === selection.target.slideId
      && current.target.elementId === selection.target.elementId
      && current.slideIndex === selection.slideIndex
      && current.element === selection.element
      && current.presentationElementIndex === selection.presentationElementIndex
    ) {
      return;
    }
    this.#snapshot = Object.freeze({ selection });
    this.#publish(reason);
  }

  #hitSlopInSlideUnits(
    presentation: ReturnType<PptxEditorSelectionControllerOptions['session']['getSnapshot']>['presentation'],
  ): number {
    const rect = this.#host.canvasElement.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return 0;
    return this.#hitSlopPx * Math.max(
      presentation.slideWidth / rect.width,
      presentation.slideHeight / rect.height,
    );
  }

  #publish(reason: PptxEditorSelectionChange['reason']): void {
    const change = Object.freeze({ reason, snapshot: this.#snapshot });
    for (const listener of [...this.#listeners]) {
      try {
        listener(change);
      } catch (cause) {
        try {
          this.#onListenerError(cause, change);
        } catch (reportingCause) {
          reportListenerError(new AggregateError(
            [cause, reportingCause],
            'PPTX editor selection listener and listener-error handler both failed',
          ));
        }
      }
    }
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new PptxEditorSelectionControllerError(
        'selection.disposed',
        'Cannot use a disposed PPTX editor selection controller',
      );
    }
  }
}

function normalizeHitSlopPx(value: number | undefined): number {
  if (value === undefined) return DEFAULT_HIT_SLOP_PX;
  return Number.isFinite(value) && value >= 0 ? value : DEFAULT_HIT_SLOP_PX;
}

function reportListenerError(cause: unknown): void {
  console.error('PPTX editor selection listener failed', cause);
}
