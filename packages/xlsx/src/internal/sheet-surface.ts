/** DOM event adapter shared by the workbook and canvas-mounted sheet facades. */
export class CanvasSurface {
  private readonly cleanups: Array<() => void> = [];
  private readonly ownerDocument: Document;
  private readonly ownerWindow: Window | null;

  constructor(
    readonly canvas: HTMLCanvasElement,
    readonly area: HTMLDivElement,
    readonly input: HTMLDivElement,
  ) {
    this.ownerDocument = input.ownerDocument ?? document;
    this.ownerWindow = this.ownerDocument.defaultView;
  }

  on<K extends keyof HTMLElementEventMap>(
    type: K,
    listener: (event: HTMLElementEventMap[K]) => void,
    options?: AddEventListenerOptions | boolean,
  ): () => void {
    this.input.addEventListener(type, listener as EventListener, options);
    const cleanup = () => this.input.removeEventListener(type, listener as EventListener, options);
    this.cleanups.push(cleanup);
    return cleanup;
  }

  get viewportSize(): { width: number; height: number } {
    return { width: this.input.clientWidth, height: this.input.clientHeight };
  }

  get dpr(): number { return this.ownerWindow?.devicePixelRatio ?? 1; }

  localPoint(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.area.getBoundingClientRect();
    return { x: clientX - rect.left, y: clientY - rect.top };
  }

  sizeCanvas(canvas: HTMLCanvasElement, cssWidth: number, cssHeight: number): number {
    const dpr = this.dpr;
    canvas.width = Math.round(cssWidth * dpr);
    canvas.height = Math.round(cssHeight * dpr);
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    return dpr;
  }

  destroy(): void {
    for (const cleanup of this.cleanups.splice(0)) cleanup();
  }
}

export interface SheetOverlayHostOptions {
  readonly commentMaxWidth: number;
  readonly commentMaxHeight: number;
  readonly validationMaxWidth: number;
  readonly validationMaxHeight: number;
}

/** Format-owned overlay DOM. Geometry and state transitions remain in the
 *  XLSX engine, while node creation and stacking stay identical in both mounts. */
export class SheetOverlayHost {
  readonly selection: HTMLDivElement;
  readonly find: HTMLDivElement;
  readonly comment: HTMLDivElement;
  readonly commentStatus: HTMLDivElement;
  readonly validation: HTMLDivElement;

  constructor(
    area: HTMLDivElement,
    canvas: HTMLCanvasElement,
    input: HTMLDivElement,
    options: SheetOverlayHostOptions,
  ) {
    const ownerDocument = area.ownerDocument ?? document;
    this.selection = ownerDocument.createElement('div');
    this.selection.style.cssText =
      `position:absolute;top:0;left:0;z-index:1;pointer-events:none;overflow:hidden;width:100%;height:100%;`;

    this.find = ownerDocument.createElement('div');
    this.find.style.cssText =
      `position:absolute;top:0;left:0;z-index:1;pointer-events:none;overflow:hidden;width:100%;height:100%;`;

    this.comment = ownerDocument.createElement('div');
    this.comment.dataset.ooxmlCommentUi = 'popup';
    this.comment.setAttribute('role', 'note');
    this.comment.setAttribute('aria-hidden', 'true');
    this.comment.style.cssText =
      `position:absolute;z-index:3;pointer-events:none;display:none;` +
      `max-width:${options.commentMaxWidth}px;max-height:${options.commentMaxHeight}px;overflow:hidden;` +
      `font:13px/1.45 var(--ooxml-comment-font-family,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif);`;

    // Keep an empty live region mounted before any popup opens. Updating a
    // stable status node is announced reliably; adding/showing the visual note
    // and its contents in the same operation is not consistently announced.
    this.commentStatus = ownerDocument.createElement('div');
    this.commentStatus.setAttribute('role', 'status');
    this.commentStatus.setAttribute('aria-live', 'polite');
    this.commentStatus.setAttribute('aria-atomic', 'true');
    this.commentStatus.setAttribute('data-xlsx-comment-status', '');
    this.commentStatus.style.cssText =
      'position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;' +
      'clip:rect(0 0 0 0);white-space:nowrap;border:0;';

    this.validation = ownerDocument.createElement('div');
    this.validation.setAttribute('data-xlsx-validation-panel', '');
    this.validation.style.cssText =
      `position:absolute;z-index:4;pointer-events:auto;display:none;` +
      `min-width:80px;max-width:${options.validationMaxWidth}px;` +
      `max-height:${options.validationMaxHeight}px;overflow-y:auto;` +
      `box-sizing:border-box;background:#fff;border:1px solid #7f7f7f;` +
      `box-shadow:1px 2px 5px rgba(0,0,0,0.25);` +
      `font:12px/1.4 sans-serif;color:#222;`;
    this.validation.addEventListener('wheel', (event) => event.stopPropagation());

    // The host owns the stable stacking contract for both public facades.
    area.appendChild(canvas);
    area.appendChild(this.selection);
    area.appendChild(this.find);
    area.appendChild(input);
    area.appendChild(this.comment);
    area.appendChild(this.commentStatus);
    area.appendChild(this.validation);
  }

  clearSelection(): void { this.selection.textContent = ''; }
  appendSelection(element: HTMLElement): void { this.selection.appendChild(element); }
  clearFind(): void { this.find.textContent = ''; }
  appendFind(element: HTMLElement): void { this.find.appendChild(element); }

  hideComment(): void {
    this.comment.style.display = 'none';
    this.comment.setAttribute('aria-hidden', 'true');
    this.commentStatus.textContent = '';
  }
  announceComment(message: string): void { this.commentStatus.textContent = message; }
  showComment(left: number, top: number): void {
    this.comment.style.left = `${left}px`;
    this.comment.style.top = `${top}px`;
    this.comment.style.display = 'block';
    this.comment.setAttribute('aria-hidden', 'false');
  }

  hideValidation(): void { this.validation.style.display = 'none'; }
  showValidation(left: number, top: number): void {
    this.validation.style.left = `${left}px`;
    this.validation.style.top = `${top}px`;
    this.validation.style.display = 'block';
  }
}
