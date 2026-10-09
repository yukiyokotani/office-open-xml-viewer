import type { XlsxConditionalFormattingReport } from '../../cf-diagnostics.js';

export const CF_NOTICE_TEXT =
  'Some conditional formatting rules in the displayed area could not be evaluated.';

/**
 * Default, nonblocking DOM notice for the Viewer's committed frame (#1547).
 * Library policy: it is driven only by the report delivered to the render
 * call that produced the committed frame, never by the workbook-wide
 * last-report getter (another viewer of the same workbook may update that).
 * It is a DOM status element, not Canvas paint, console output or onError.
 */
export class ConditionalFormattingNotice {
  readonly element: HTMLDivElement;

  constructor(ownerDocument: Document, parent: HTMLElement) {
    const element = ownerDocument.createElement('div');
    element.setAttribute('data-xlsx-cf-notice', '');
    element.setAttribute('role', 'status');
    element.setAttribute('aria-live', 'polite');
    element.hidden = true;
    // No `display` here: the `hidden` attribute must keep working.
    element.style.cssText =
      'position:absolute;right:24px;bottom:24px;z-index:3;' +
      'max-width:calc(100% - 48px);box-sizing:border-box;pointer-events:none;' +
      'padding:3px 8px;border-radius:4px;font:12px/1.4 sans-serif;' +
      'color:var(--ooxml-xlsx-chrome-text,#444);' +
      'background:var(--ooxml-xlsx-chrome-surface,#fff);' +
      'border:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);' +
      'box-shadow:0 1px 3px rgba(0,0,0,0.15);';
    parent.appendChild(element);
    this.element = element;
  }

  /** Apply the committed frame's report. A report without diagnostics, or no
   *  report, hides the notice. */
  update(report: XlsxConditionalFormattingReport | null): void {
    if (report && report.diagnostics.length > 0) {
      // Set text only when showing, so live regions announce the change.
      if (this.element.textContent !== CF_NOTICE_TEXT) this.element.textContent = CF_NOTICE_TEXT;
      this.element.hidden = false;
    } else {
      this.clear();
    }
  }

  clear(): void {
    this.element.hidden = true;
    this.element.textContent = '';
  }

  destroy(): void {
    this.clear();
    this.element.remove();
  }
}
