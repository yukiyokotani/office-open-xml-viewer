/** Shared DOM policy for the built-in read-only comment margin.
 *
 * OOXML defines comment data and anchors but not this UI. Format packages own
 * anchor projection (DOCX text ranges, PPTX slide coordinates); this module
 * owns only the accessible card list. Cards stay in authored order and the
 * margin scrolls independently, avoiding text-length or line-count estimates.
 */

export const READ_ONLY_COMMENT_MARGIN_WIDTH_PX = 280;

export interface ReadOnlyCommentCard {
  readonly id: string;
  readonly author?: string;
  readonly date?: string;
  readonly text: string;
  readonly replies?: readonly ReadOnlyCommentCard[];
}

export interface ReadOnlyCommentCardRenderContext {
  readonly comment: ReadOnlyCommentCard;
  readonly active: boolean;
  readonly activate: () => void;
}

export type ReadOnlyCommentCardRenderer = (
  host: HTMLElement,
  context: ReadOnlyCommentCardRenderContext,
) => void | (() => void);

const cleanupByMargin = new WeakMap<HTMLDivElement, readonly (() => void)[]>();

/** Release consumer-owned card resources before a virtualized margin is pooled. */
export function disposeReadOnlyCommentMargin(margin: HTMLDivElement): void {
  for (const cleanup of cleanupByMargin.get(margin) ?? []) cleanup();
  cleanupByMargin.delete(margin);
  margin.replaceChildren();
}

function createDiv(owner: Document, cssText: string): HTMLDivElement {
  const element = owner.createElement('div');
  element.style.cssText = cssText;
  return element;
}

function appendCommentBody(host: HTMLElement, comment: ReadOnlyCommentCard, reply: boolean): void {
  const owner = host.ownerDocument;
  const block = createDiv(
    owner,
    `${reply ? 'margin:8px 0 0 12px;padding-left:10px;border-left:2px solid rgba(37,99,235,.18);' : ''}`,
  );
  const author = createDiv(
    owner,
    'font:600 12px/1.35 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;' +
      'color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;',
  );
  author.textContent = [comment.author, comment.date].filter(Boolean).join(' · ') || 'Comment';
  const body = createDiv(
    owner,
    'margin-top:4px;font:13px/1.45 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;' +
      'color:#0f172a;white-space:pre-wrap;overflow-wrap:anywhere;',
  );
  body.textContent = comment.text;
  block.append(author, body);
  host.appendChild(block);
}

/** Rebuild one read-only margin. The returned buttons use normal document flow,
 * so their heights come from the browser and arbitrary comment text never needs
 * an estimated line count. */
export function buildReadOnlyCommentMargin(
  margin: HTMLDivElement,
  comments: readonly ReadOnlyCommentCard[],
  activeId: string | null,
  onActivate: (id: string | null) => void,
  renderCard?: ReadOnlyCommentCardRenderer,
): void {
  disposeReadOnlyCommentMargin(margin);
  margin.setAttribute('role', 'list');
  if (comments.length === 0) return;

  const cleanups: (() => void)[] = [];
  for (const comment of comments) {
    const active = activeId === comment.id;
    const item = createDiv(margin.ownerDocument, 'margin:0;padding:0;');
    item.setAttribute('role', 'listitem');
    if (renderCard) {
      const host = createDiv(margin.ownerDocument, 'display:block;width:100%;box-sizing:border-box;');
      host.dataset.ooxmlCommentId = comment.id;
      const cleanup = renderCard(host, {
        comment,
        active,
        activate: () => onActivate(active ? null : comment.id),
      });
      if (typeof cleanup === 'function') cleanups.push(cleanup);
      item.appendChild(host);
      margin.appendChild(item);
      continue;
    }
    const card = margin.ownerDocument.createElement('button');
    card.type = 'button';
    card.dataset.ooxmlCommentId = comment.id;
    card.setAttribute('aria-pressed', String(active));
    card.style.cssText =
      'display:block;width:100%;box-sizing:border-box;margin:0 0 8px;padding:10px 12px;' +
      'border:0;border-radius:8px;text-align:left;cursor:pointer;' +
      `background:${active ? 'var(--ooxml-comment-card-active-background,#dbeafe)' : 'var(--ooxml-comment-card-background,#f1f5f9)'};` +
      `box-shadow:${active ? '0 0 0 2px rgba(37,99,235,.38)' : '0 1px 2px rgba(15,23,42,.12)'};`;
    card.addEventListener('click', () => onActivate(active ? null : comment.id));
    appendCommentBody(card, comment, false);
    for (const reply of comment.replies ?? []) appendCommentBody(card, reply, true);
    item.appendChild(card);
    margin.appendChild(item);
  }
  if (cleanups.length > 0) cleanupByMargin.set(margin, Object.freeze(cleanups));
}
