/** Shared DOM policy for the built-in read-only comment margin.
 *
 * OOXML defines comment data and anchors but not this UI. Format packages own
 * anchor projection (DOCX text ranges, PPTX slide coordinates); this module
 * owns only the accessible card list. Cards stay in authored order and the
 * margin scrolls independently, avoiding text-length or line-count estimates.
 */

import type {
  ViewerCommentCard,
  ViewerCommentCardRenderContext,
  ViewerCommentCardRenderer,
} from '../comment-card.js';

export const READ_ONLY_COMMENT_MARGIN_WIDTH_PX = 280;

export type ReadOnlyCommentCard = ViewerCommentCard;
export type ReadOnlyCommentCardRenderContext = ViewerCommentCardRenderContext;
export type ReadOnlyCommentCardRenderer = ViewerCommentCardRenderer;

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

function displayDate(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const instant = new Date(value);
  if (!Number.isFinite(instant.getTime())) return value;
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(instant);
}

function appendCommentBody(host: HTMLElement, comment: ReadOnlyCommentCard, reply: boolean): void {
  const owner = host.ownerDocument;
  const block = createDiv(
    owner,
    `display:flex;gap:.65em;${reply ? 'margin:.78em 0 0;padding-top:.72em;border-top:.08em solid rgba(100,116,139,.2);' : ''}`,
  );
  block.dataset.ooxmlCommentPart = reply ? 'reply' : 'comment';
  const avatar = createDiv(
    owner,
    `display:grid;place-items:center;flex:0 0 auto;width:${reply ? '1.8em' : '2.3em'};height:${reply ? '1.8em' : '2.3em'};` +
      'border-radius:.7em;background:#2563eb;color:#fff;font:700 .72em/1 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;',
  );
  avatar.dataset.ooxmlCommentPart = 'avatar';
  avatar.textContent = (comment.author || 'C').trim().slice(0, 1).toUpperCase();
  const content = createDiv(owner, 'min-width:0;flex:1;');
  const identity = createDiv(owner, 'min-width:0;');
  const author = createDiv(
    owner,
    'min-width:0;font:700 .96em/1.3 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;' +
      'color:#0f172a;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;',
  );
  author.dataset.ooxmlCommentPart = 'author';
  author.textContent = comment.author || 'Comment';
  identity.appendChild(author);
  const formattedDate = displayDate(comment.date);
  if (formattedDate) {
    const date = createDiv(
      owner,
      'margin-top:.08em;font:500 .72em/1.35 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;' +
        'color:#64748b;white-space:nowrap;',
    );
    date.dataset.ooxmlCommentPart = 'date';
    date.textContent = formattedDate;
    date.setAttribute('title', comment.date as string);
    identity.appendChild(date);
  }
  const body = createDiv(
    owner,
    'margin-top:.62em;font:400 .92em/1.5 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;' +
      'color:#334155;white-space:pre-wrap;overflow-wrap:anywhere;',
  );
  body.dataset.ooxmlCommentPart = 'body';
  body.textContent = comment.text;
  content.append(identity, body);
  block.append(avatar, content);
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
        view: comment,
        active,
        zoom: Number(margin.dataset.ooxmlCommentZoom ?? '1'),
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
      'display:block;width:100%;box-sizing:border-box;margin:0 0 .62em;padding:.78em .92em;' +
      'border:0;border-radius:.62em;text-align:left;cursor:pointer;font:inherit;' +
      `background:${active ? 'var(--ooxml-comment-card-active-background,#dbeafe)' : 'var(--ooxml-comment-card-background,#f1f5f9)'};` +
      `box-shadow:${active ? '0 0 0 .15em rgba(37,99,235,.38)' : '0 .08em .16em rgba(15,23,42,.12)'};`;
    card.addEventListener('click', () => onActivate(active ? null : comment.id));
    appendCommentBody(card, comment, false);
    for (const reply of comment.replies ?? []) appendCommentBody(card, reply, true);
    item.appendChild(card);
    margin.appendChild(item);
  }
  if (cleanups.length > 0) cleanupByMargin.set(margin, Object.freeze(cleanups));
}
