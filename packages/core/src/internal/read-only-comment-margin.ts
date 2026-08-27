/** Shared DOM policy for the built-in read-only comment margin.
 *
 * OOXML defines comment data and anchors but not this UI. Format packages own
 * anchor projection (DOCX text ranges, PPTX slide coordinates); this module
 * owns the deliberately plain, themeable card list. Applications that need a
 * different structure use the format packages' comment and geometry APIs.
 */

export {
  READ_ONLY_COMMENT_MARGIN_WIDTH_PX,
  READ_ONLY_COMMENT_MARKER_SIZE_PX,
} from './read-only-comment-contract.js';
export type {
  ReadOnlyCommentMessage,
  ReadOnlyCommentThread,
} from './read-only-comment-contract.js';
import {
  READ_ONLY_COMMENT_MARKER_SIZE_PX,
  type ReadOnlyCommentMessage,
  type ReadOnlyCommentThread,
} from './read-only-comment-contract.js';

interface MountedCard {
  readonly item: HTMLDivElement;
  readonly card: HTMLButtonElement;
  readonly onClick: () => void;
  readonly onFocus: () => void;
  readonly onBlur: () => void;
  thread: ReadOnlyCommentThread;
  painted: boolean;
  focused: boolean;
  active: boolean;
  onSetActive: (id: string, active: boolean) => void;
  committedTop: number;
}

interface MarginState {
  readonly cards: Map<string, MountedCard>;
  readonly onScroll: () => void;
  zoom: number;
  resizeObserver?: ResizeObserver;
  onGeometryChange?: () => void;
  onScrollGeometryChange?: () => void;
}

let nextCommentCardDomId = 1;

const READ_ONLY_COMMENT_STYLES = `
:where(.ooxml-comment-marker) {
  padding: 0;
  border: 0;
  cursor: pointer;
  pointer-events: auto;
  background: transparent;
  box-shadow: none;
  color: var(--ooxml-comment-marker-color, var(--ooxml-comment-author-accent));
}
:where([data-ooxml-comment-highlight]) {
  position: absolute;
  pointer-events: none;
  background: var(--ooxml-comment-highlight, color-mix(in srgb, var(--ooxml-comment-author-accent) 18%, transparent));
  transform-origin: top left;
}
:where([data-ooxml-comment-highlight][data-active="true"]) {
  background: var(--ooxml-comment-highlight-active, color-mix(in srgb, var(--ooxml-comment-author-accent) 34%, transparent));
}
:where([data-ooxml-comment-target]) {
  position: absolute;
  box-sizing: border-box;
  pointer-events: none;
  border-style: solid;
  border-color: var(--ooxml-comment-target-border, #2563eb);
  background: var(--ooxml-comment-target-background, rgba(37, 99, 235, .06));
  transform-origin: center;
}
:where(.ooxml-comment-card) {
  position: relative;
  display: block;
  width: 100%;
  margin: 0 0 var(--ooxml-comment-card-gap, .42em);
  box-sizing: border-box;
  padding: var(--ooxml-comment-card-padding, .56em .68em);
  border: 0;
  border-radius: var(--ooxml-comment-card-radius, .3em);
  text-align: start;
  cursor: default;
  font: inherit;
  outline: none;
  background: var(--ooxml-comment-card-background, #fff);
  box-shadow: var(--ooxml-comment-card-shadow, none);
}
:where(button.ooxml-comment-card) {
  cursor: pointer;
}
:where(.ooxml-comment-card[data-standalone="true"]) {
  position: absolute;
  width: max-content;
  max-width: 100%;
  margin: 0;
  z-index: 3;
  pointer-events: none;
  overflow: hidden;
  font-size: 13px;
}
:where(.ooxml-comment-card[data-active="true"]) {
  background: var(--ooxml-comment-card-active-background, #eff6ff);
  box-shadow: var(--ooxml-comment-card-active-shadow, none);
}
:where(.ooxml-comment-card[data-focused="true"]) {
  box-shadow: var(--ooxml-comment-card-focus-shadow, 0 0 0 .12em rgba(37, 99, 235, .65));
}
:where(.ooxml-comment-card [data-ooxml-comment-part="content"]) {
  min-width: 0;
  flex: 1;
}
:where(.ooxml-comment-card [data-ooxml-comment-part="identity"]) {
  display: flex;
  align-items: baseline;
  gap: .48em;
  min-width: 0;
}
:where(.ooxml-comment-card__author) {
  min-width: 0;
  font: 700 .84em/1.3 var(--ooxml-comment-font-family, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif);
  color: var(--ooxml-comment-author-color, #0f172a);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
:where(.ooxml-comment-card__date) {
  font: 500 .66em/1.35 var(--ooxml-comment-date-font-family, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace);
  color: var(--ooxml-comment-muted-color, #64748b);
  white-space: nowrap;
}
:where(.ooxml-comment-card__body) {
  margin-top: .28em;
  font: 400 .84em/1.45 var(--ooxml-comment-font-family, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif);
  color: var(--ooxml-comment-body-color, #334155);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
:where(.ooxml-comment-card__reply) {
  margin: .55em 0 0 .45em;
  padding: .08em 0 0 .65em;
  border-left: .08em solid var(--ooxml-comment-reply-border, rgba(100, 116, 139, .24));
}
:where(.ooxml-comment-card [data-ooxml-comment-part="frame"]) {
  position: absolute;
  inset: 0;
  box-sizing: border-box;
  pointer-events: none;
  border-radius: inherit;
  border: var(--ooxml-comment-card-border, 1px solid rgba(148, 163, 184, .34));
  border-left: var(--ooxml-comment-card-border-left, .14em solid var(--ooxml-comment-author-accent));
  border-right: var(--ooxml-comment-card-border-right, var(--ooxml-comment-card-border, 1px solid rgba(148, 163, 184, .34)));
}
:where(.ooxml-comment-card[data-active="true"] [data-ooxml-comment-part="frame"]) {
  border: var(--ooxml-comment-card-active-border, 1px solid rgba(37, 99, 235, .5));
  border-left: var(--ooxml-comment-card-border-left, .14em solid var(--ooxml-comment-author-accent));
  border-right: var(--ooxml-comment-card-border-right, var(--ooxml-comment-card-active-border, 1px solid rgba(37, 99, 235, .5)));
}
:where([data-ooxml-comment-ui="margin"] > [data-ooxml-comment-item]) {
  position: absolute;
  left: 0;
  padding: 0 .14em;
  box-sizing: border-box;
  transform-origin: 0 0;
}
:where([data-ooxml-comment-ui="margin"] > [data-ooxml-comment-item]:last-child > .ooxml-comment-card) {
  margin-bottom: 0;
}
`;

export function ensureReadOnlyCommentStyles(owner: Document): void {
  const head = (owner as Document & { head?: HTMLHeadElement }).head;
  if (!head || head.querySelector('style[data-ooxml-comment-styles]')) return;
  const style = owner.createElement('style');
  style.dataset.ooxmlCommentStyles = '';
  style.textContent = READ_ONLY_COMMENT_STYLES;
  head.appendChild(style);
}

/** Stable built-in accent for one author. This is presentation policy rather
 * than OOXML data, so applications can replace it through the CSS variables. */
export function readOnlyCommentAuthorAccent(author: string | undefined): string {
  if (!author) return '#2563eb';
  let hash = 2166136261;
  for (const character of author.normalize('NFKC')) {
    hash ^= character.codePointAt(0) ?? 0;
    hash = Math.imul(hash, 16777619);
  }
  return `hsl(${(hash >>> 0) % 360} 68% 42%)`;
}

export interface ReadOnlyCommentMarkerOptions {
  readonly occurrenceKey: string;
  readonly visibleIndex: number;
  readonly author?: string;
  readonly active: boolean;
  readonly zoom: number;
  readonly onSetActive: (id: string, active: boolean) => void;
}

/** Create the shared, standalone comment glyph used by the DOCX/PPTX built-in UI. */
export function createReadOnlyCommentMarker(
  owner: Document,
  options: ReadOnlyCommentMarkerOptions,
): HTMLButtonElement {
  ensureReadOnlyCommentStyles(owner);
  const marker = owner.createElement('button');
  marker.type = 'button';
  marker.setAttribute('class', 'ooxml-comment-marker');
  marker.dataset.ooxmlCommentId = options.occurrenceKey;
  marker.dataset.ooxmlCommentMarker = '';
  marker.dataset.active = String(options.active);
  marker.setAttribute('aria-label', `Comment ${options.visibleIndex + 1}`);
  marker.setAttribute('aria-pressed', String(options.active));
  marker.style.cssText =
    `--ooxml-comment-author-accent:${readOnlyCommentAuthorAccent(options.author)};` +
    `position:absolute;transform:translate(-50%,-50%);width:${READ_ONLY_COMMENT_MARKER_SIZE_PX * options.zoom}px;` +
    `height:${READ_ONLY_COMMENT_MARKER_SIZE_PX * options.zoom}px;`;
  marker.innerHTML =
    '<svg viewBox="0 0 24 24" width="100%" height="100%" aria-hidden="true">' +
    '<path fill="currentColor" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M18 4C18.7956 4 19.5587 4.31607 20.1213 4.87868C20.6839 5.44129 21 6.20435 21 7V15C21 15.7956 20.6839 16.5587 20.1213 17.1213C19.5587 17.6839 18.7956 18 18 18H13L8 21V18H6C5.20435 18 4.44129 17.6839 3.87868 17.1213C3.31607 16.5587 3 15.7956 3 15V7C3 6.20435 3.31607 5.44129 3.87868 4.87868C4.44129 4.31607 5.20435 4 6 4H18Z"/>' +
    '</svg>';
  marker.addEventListener('click', () => options.onSetActive(
    options.occurrenceKey,
    !options.active,
  ));
  return marker;
}

export interface ReadOnlyCommentMarginOptions {
  readonly activeId: string | null;
  readonly zoom: number;
  /** Card-list width before Viewer zoom is applied. */
  readonly logicalWidth: number;
  readonly onSetActive: (id: string, active: boolean) => void;
  /** Called when a card's measured geometry changes. */
  readonly onGeometryChange?: () => void;
  /** Called when only the margin scroll offset changed. */
  readonly onScrollGeometryChange?: () => void;
  /** Preferred card top in the same CSS-pixel coordinate space as the margin. */
  readonly preferredTopById?: ReadonlyMap<string, number>;
}

export interface ReadOnlyCommentCardLayoutItem {
  readonly occurrenceKey: string;
  readonly preferredTop: number;
  readonly height: number;
}

const stateByMargin = new WeakMap<HTMLDivElement, MarginState>();

function destroyMountedCard(card: MountedCard, observer?: ResizeObserver): void {
  if (observer && typeof observer.unobserve === 'function') observer.unobserve(card.card);
  card.card.removeEventListener('click', card.onClick);
  card.card.removeEventListener('focus', card.onFocus);
  card.card.removeEventListener('blur', card.onBlur);
  card.item.remove();
}

/** Release built-in card resources before a virtualized margin is pooled. */
export function disposeReadOnlyCommentMargin(margin: HTMLDivElement): void {
  const state = stateByMargin.get(margin);
  stateByMargin.delete(margin);
  for (const card of state?.cards.values() ?? []) {
    destroyMountedCard(card, state?.resizeObserver);
  }
  state?.cards.clear();
  if (state) margin.removeEventListener('scroll', state.onScroll);
  state?.resizeObserver?.disconnect();
  margin.replaceChildren();
}

/** Scale the last committed card positions without measuring or rebuilding the
 * margin. Used only during a transient Viewer zoom; the settled render replaces
 * these preview positions with a newly measured layout. */
export function previewReadOnlyCommentMargin(
  margin: HTMLDivElement,
  ratio: number,
): boolean {
  const state = stateByMargin.get(margin);
  if (!state || state.cards.size === 0 || !Number.isFinite(ratio) || ratio <= 0) return false;
  const dataZoom = Number.parseFloat(margin.dataset.ooxmlCommentZoom ?? '');
  const zoom = Number.isFinite(dataZoom) && dataZoom > 0 ? dataZoom : state.zoom * ratio;
  for (const mounted of state.cards.values()) {
    mounted.item.style.top = `${mounted.committedTop * zoom}px`;
    mounted.item.style.transform = `scale(${zoom})`;
  }
  return true;
}

function createDiv(owner: Document, cssText = ''): HTMLDivElement {
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

function sameMessage(left: ReadOnlyCommentMessage, right: ReadOnlyCommentMessage): boolean {
  return left.messageKey === right.messageKey &&
    left.sourceId === right.sourceId &&
    left.author === right.author &&
    left.date === right.date &&
    left.text === right.text &&
    left.status === right.status;
}

function sameThread(left: ReadOnlyCommentThread, right: ReadOnlyCommentThread): boolean {
  return left.occurrenceKey === right.occurrenceKey &&
    sameMessage(left.root, right.root) &&
    left.replies.length === right.replies.length &&
    left.replies.every((reply, index) => sameMessage(reply, right.replies[index]));
}

/** Place cards near their anchors while preserving a non-overlapping reading
 * order. When the cards cannot fit in the page, fall back to a top-packed
 * scroll list so every thread remains reachable. */
export function layoutReadOnlyCommentCards(
  items: readonly ReadOnlyCommentCardLayoutItem[],
  availableHeight: number,
  gap: number,
): ReadonlyMap<string, number> {
  const ordered = items
    .map((item, index) => ({ ...item, index }))
    .sort((left, right) => left.preferredTop - right.preferredTop || left.index - right.index);
  const safeGap = Math.max(0, gap);
  const totalHeight = ordered.reduce((sum, item) => sum + Math.max(0, item.height), 0) +
    safeGap * Math.max(0, ordered.length - 1);
  if (availableHeight <= 0 || totalHeight > availableHeight) {
    let top = 0;
    return new Map(ordered.map((item) => {
      const entry = [item.occurrenceKey, top] as const;
      top += Math.max(0, item.height) + safeGap;
      return entry;
    }));
  }

  const tops = ordered.map((item) => Math.max(
    0,
    Math.min(item.preferredTop, availableHeight - item.height),
  ));
  for (let index = 1; index < ordered.length; index++) {
    const previousBottom = tops[index - 1] + ordered[index - 1].height + safeGap;
    tops[index] = Math.max(tops[index], previousBottom);
  }
  if (ordered.length > 0) {
    const last = ordered.length - 1;
    tops[last] = Math.min(tops[last], availableHeight - ordered[last].height);
    for (let index = last - 1; index >= 0; index--) {
      tops[index] = Math.min(
        tops[index],
        tops[index + 1] - safeGap - ordered[index].height,
      );
    }
    if (tops[0] < 0) {
      const shift = -tops[0];
      for (let index = 0; index < tops.length; index++) tops[index] += shift;
    }
  }
  return new Map(ordered.map((item, index) => [item.occurrenceKey, tops[index]]));
}

function appendCommentBody(host: HTMLElement, comment: ReadOnlyCommentMessage, reply: boolean): void {
  const owner = host.ownerDocument;
  const block = createDiv(owner);
  if (reply) block.setAttribute('class', 'ooxml-comment-card__reply');
  block.dataset.ooxmlCommentPart = reply ? 'reply' : 'comment';
  const content = createDiv(owner);
  content.dataset.ooxmlCommentPart = 'content';
  const identity = createDiv(owner);
  identity.dataset.ooxmlCommentPart = 'identity';
  const author = createDiv(owner);
  author.setAttribute('class', 'ooxml-comment-card__author');
  author.dataset.ooxmlCommentPart = 'author';
  author.textContent = comment.author || 'Comment';
  identity.appendChild(author);
  const formattedDate = displayDate(comment.date);
  if (formattedDate) {
    const date = createDiv(owner);
    date.setAttribute('class', 'ooxml-comment-card__date');
    date.dataset.ooxmlCommentPart = 'date';
    date.textContent = formattedDate;
    date.setAttribute('title', comment.date as string);
    identity.appendChild(date);
  }
  const body = createDiv(owner);
  body.setAttribute('class', 'ooxml-comment-card__body');
  body.dataset.ooxmlCommentPart = 'body';
  body.textContent = comment.text;
  content.appendChild(identity);
  content.appendChild(body);
  block.appendChild(content);
  host.appendChild(block);
}

export interface ReadOnlyCommentCardOptions {
  readonly active?: boolean;
  readonly focused?: boolean;
  /** Margin cards are buttons; anchored popups are intentionally non-interactive. */
  readonly interactive?: boolean;
  /** Remove list spacing and let an absolutely positioned popup size to its content. */
  readonly standalone?: boolean;
}

/** Paint the shared built-in comment-card structure into any format-owned host. */
export function paintReadOnlyCommentCard(
  card: HTMLElement,
  thread: ReadOnlyCommentThread,
  options: ReadOnlyCommentCardOptions = {},
): void {
  ensureReadOnlyCommentStyles(card.ownerDocument);
  const active = options.active ?? false;
  const focused = options.focused ?? false;
  const interactive = options.interactive ?? card.tagName === 'BUTTON';
  const accent = readOnlyCommentAuthorAccent(thread.root.author);
  card.setAttribute('class', 'ooxml-comment-card');
  card.dataset.ooxmlCommentId = thread.occurrenceKey;
  card.dataset.active = String(active);
  card.dataset.focused = String(focused);
  card.dataset.standalone = String(options.standalone ?? false);
  card.dataset.ooxmlCommentCard = '';
  if (interactive) card.setAttribute('aria-pressed', String(active));
  else card.removeAttribute('aria-pressed');
  card.style.cssText = `--ooxml-comment-author-accent:${accent};`;
  card.replaceChildren();
  appendCommentBody(card, thread.root, false);
  for (const reply of thread.replies) appendCommentBody(card, reply, true);
  // Keep decorative borders out of the card's box model. Browsers clamp very
  // thin CSS borders to one device-independent pixel, so a real card border
  // would consume a different share of the content width at each Viewer zoom
  // and could change line wrapping. The overlay preserves the familiar border
  // CSS API without letting it participate in text layout.
  const frame = createDiv(card.ownerDocument);
  frame.dataset.ooxmlCommentPart = 'frame';
  frame.setAttribute('aria-hidden', 'true');
  card.appendChild(frame);
}

/** Reconcile one built-in margin by occurrence key without replacing card nodes. */
export function buildReadOnlyCommentMargin(
  margin: HTMLDivElement,
  threads: readonly ReadOnlyCommentThread[],
  options: ReadOnlyCommentMarginOptions,
): ReadonlyMap<string, HTMLElement> {
  margin.setAttribute('role', 'list');
  margin.setAttribute('aria-label', 'Comments');
  margin.dataset.ooxmlCommentUi = 'margin';
  margin.dataset.ooxmlCommentZoom = String(options.zoom);
  let state = stateByMargin.get(margin);
  if (!state) {
    const created: MarginState = {
      cards: new Map<string, MountedCard>(),
      onScroll: () => created.onScrollGeometryChange?.(),
      zoom: options.zoom,
      onGeometryChange: options.onGeometryChange,
      onScrollGeometryChange: options.onScrollGeometryChange,
    };
    if (options.onGeometryChange || options.onScrollGeometryChange) {
      margin.addEventListener('scroll', created.onScroll, { passive: true });
    }
    if (options.onGeometryChange) {
      const ResizeObserverClass = margin.ownerDocument.defaultView?.ResizeObserver ??
        globalThis.ResizeObserver;
      if (ResizeObserverClass) {
        created.resizeObserver = new ResizeObserverClass(() => created.onGeometryChange?.());
      }
    }
    state = created;
    stateByMargin.set(margin, state);
  }
  state.zoom = options.zoom;
  state.onGeometryChange = options.onGeometryChange;
  state.onScrollGeometryChange = options.onScrollGeometryChange;

  const desired = new Set<string>();
  for (const thread of threads) {
    if (desired.has(thread.occurrenceKey)) {
      throw new Error(`Duplicate comment occurrence key: ${thread.occurrenceKey}`);
    }
    desired.add(thread.occurrenceKey);
  }

  for (const [id, mounted] of [...state.cards]) {
    if (!desired.has(id)) {
      state.cards.delete(id);
      destroyMountedCard(mounted, state.resizeObserver);
    }
  }

  for (const thread of threads) {
    let mounted = state.cards.get(thread.occurrenceKey);
    if (!mounted) {
      const item = createDiv(margin.ownerDocument);
      item.setAttribute('role', 'listitem');
      item.dataset.ooxmlCommentItem = '';
      const card = margin.ownerDocument.createElement('button');
      card.type = 'button';
      card.id = `ooxml-comment-card-${nextCommentCardDomId++}`;
      card.dataset.ooxmlCommentCard = '';
      const created: MountedCard = {
        item,
        card,
        thread,
        painted: false,
        focused: false,
        active: false,
        committedTop: 0,
        onSetActive: options.onSetActive,
        onClick: () => created.onSetActive(thread.occurrenceKey, !created.active),
        onFocus: () => {
          created.focused = true;
          created.card.dataset.focused = 'true';
        },
        onBlur: () => {
          created.focused = false;
          created.card.dataset.focused = 'false';
        },
      };
      card.addEventListener('click', created.onClick);
      card.addEventListener('focus', created.onFocus);
      card.addEventListener('blur', created.onBlur);
      item.appendChild(card);
      state.cards.set(thread.occurrenceKey, created);
      state.resizeObserver?.observe(card);
      mounted = created;
    }
    const active = options.activeId === thread.occurrenceKey;
    mounted.onSetActive = options.onSetActive;
    if (!mounted.painted || mounted.active !== active || !sameThread(mounted.thread, thread)) {
      paintReadOnlyCommentCard(mounted.card, thread, {
        active,
        focused: mounted.focused,
        interactive: true,
      });
    }
    mounted.painted = true;
    mounted.active = active;
    mounted.thread = thread;
  }

  const visualThreads = threads
    .map((thread, index) => ({ thread, index }))
    .sort((left, right) =>
      (options.preferredTopById?.get(left.thread.occurrenceKey) ?? 0) -
        (options.preferredTopById?.get(right.thread.occurrenceKey) ?? 0) ||
      left.index - right.index)
    .map(({ thread }) => thread);
  const orderedItems = visualThreads.flatMap((thread) => {
    const mounted = state.cards.get(thread.occurrenceKey);
    return mounted ? [mounted.item] : [];
  });
  const orderChanged = orderedItems.length !== margin.children.length ||
    orderedItems.some((item, index) => margin.children[index] !== item);
  if (orderChanged) margin.replaceChildren(...orderedItems);

  // Lay cards out once in a zoom-independent coordinate space, then scale the
  // finished boxes. Re-running browser text layout at each font size can move a
  // word across a line because of glyph and sub-pixel rounding. Temporarily
  // remove any preview transform so measurements stay in the canonical space.
  for (const mounted of state.cards.values()) {
    mounted.item.style.cssText =
      `width:${options.logicalWidth}px;top:${mounted.committedTop * options.zoom}px;`;
  }

  const firstCard = orderedItems[0]?.children[0] as HTMLElement | undefined;
  const computedGap = firstCard && margin.ownerDocument.defaultView?.getComputedStyle
    ? Number.parseFloat(margin.ownerDocument.defaultView.getComputedStyle(firstCard).marginBottom)
    : 0;
  const layout = layoutReadOnlyCommentCards(
    threads.map((thread) => {
      const mounted = state.cards.get(thread.occurrenceKey);
      return {
        occurrenceKey: thread.occurrenceKey,
        preferredTop: (options.preferredTopById?.get(thread.occurrenceKey) ?? 0) / options.zoom,
        height: mounted?.card.getBoundingClientRect().height ?? 0,
      };
    }),
    margin.clientHeight / options.zoom,
    Number.isFinite(computedGap) ? computedGap : 0,
  );
  for (const mounted of state.cards.values()) {
    mounted.committedTop = layout.get(mounted.thread.occurrenceKey) ?? 0;
    mounted.item.style.top = `${mounted.committedTop * options.zoom}px`;
    mounted.item.style.transform = `scale(${options.zoom})`;
  }

  return new Map(
    threads.flatMap((thread) => {
      const mounted = state.cards.get(thread.occurrenceKey);
      return mounted ? [[thread.occurrenceKey, mounted.card] as const] : [];
    }),
  );
}
