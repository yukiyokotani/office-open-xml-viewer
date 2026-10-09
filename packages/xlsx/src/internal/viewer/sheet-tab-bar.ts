import { HEADER_W } from '../../renderer.js';
import type { HiddenSheetMode } from '../../viewer.js';
import { ListenerScope } from './listener-scope.js';

/** Height of the composite viewer's footer (tab strip + zoom control). */
export const TAB_BAR_H = 30;
// Footer chrome stays in screen pixels: sheet zoom scales grid cells and their
// row/column headers, but must not resize the tab-navigation controls.
const TAB_NAV_W = HEADER_W;
// Gap between adjacent sheet tabs. The first tab also gets this much leading
// space so it is offset from the row-header boundary by the same margin that
// separates tabs from each other.
const TAB_GAP = 1;
/** `'dim'`-mode tab opacity: hidden/veryHidden tabs are greyed but selectable.
 *  A UI-presentation default (ECMA-376 defines no hidden-tab rendering); mirrors
 *  the named pptx `DEFAULT_HIDDEN_DIM` constant. */
const HIDDEN_TAB_DIM_OPACITY = 0.45;
/** Viewport margin kept around the sheet-list popover, in CSS px. */
const SHEET_LIST_EDGE = 8;
/** Widest the sheet list grows; longer names ellipsize (the full name is the title). */
const SHEET_LIST_MAX_W = 320;
/** Tooltip suffix on the scroll buttons where the sheet list is available. */
const SHEET_LIST_HINT = ' (right-click for sheet list)';

/** Workbook facts and navigation the tab strip needs from its engine. */
export interface SheetTabBarHost {
  hiddenSheetMode(): HiddenSheetMode;
  /** Whether sheet `index` is hidden/veryHidden (`<sheet state>`, §18.2.19). */
  isHidden(index: number): boolean;
  /** Activate sheet `index` (a tab click). */
  selectSheet(index: number): void;
}

/**
 * The composite viewer's footer: Excel-style tab-scroll buttons, the
 * scrollable sheet-tab strip and a slot for the zoom control. Owns its DOM and
 * every listener on it; the active sheet itself belongs to the engine.
 *
 * Right-clicking (or the keyboard context-menu gesture on) either scroll button
 * opens a list of the workbook's sheets. This is library UI policy, not an
 * OOXML rule. The list is a native `popover="auto"` element: the browser
 * supplies top-layer display, Escape and light dismissal, and focus return to
 * the scroll button. It stays a DOM descendant of the footer, so it inherits
 * the viewer's chrome CSS variables and belongs to the footer's document
 * realm. Its items exist only while it is open and are built from the names
 * already passed to {@link build}; nothing is parsed or rendered for it.
 * A pending right-press (waiting for its release, or queued to open after it)
 * is discarded by pointercancel, {@link build} and {@link destroy}, so a stale
 * gesture never opens a rebuilt list or moves focus.
 * Limitations: where the Popover API is unavailable the list is not created
 * and the browser's own context menu is left alone; the list is placed once
 * when opened and does not follow page scrolling/resizing while it is open.
 */
export class SheetTabBar {
  readonly tabBar: HTMLDivElement;
  readonly tabStrip: HTMLDivElement;
  /** Direction-aware flex row inside the LTR scroll host. Keeping direction on
   *  this inner row avoids browser-specific negative scrollLeft semantics. */
  readonly tabList: HTMLDivElement;
  readonly navPrev: HTMLButtonElement;
  readonly navNext: HTMLButtonElement;
  /** Popover listing every sheet; attached only where the Popover API exists. */
  readonly sheetList: HTMLDivElement;
  tabs: HTMLButtonElement[] = [];
  /** Per-tab colors parallel to `tabs`, from `<sheetPr><tabColor>`. */
  tabColors: (string | null)[] = [];
  private readonly ownerDocument: Document;
  private readonly navGroup: HTMLDivElement;
  private readonly popoverSupported: boolean;
  private readonly listeners = new ListenerScope();
  /** Listeners on the current tab buttons; replaced on every rebuild. */
  private tabListeners = new ListenerScope();
  private sheetNames: readonly string[] = [];
  private activeIndex = 0;
  /** Items of the open sheet list indexed by workbook sheet index, so a
   *  `'skip'`ped sheet is a hole and indices never shift. Empty while closed. */
  private listItems: HTMLButtonElement[] = [];
  /** Pointer last pressed on the scroll buttons, until its release/cancel. */
  private heldPointer: number | null = null;
  /** Pointer this bar explicitly captured on the nav group for a pending list. */
  private capturedPointer: number | null = null;
  /** Button whose list waits for {@link heldPointer}'s release. */
  private pendingOpener: HTMLButtonElement | null = null;
  /** Zero-delay open queued after a release (see {@link onNavPointerUp}). */
  private openTimer: number | undefined;

  constructor(ownerDocument: Document, private readonly host: SheetTabBarHost) {
    this.ownerDocument = ownerDocument;
    this.tabBar = ownerDocument.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:var(--ooxml-xlsx-chrome-background,#f0f0f0);` +
      `border-top:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);`;

    // Excel-style scroll buttons. They scroll the tab strip; they do NOT change
    // the active sheet. Disabled (greyed) at the ends / when there is no overflow.
    this.navPrev = this.makeNavButton('◀', 'Scroll tabs left', () => this.scrollTabs(-1));
    this.navNext = this.makeNavButton('▶', 'Scroll tabs right', () => this.scrollTabs(1));
    this.navPrev.dataset.xlsxTabNav = 'prev';
    this.navNext.dataset.xlsxTabNav = 'next';

    // Keep the two-button footer control at the row-header width from the 100%
    // view. It is viewer chrome, so workbook zoom must not resize or shift it.
    const navGroup = ownerDocument.createElement('div');
    navGroup.style.cssText =
      `display:flex;flex-shrink:0;width:${TAB_NAV_W}px;height:100%;`;
    navGroup.appendChild(this.navPrev);
    navGroup.appendChild(this.navNext);
    this.navGroup = navGroup;

    // The scrollable strip that actually holds the sheet tabs. position:relative
    // so each tab's offsetLeft is measured against the strip's scroll content.
    this.tabStrip = ownerDocument.createElement('div');
    // Keep the scroll host itself LTR so scrollLeft is consistently 0..max in
    // every browser. The inner tabList owns visual LTR/RTL ordering.
    this.tabStrip.style.cssText =
      `position:relative;display:block;flex:1;min-width:0;height:100%;` +
      `margin-left:${TAB_GAP}px;overflow-x:auto;overflow-y:hidden;scrollbar-width:none;`;
    this.tabStrip.classList.add('xlsx-tab-strip');
    this.listeners.on(this.tabStrip, 'scroll', () => this.updateNavButtons());

    // width:max-content preserves overflow scrolling; min-width:100% makes a
    // short RTL tab row fill the strip so row-reverse can right-align it.
    this.tabList = ownerDocument.createElement('div');
    this.tabList.style.cssText =
      `display:flex;align-items:flex-end;height:100%;` +
      `gap:${TAB_GAP}px;box-sizing:border-box;`;
    this.tabList.style.width = 'max-content';
    this.tabList.style.minWidth = '100%';
    this.tabStrip.appendChild(this.tabList);

    this.tabBar.appendChild(navGroup);
    this.tabBar.appendChild(this.tabStrip);

    this.sheetList = ownerDocument.createElement('div');
    this.popoverSupported = typeof this.sheetList.showPopover === 'function';
    if (this.popoverSupported) this.initSheetList();
  }

  /** Append trailing footer chrome (the zoom control). */
  append(element: HTMLElement): void {
    this.tabBar.appendChild(element);
  }

  /** Rebuild one tab per sheet, honoring the hidden-sheet mode. */
  build(sheetNames: readonly string[], tabColors: (string | null)[]): void {
    // A list opened (or about to open) for the previous names would map to the
    // wrong sheets, so drop both the open list and any pending right-press.
    this.cancelPendingOpen();
    this.closeSheetList();
    this.sheetNames = sheetNames;
    this.tabListeners.dispose();
    this.tabListeners = new ListenerScope();
    this.tabList.innerHTML = '';
    this.tabs = [];
    this.tabColors = tabColors;
    sheetNames.forEach((name, i) => {
      const btn = this.ownerDocument.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabCss(i, false);
      this.tabListeners.on(btn, 'click', () => this.host.selectSheet(i));
      this.tabList.appendChild(btn);
      this.tabs.push(btn);
    });
    this.updateNavButtons();
  }

  private makeNavButton(glyph: string, label: string, onClick: () => void): HTMLButtonElement {
    const btn = this.ownerDocument.createElement('button');
    btn.textContent = glyph;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.classList.add('xlsx-tab-nav');
    btn.style.cssText = this.navButtonStyle(false);
    this.listeners.on(btn, 'click', onClick);
    return btn;
  }

  private navButtonStyle(disabled: boolean): string {
    // Plain triangle icons — no border / tab chrome. The background (incl. the
    // hover tint) lives in the injected `.xlsx-tab-nav` stylesheet so the inline
    // style does not shadow the `:hover` rule.
    const base =
      `flex:1;height:100%;padding:0;` +
      `display:flex;align-items:center;justify-content:center;` +
      `border:none;color:var(--ooxml-xlsx-chrome-text-muted,#666);font-size:9px;line-height:1;` +
      `box-sizing:border-box;outline:none;`;
    // A disabled button ignores pointers, so its right-click reaches the nav
    // group, which owns the sheet-list listener (see onNavContextMenu).
    return disabled
      ? base + `opacity:0.3;cursor:default;pointer-events:none;`
      : base + `cursor:pointer;`;
  }

  scrollTabs(dir: -1 | 1): void {
    const strip = this.tabStrip;
    const viewLeft = strip.scrollLeft;
    const viewRight = viewLeft + strip.clientWidth;
    let target: number | null = null;
    if (dir === 1) {
      // Nearest tab clipped on the physical right; align its right edge.
      let nearestRight = Number.POSITIVE_INFINITY;
      for (const tab of this.tabs) {
        const right = tab.offsetLeft + tab.offsetWidth;
        if (right > viewRight + 1) nearestRight = Math.min(nearestRight, right);
      }
      if (Number.isFinite(nearestRight)) target = nearestRight - strip.clientWidth;
    } else {
      // Nearest tab clipped on the physical left; align its left edge. Search
      // by geometry, not DOM order, because RTL reverses the visual tab row.
      let nearestLeft = Number.NEGATIVE_INFINITY;
      for (const tab of this.tabs) {
        const left = tab.offsetLeft;
        if (left < viewLeft - 1) nearestLeft = Math.max(nearestLeft, left);
      }
      if (Number.isFinite(nearestLeft)) target = nearestLeft;
    }
    if (target !== null) {
      // Instant (not smooth) so the disabled state is consistent the moment the
      // click resolves — keeps the interaction deterministic to drive/test.
      strip.scrollLeft = Math.max(0, Math.min(target, strip.scrollWidth - strip.clientWidth));
    }
    this.updateNavButtons();
  }

  updateNavButtons(): void {
    const strip = this.tabStrip;
    const atStart = strip.scrollLeft <= 0;
    const atEnd = strip.scrollLeft + strip.clientWidth >= strip.scrollWidth - 1;
    // No overflow => scrollWidth ≈ clientWidth => both ends true => both disabled.
    this.navPrev.style.cssText = this.navButtonStyle(atStart);
    this.navNext.style.cssText = this.navButtonStyle(atEnd);
  }

  /** Restyle every tab for the active sheet and keep that tab in view. */
  setActive(index: number): void {
    const previous = this.activeIndex;
    this.activeIndex = index;
    // An open sheet list updates its two affected items in place.
    this.styleListItem(previous);
    this.styleListItem(index);
    this.tabs.forEach((btn, i) => {
      btn.style.cssText = this.tabCss(i, i === index);
    });
    // Keep the active tab visible by scrolling the tab strip HORIZONTALLY only.
    // `scrollIntoView` walks every scrollable ancestor, so it also scrolls the
    // page vertically — on first load that jumped the whole page down to the
    // tab bar (the active sheet is set during load). Adjust the strip's
    // scrollLeft directly so the page never moves.
    // `offsetParent === null` for a `display:none` tab (a hidden sheet reached
    // by an explicit goToSheet in 'skip' mode). Its getBoundingClientRect is all
    // zeros, which would spuriously scroll the strip — skip the scroll for it.
    const tab = this.tabs[index];
    if (tab && tab.offsetParent !== null) {
      const strip = this.tabStrip;
      const tabRect = tab.getBoundingClientRect();
      const stripRect = strip.getBoundingClientRect();
      if (tabRect.left < stripRect.left) {
        strip.scrollLeft -= stripRect.left - tabRect.left;
      } else if (tabRect.right > stripRect.right) {
        strip.scrollLeft += tabRect.right - stripRect.right;
      }
    }
    this.updateNavButtons();
  }

  /** Mirror the workbook footer around the sheet-tab strip for an RTL sheet.
   *  The DOM order remains navigation → tabs → zoom, which is also the
   *  logical reading order; `row-reverse` places that sequence right-to-left.
   *  Move the strip's leading gap with it so the spacing stays symmetric. */
  setDirection(rtl: boolean): void {
    this.tabBar.style.flexDirection = rtl ? 'row-reverse' : 'row';
    this.tabStrip.style.marginLeft = rtl ? '0' : `${TAB_GAP}px`;
    this.tabStrip.style.marginRight = rtl ? `${TAB_GAP}px` : '0';
    this.tabList.style.flexDirection = rtl ? 'row-reverse' : 'row';
  }

  private tabStyle(active: boolean, tabColor?: string | null): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;flex:none;padding:0 14px;position:relative;` +
      `border:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);border-bottom:none;` +
      `border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    // `<sheetPr><tabColor>` renders as a color bar along the tab's bottom edge
    // (Excel's "sheet tab color" treatment), drawn as an inset bottom shadow so
    // it doesn't fight the tab's own border/background. The active tab keeps a
    // thinner bar since its bottom merges into the white sheet body.
    const bar = tabColor
      ? `box-shadow:inset 0 -${active ? 2 : 3}px 0 0 ${tabColor};`
      : '';
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:var(--ooxml-xlsx-chrome-surface,#fff);` +
        `color:var(--ooxml-xlsx-chrome-text,#000);` +
        `border-bottom:1px solid var(--ooxml-xlsx-chrome-surface,#fff);` +
        `font-weight:600;top:1px;` +
        bar
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:var(--ooxml-xlsx-chrome-surface-muted,#e0e0e0);` +
        `color:var(--ooxml-xlsx-chrome-text-muted,#555);` +
        bar;
  }

  /**
   * Full inline style for the tab of sheet `i`, honoring the hidden-sheet mode:
   * `'skip'` hides the tab of a hidden/veryHidden sheet (`display:none`); `'dim'`
   * greys it but leaves it clickable; `'show'` styles every tab normally. Used
   * by both build and setActive so navigation never wipes the styling.
   */
  private tabCss(i: number, active: boolean): string {
    let css = this.tabStyle(active, this.tabColors[i]);
    const mode = this.host.hiddenSheetMode();
    if (mode !== 'show' && this.host.isHidden(i)) {
      css += mode === 'skip' ? 'display:none;' : `opacity:${HIDDEN_TAB_DIM_OPACITY};`;
    }
    return css;
  }

  private initSheetList(): void {
    const list = this.sheetList;
    list.setAttribute('popover', 'auto');
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', 'Sheets');
    list.dataset.xlsxSheetList = '';
    // No inline `display`: the UA popover rule hides the element while closed.
    // `inset:auto;margin:0` replace the UA centering; placeSheetList sets the
    // position, and the viewport-bounded max-height makes long lists scroll.
    list.style.cssText =
      `inset:auto;margin:0;padding:4px 0;box-sizing:border-box;min-width:${TAB_NAV_W}px;` +
      `overflow-x:hidden;overflow-y:auto;overscroll-behavior:contain;` +
      `background:var(--ooxml-xlsx-chrome-surface,#fff);color:var(--ooxml-xlsx-chrome-text,#000);` +
      `border:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);border-radius:4px;` +
      `box-shadow:0 4px 16px rgba(0,0,0,0.2);`;
    this.tabBar.appendChild(list);
    // Tooltip-only discoverability hint. The accessible name (aria-label) and
    // roles are unchanged: the list is a role=group, not an ARIA menu, so no
    // aria-haspopup is claimed.
    this.navPrev.title += SHEET_LIST_HINT;
    this.navNext.title += SHEET_LIST_HINT;
    this.listeners.on(this.navGroup, 'contextmenu', (event) => this.onNavContextMenu(event));
    this.listeners.on(this.navGroup, 'pointerdown', (event) => this.onNavPointerDown(event));
    this.listeners.on(this.navGroup, 'pointerup', (event) => this.onNavPointerUp(event));
    this.listeners.on(this.navGroup, 'pointercancel', (event) => this.onNavPointerCancel(event));
    this.listeners.on(list, 'click', (event) => this.onSheetListClick(event));
    this.listeners.on(list, 'keydown', (event) => this.onSheetListKeyDown(event));
    this.listeners.on(list, 'toggle', (event) => {
      // Items exist only while open, so navigation never restyles a closed list.
      if ((event as Event & { newState?: string }).newState !== 'closed') return;
      this.listItems = [];
      this.sheetList.replaceChildren();
    });
  }

  private onNavPointerDown(event: PointerEvent): void {
    // While a list waits on a captured press, another pointer is unrelated and
    // must not replace the pointer whose release opens it. Otherwise track the
    // newest press (this also recovers from a release that never reached us).
    if (this.pendingOpener) return;
    this.heldPointer = event.pointerId;
  }

  private onNavContextMenu(event: MouseEvent): void {
    if (this.sheetNames.length === 0) return;
    event.preventDefault();
    this.clearOpenTimer();
    const opener = this.navOpener(event);
    // Keyboard menus and platforms that fire contextmenu on release open now.
    if (event.buttons === 0 || this.heldPointer === null) {
      // The new gesture supersedes any held press as well as its queued open.
      this.cancelPendingOpen();
      this.openSheetList(opener);
      return;
    }
    // Fired while the button is still pressed: wait for the release, captured
    // so it reaches the nav group even if the pointer moved away.
    const pointerId = this.heldPointer;
    try {
      this.navGroup.setPointerCapture(pointerId);
    } catch {
      // The pointer is no longer active (its release was missed): open now.
      this.cancelPendingOpen();
      this.openSheetList(opener);
      return;
    }
    this.capturedPointer = pointerId;
    this.pendingOpener = opener;
  }

  /** The scroll button a context-menu gesture belongs to. A disabled button
   *  has `pointer-events:none`, so its pointer right-click targets the group
   *  and is resolved by position. */
  private navOpener(event: MouseEvent): HTMLButtonElement {
    if (event.target === this.navPrev) return this.navPrev;
    if (event.target === this.navNext) return this.navNext;
    const next = this.navNext.getBoundingClientRect();
    return event.clientX >= next.left && event.clientX < next.right ? this.navNext : this.navPrev;
  }

  /** Normal completion: the held pointer was released. */
  private onNavPointerUp(event: PointerEvent): void {
    if (event.pointerId !== this.heldPointer) return;
    this.heldPointer = null;
    // The browser releases capture itself after pointerup.
    this.capturedPointer = null;
    const opener = this.pendingOpener;
    if (!opener) return;
    this.pendingOpener = null;
    // Open after this pointerup finishes dispatching. HTML popover light
    // dismissal runs on pointerup; because the matching pointerdown happened
    // while no popover was open, it would hide a list opened during the press.
    // https://html.spec.whatwg.org/multipage/popover.html#popover-light-dismiss
    const view = this.ownerDocument.defaultView;
    if (!view) return;
    this.openTimer = view.setTimeout(() => {
      this.openTimer = undefined;
      this.openSheetList(opener);
    }, 0);
  }

  /** The held pointer's gesture was cancelled: discard, never open. */
  private onNavPointerCancel(event: PointerEvent): void {
    if (event.pointerId !== this.heldPointer) return;
    this.cancelPendingOpen();
  }

  private clearOpenTimer(): void {
    if (this.openTimer !== undefined) this.ownerDocument.defaultView?.clearTimeout(this.openTimer);
    this.openTimer = undefined;
  }

  /** Discard a right-press interaction in any stage: still held (pending
   *  release) or released with its open already queued. Shared by
   *  pointercancel, build and destroy. Releases only a capture this bar took
   *  and still owns; tab clicks and scroll-button clicks are unaffected. */
  private cancelPendingOpen(): void {
    this.clearOpenTimer();
    this.pendingOpener = null;
    this.heldPointer = null;
    const captured = this.capturedPointer;
    this.capturedPointer = null;
    if (captured !== null && this.navGroup.hasPointerCapture(captured)) {
      this.navGroup.releasePointerCapture(captured);
    }
  }

  private openSheetList(opener: HTMLButtonElement): void {
    const list = this.sheetList;
    if (!list.isConnected) return;
    this.renderSheetList();
    // showPopover records the focused element; native Escape and hidePopover
    // return focus to it, so focus the scroll button first.
    opener.focus({ preventScroll: true });
    if (!list.matches(':popover-open')) list.showPopover();
    this.placeSheetList(opener);
    const current = this.listItems[this.activeIndex] ?? this.listItems.find((item) => item !== undefined);
    if (current) this.focusListItem(current);
  }

  private renderSheetList(): void {
    const skip = this.host.hiddenSheetMode() === 'skip';
    this.listItems = [];
    this.sheetList.replaceChildren();
    this.sheetNames.forEach((name, i) => {
      if (skip && this.host.isHidden(i)) return;
      const item = this.ownerDocument.createElement('button');
      item.type = 'button';
      // Literal text: sheet names are arbitrary workbook strings.
      item.textContent = name;
      item.title = name;
      this.listItems[i] = item;
      this.styleListItem(i);
      this.sheetList.appendChild(item);
    });
  }

  /** Mark the active sheet and dim a hidden one (`'dim'` mode), matching the tabs. */
  private styleListItem(index: number): void {
    const item = this.listItems[index];
    if (!item) return;
    const active = index === this.activeIndex;
    let css =
      `display:block;width:100%;box-sizing:border-box;margin:0;padding:4px 16px 4px 12px;` +
      `border:none;text-align:start;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;` +
      `font-size:12px;color:inherit;cursor:pointer;`;
    css += active
      ? `font-weight:600;background:var(--ooxml-xlsx-chrome-surface-muted,#e0e0e0);`
      : `background:transparent;`;
    if (this.host.hiddenSheetMode() === 'dim' && this.host.isHidden(index)) {
      css += `opacity:${HIDDEN_TAB_DIM_OPACITY};`;
    }
    item.style.cssText = css;
    if (active) item.setAttribute('aria-current', 'true');
    else item.removeAttribute('aria-current');
  }

  /** Place the open list beside the scroll button on the side with more room,
   *  so it never covers the footer, and bound it to the viewport. */
  private placeSheetList(opener: HTMLElement): void {
    const list = this.sheetList;
    const viewport = this.ownerDocument.documentElement;
    const anchor = opener.getBoundingClientRect();
    const spaceAbove = anchor.top - SHEET_LIST_EDGE;
    const spaceBelow = viewport.clientHeight - anchor.bottom - SHEET_LIST_EDGE;
    const above = spaceAbove >= spaceBelow;
    list.style.maxHeight = `${Math.max(0, Math.floor(above ? spaceAbove : spaceBelow))}px`;
    list.style.maxWidth =
      `${Math.max(0, Math.min(SHEET_LIST_MAX_W, viewport.clientWidth - 2 * SHEET_LIST_EDGE))}px`;
    const { width, height } = list.getBoundingClientRect();
    list.style.top = `${above ? anchor.top - height : anchor.bottom}px`;
    list.style.left =
      `${Math.max(SHEET_LIST_EDGE, Math.min(anchor.left, viewport.clientWidth - width - SHEET_LIST_EDGE))}px`;
  }

  /** Focus without scrolling the page; scroll only the list to the item. */
  private focusListItem(item: HTMLButtonElement): void {
    item.focus({ preventScroll: true });
    const list = this.sheetList;
    const top = item.offsetTop;
    const bottom = top + item.offsetHeight;
    if (top < list.scrollTop) list.scrollTop = top;
    else if (bottom > list.scrollTop + list.clientHeight) list.scrollTop = bottom - list.clientHeight;
  }

  private onSheetListClick(event: MouseEvent): void {
    const index = this.listItems.indexOf(event.target as HTMLButtonElement);
    if (index < 0) return;
    this.sheetList.hidePopover();
    this.host.selectSheet(index);
  }

  /** Arrow keys move between items (wrapping); Tab, Enter and Space are native. */
  private onSheetListKeyDown(event: KeyboardEvent): void {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    event.preventDefault();
    const items = this.listItems.filter((item) => item !== undefined);
    if (items.length === 0) return;
    const current = items.indexOf(this.ownerDocument.activeElement as HTMLButtonElement);
    const step = event.key === 'ArrowDown' ? 1 : -1;
    const next = current < 0
      ? (step > 0 ? 0 : items.length - 1)
      : (current + step + items.length) % items.length;
    this.focusListItem(items[next]);
  }

  private closeSheetList(): void {
    if (this.popoverSupported && this.sheetList.isConnected && this.sheetList.matches(':popover-open')) {
      this.sheetList.hidePopover();
    }
  }

  /** Detach every tab-bar listener (the DOM leaves with the viewer subtree). */
  destroy(): void {
    this.cancelPendingOpen();
    this.closeSheetList();
    this.listItems = [];
    this.tabListeners.dispose();
    this.listeners.dispose();
  }
}
