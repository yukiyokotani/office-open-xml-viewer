import { describe, expect, it, vi } from 'vitest';
import { SheetTabBar, type SheetTabBarHost } from './internal/viewer/sheet-tab-bar.js';
import { makeContainer, makeDocument, type FakeDocument, type FakeEl } from './viewer-destroy-test-dom.js';
import type { HiddenSheetMode } from './viewer.js';

type PopoverEl = FakeEl & { matches(selector: string): boolean };

/** Recording surfaces a browser provides on top of the fake DOM: Popover API,
 *  focus, pointer capture and a manually flushed zero-delay timer queue. */
function makePopoverDocument() {
  const doc = makeDocument();
  const focused = { element: null as FakeEl | null };
  const released: number[] = [];
  const timers = new Map<number, () => void>();
  let nextTimer = 0;
  const createElement = doc.createElement;
  doc.createElement = (tag: string) => {
    const el = createElement(tag);
    let open = false;
    const captured = new Set<number>();
    return Object.assign(el, {
      isConnected: true,
      focus: () => { focused.element = el; },
      showPopover: () => { open = true; },
      hidePopover: () => { open = false; },
      matches: (selector: string) => selector === ':popover-open' && open,
      setPointerCapture: (id: number) => { captured.add(id); },
      hasPointerCapture: (id: number) => captured.has(id),
      releasePointerCapture: (id: number) => { captured.delete(id); released.push(id); },
    });
  };
  Object.assign(doc, { documentElement: makeContainer(1000, 800) });
  Object.assign(doc.defaultView, {
    setTimeout: (fn: () => void) => { timers.set(++nextTimer, fn); return nextTimer; },
    clearTimeout: (id: number) => { timers.delete(id); },
  });
  const flushTimers = () => {
    const queued = [...timers.values()];
    timers.clear();
    for (const fn of queued) fn();
  };
  return { doc, focused, released, timers, flushTimers };
}

function makeBar(doc: FakeDocument, mode: HiddenSheetMode, hidden: number[]) {
  const selectSheet = vi.fn();
  const host: SheetTabBarHost = {
    hiddenSheetMode: () => mode,
    isHidden: (index) => hidden.includes(index),
    selectSheet,
  };
  const bar = new SheetTabBar(doc as unknown as Document, host);
  const navGroup = (bar.navPrev as unknown as FakeEl).parentElement as FakeEl;
  const tabBar = bar.tabBar as unknown as FakeEl;
  const list = () => tabBar.querySelector('[data-xlsx-sheet-list]') as PopoverEl | null;
  return { bar, selectSheet, navGroup, tabBar, list };
}

const contextMenu = (target: unknown, buttons?: number) =>
  ({ target, clientX: 0, buttons, preventDefault: vi.fn() });

/** Right-press on the prev button whose contextmenu fires while still held. */
function pressRightAndHold(navGroup: FakeEl, target: unknown, pointerId = 1): void {
  navGroup.dispatch('pointerdown', { pointerId, button: 2 });
  navGroup.dispatch('contextmenu', contextMenu(target, 2));
}

describe('SheetTabBar sheet list', () => {
  it('lists skip-mode sheets by workbook index as literal text and selects through the host', () => {
    const { doc, focused } = makePopoverDocument();
    const { bar, selectSheet, navGroup, list: getList } = makeBar(doc, 'skip', [1]);
    const unsafeName = '<img src=x onerror=alert(1)>';
    bar.build([unsafeName, 'Hidden', 'Third'], [null, null, null]);
    bar.setActive(2);
    const list = getList() as PopoverEl;

    // Tooltip hint only; the accessible name and popup semantics are unchanged.
    expect(bar.navPrev.title).toMatch(/right-click/i);
    expect(bar.navPrev.getAttribute('aria-label')).toBe('Scroll tabs left');
    expect(bar.navPrev.hasAttribute('aria-haspopup')).toBe(false);

    const event = contextMenu(bar.navPrev);
    navGroup.dispatch('contextmenu', event);

    expect(event.preventDefault).toHaveBeenCalledOnce();
    expect(list.matches(':popover-open')).toBe(true);
    expect(list.children.map((item) => item.textContent)).toEqual([unsafeName, 'Third']);
    const [first, third] = list.children;
    expect(third.getAttribute('aria-current')).toBe('true');
    expect(focused.element).toBe(third);

    // Navigation while open moves the marker on the existing items.
    bar.setActive(0);
    expect(list.children[0]).toBe(first);
    expect(first.getAttribute('aria-current')).toBe('true');
    expect(third.getAttribute('aria-current')).toBeNull();

    list.dispatch('click', { target: third });
    expect(selectSheet).toHaveBeenCalledWith(2);
    expect(list.matches(':popover-open')).toBe(false);
  });

  it('keeps the native context menu and adds no element without the Popover API', () => {
    const { bar, navGroup, list } = makeBar(makeDocument(), 'show', []);
    bar.build(['Only'], [null]);

    const event = contextMenu(bar.navNext);
    navGroup.dispatch('contextmenu', event);

    expect(event.preventDefault).not.toHaveBeenCalled();
    expect(list()).toBeNull();
    expect(bar.navNext.title).not.toMatch(/right-click/i);
  });

  it('opens a held right-press only after its own release has finished dispatching', () => {
    const { doc, focused, timers, flushTimers } = makePopoverDocument();
    const { bar, navGroup, list } = makeBar(doc, 'show', []);
    bar.build(['A', 'B'], [null, null]);

    pressRightAndHold(navGroup, bar.navPrev, 1);
    expect(list()!.matches(':popover-open')).toBe(false);

    // An unrelated pointer's release neither completes nor cancels the press.
    navGroup.dispatch('pointerup', { pointerId: 7 });
    navGroup.dispatch('pointercancel', { pointerId: 7 });
    expect(timers.size).toBe(0);

    navGroup.dispatch('pointerup', { pointerId: 1 });
    expect(list()!.matches(':popover-open')).toBe(false);
    expect(timers.size).toBe(1);

    flushTimers();
    expect(list()!.matches(':popover-open')).toBe(true);
    expect(focused.element).toBe(list()!.children[0]);
  });

  it('discards a held right-press on pointercancel without opening or focusing', () => {
    const { doc, focused, released, timers, flushTimers } = makePopoverDocument();
    const { bar, navGroup, list } = makeBar(doc, 'show', []);
    bar.build(['A', 'B'], [null, null]);

    pressRightAndHold(navGroup, bar.navPrev, 1);
    navGroup.dispatch('pointercancel', { pointerId: 1 });
    expect(released).toEqual([1]);
    expect(timers.size).toBe(0);

    // A late release for the cancelled pointer is no longer ours.
    navGroup.dispatch('pointerup', { pointerId: 1 });
    flushTimers();
    expect(list()!.matches(':popover-open')).toBe(false);
    expect(focused.element).toBeNull();
  });

  it('a keyboard context menu supersedes a held press without reopening after selection', () => {
    const { doc, flushTimers } = makePopoverDocument();
    const { bar, navGroup, list, selectSheet } = makeBar(doc, 'show', []);
    bar.build(['A', 'B'], [null, null]);
    pressRightAndHold(navGroup, bar.navPrev, 1);

    navGroup.dispatch('contextmenu', contextMenu(bar.navNext, 0));
    const popup = list() as PopoverEl;
    expect(popup.matches(':popover-open')).toBe(true);
    popup.dispatch('click', { target: popup.children[1] });
    expect(selectSheet).toHaveBeenCalledWith(1);

    navGroup.dispatch('pointerup', { pointerId: 1 });
    flushTimers();
    expect(popup.matches(':popover-open')).toBe(false);
  });

  it('rebuilding while the right-press is still held drops it and releases our capture', () => {
    const { doc, focused, released, timers, flushTimers } = makePopoverDocument();
    const { bar, navGroup, list } = makeBar(doc, 'show', []);
    bar.build(['A', 'B'], [null, null]);

    pressRightAndHold(navGroup, bar.navPrev, 1);
    bar.build(['X', 'Y', 'Z'], [null, null, null]);
    expect(released).toEqual([1]);

    navGroup.dispatch('pointerup', { pointerId: 1 });
    expect(timers.size).toBe(0);
    flushTimers();
    expect(list()!.matches(':popover-open')).toBe(false);
    expect(focused.element).toBeNull();
  });

  it('rebuilding or destroying after release but before the queued open cancels it', () => {
    for (const teardown of ['build', 'destroy'] as const) {
      const { doc, focused, released, timers, flushTimers } = makePopoverDocument();
      const { bar, navGroup, list } = makeBar(doc, 'show', []);
      bar.build(['A', 'B'], [null, null]);

      pressRightAndHold(navGroup, bar.navNext, 1);
      navGroup.dispatch('pointerup', { pointerId: 1 });
      expect(timers.size).toBe(1);

      if (teardown === 'build') bar.build(['X'], [null]);
      else bar.destroy();

      expect(timers.size).toBe(0);
      // The browser already ended the capture on pointerup; nothing to release.
      expect(released).toEqual([]);
      flushTimers();
      expect(list()!.matches(':popover-open')).toBe(false);
      expect(focused.element).toBeNull();
    }
  });
});
