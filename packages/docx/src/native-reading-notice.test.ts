import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import { nativeReadingNotices, noReadingNotices, updateNativeReadingNotice } from './native-reading-notice.js';
import type { DocumentLayout } from './layout/types.js';
function layout(relocated: boolean): DocumentLayout {
  return { pages: [{ layers: { body: [{ kind: 'paragraph', ...(relocated ? { nativeReadingRelocations: ['actual-retained-drawing'] } : {}) }] } }], diagnostics: [] } as unknown as DocumentLayout;
}
describe('changed reading-layout disclosure', () => {
  it('projects only actual retained scenes, preserves strict empty and returns immutable public text', () => {
    assert.strictEqual(nativeReadingNotices(layout(false)), noReadingNotices);
    const notices = nativeReadingNotices(layout(true));
    assert.equal(notices[0].message, 'Some drawings were moved into separate blocks for readability');
    assert.equal(notices[0].code, 'DRAWINGS_RELOCATED_FOR_READING');
    assert.ok(Object.isFrozen(notices)); assert.ok(Object.isFrozen(notices[0]));
  });
  it('inserts accessible text before the first page and removes it on replacement or failure', () => {
    const children: object[] = [{ name: 'first-page' }];
    const attributes: Record<string, string> = {};
    const element = { textContent: '', setAttribute: (key: string, value: string) => { attributes[key] = value; }, remove: () => { children.splice(children.indexOf(element), 1); } };
    const container = { ownerDocument: { createElement: () => element }, firstChild: children[0], insertBefore: (notice: object, page: object) => children.splice(children.indexOf(page), 0, notice) };
    const inserted = updateNativeReadingNotice(container as unknown as HTMLElement, null, nativeReadingNotices(layout(true)));
    assert.strictEqual(children[0], element); assert.equal(attributes.role, 'status'); assert.equal(attributes['aria-live'], 'polite');
    assert.equal(element.textContent, nativeReadingNotices(layout(true))[0].message);
    assert.strictEqual(updateNativeReadingNotice(container as unknown as HTMLElement, inserted, noReadingNotices), null);
    assert.equal(children.length, 1); assert.deepEqual(children[0], { name: 'first-page' });
  });
});
