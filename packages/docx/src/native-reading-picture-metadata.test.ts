import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import { anchorAcquisitionInput } from './parser-model.js';
import { nativeReadingNotices, retainedReadingNotices } from './native-reading-notice.js';
import { appendNativeReadingScenes } from './layout/native-reading-paragraph.js';
import { snapshotPlainData } from './layout/plain-data.js';
import { createPaintResourceRegistry } from './layout/paint-resources.js';
import { imageResourceKey } from './layout/source-key.js';
import type { NativePictureMetadataInput } from './layout/anchor-input.js';
import type { ParagraphAcquisitionInput } from './layout/text.js';
import type { AcquiredParagraphLayoutInput, DocumentLayout } from './layout/types.js';

const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
const key = imageResourceKey({ ...source, path: [0, 2] }, 'invented.jpeg');
const registry = createPaintResourceRegistry([{ kind: 'image', resourceKey: key, partPath: 'invented.jpeg', mimeType: 'image/jpeg', intrinsicSize: { widthPt: 40, heightPt: 20 } }]);
function metadata(retention: 'inactiveOpaqueNotDecoded' | 'ignoredZeroIndex'): NativePictureMetadataInput {
  return { blipName: null, shapeName: null,
    description: { scope: 'shape', opid: 0x8381, value: 4, text: 'D', rawBytes: [68, 0, 0, 0], retention: 'passiveName' },
    inactiveFillCarrier: { scope: 'documentDefault', opid: retention === 'ignoredZeroIndex' ? 0x4186 : 0x8186,
      value: 0, text: null, rawBytes: [], retention }, inactiveLineCarrier: null };
}
function paragraph(value?: NativePictureMetadataInput): ParagraphAcquisitionInput {
  const input = anchorAcquisitionInput({ __anchorAcquisition: {
    occurrenceId: 'picture', nativeReadingRelocation: 'completeScene', group: null,
    ...(value ? { nativePictureMetadata: value } : {}),
    extent: { widthPt: 40, heightPt: 20, widthStatus: 'valid', heightStatus: 'valid' },
    relativeSize: { horizontal: null, vertical: null }, wrap: { kind: 'tight', polygon: null },
  } });
  return snapshotPlainData({ runs: [{ type: 'text', text: 'all source text' },
    { type: 'anchorHost', anchorOccurrenceId: 'picture' },
    { type: 'image', imagePath: 'invented.jpeg', mimeType: 'image/jpeg', anchorAcquisitionInput: input },
  ] }, 'invented passive picture properties') as unknown as ParagraphAcquisitionInput;
}
function base(): AcquiredParagraphLayoutInput {
  return { kind: 'paragraph', id: 'paragraph', source, flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 30, yPt: 100, widthPt: 80, heightPt: 18 }, inkBounds: { xPt: 30, yPt: 100, widthPt: 50, heightPt: 12 },
    spacing: { beforePt: 0, afterPt: 6 }, lines: [{ range: { start: 0, end: 15 }, bounds: { xPt: 30, yPt: 100, widthPt: 50, heightPt: 12 }, baselinePt: 110, advancePt: 12, placements: [] }],
    resources: [], drawings: [], textBoxes: [], borders: [], events: [], exclusions: [] };
}
const frame = { xPt: 30, yPt: 30, widthPt: 80, heightPt: 100 };
function document(node: AcquiredParagraphLayoutInput): DocumentLayout {
  return { pages: [{ layers: { body: [node] } }], diagnostics: [] } as unknown as DocumentLayout;
}
describe('passive native picture metadata reading contract', () => {
  it('retains raw metadata through the real wire boundary without changing paint resources or source text', () => {
    const raw = metadata('inactiveOpaqueNotDecoded'), input = paragraph(raw), before = JSON.stringify(input);
    const projected = appendNativeReadingScenes(input, base(), frame, registry);
    const control = appendNativeReadingScenes(paragraph(), base(), frame, registry);
    assert.deepEqual(projected.drawings, control.drawings);
    assert.deepEqual(projected.lines, control.lines);
    assert.deepEqual(projected.resources, control.resources);
    assert.equal(JSON.stringify(input), before);
    const retained = input.runs[2]; assert.equal(retained?.type, 'image');
    if (retained?.type !== 'image') throw new Error('invented fixture lost image');
    assert.deepEqual(retained.anchorAcquisitionInput?.nativePictureMetadata, raw);
    assert.deepEqual(structuredClone(retained.anchorAcquisitionInput?.nativePictureMetadata), raw);
    assert.equal(projected.nativeReadingInactivePictureData, true);
    const notices = nativeReadingNotices(document(projected));
    assert.deepEqual(notices.map(notice => notice.code), ['DRAWINGS_RELOCATED_FOR_READING', 'INACTIVE_PICTURE_DATA_RETAINED']);
    assert.equal(notices[1]?.message, 'Inactive picture formatting was retained as metadata without interpretation');
    assert.deepEqual(retainedReadingNotices(structuredClone(notices)), notices);
    assert.ok(Object.isFrozen(retainedReadingNotices(structuredClone(notices))));
  });
  it('does not label ignored zero indices or names as undecoded inactive paint, and rejects forged disclosures', () => {
    const ignored = appendNativeReadingScenes(paragraph(metadata('ignoredZeroIndex')), base(), frame, registry);
    assert.equal(ignored.nativeReadingInactivePictureData, undefined);
    assert.equal(nativeReadingNotices(document(ignored)).length, 1);
    const passive = { ...metadata('ignoredZeroIndex'), inactiveFillCarrier: null };
    const named = appendNativeReadingScenes(paragraph(passive), base(), frame, registry);
    assert.equal(nativeReadingNotices(document(named)).length, 1);
    assert.throws(() => retainedReadingNotices([{ code: 'INACTIVE_PICTURE_DATA_RETAINED', message: 'Inactive picture formatting was retained as metadata without interpretation' }]), /Invalid reading-layout disclosure/);
    assert.throws(() => retainedReadingNotices([{ code: 'DRAWINGS_RELOCATED_FOR_READING', message: 'arbitrary text' }]), /Invalid reading-layout disclosure/);
  });
});
