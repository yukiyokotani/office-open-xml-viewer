import { projectBodyOccurrence } from './occurrence-projection.js';
import { buildPageLayers } from './page-graph.js';
import { renderSelectedDocumentPage } from '../paint/canvas-document.js';
import type { LayoutPage } from './types.js';
import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import { snapshotPlainData } from './plain-data.js';
import { appendNativeReadingScenes, nativeReadingOccurrenceIds } from './native-reading-paragraph.js';
import { layoutParagraph } from './paragraph.js';
import { selectParagraphFragment } from './paragraph-pagination.js';
import { createPaintResourceRegistry } from './paint-resources.js';
import { imageResourceKey } from './source-key.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { AcquiredParagraphLayoutInput } from './types.js';
const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
const key = imageResourceKey({ ...source, path: [0, 2] }, 'invented.jpeg');
const registry = createPaintResourceRegistry([{ kind: 'image', resourceKey: key, partPath: 'invented.jpeg', mimeType: 'image/jpeg', intrinsicSize: { widthPt: 40, heightPt: 20 } }]);
function paragraph(requested = true): ParagraphAcquisitionInput {
  return snapshotPlainData({ runs: [ { type: 'text', text: 'all source text' }, { type: 'anchorHost', anchorOccurrenceId: 'picture' },
    { type: 'image', imagePath: 'invented.jpeg', mimeType: 'image/jpeg', anchorAcquisitionInput: {
      occurrenceId: 'picture', ...(requested ? { nativeReadingRelocation: 'completeScene' } : {}), group: null,
      extent: { widthPt: 40, heightPt: 20, widthStatus: 'valid', heightStatus: 'valid' },
      relativeSize: { horizontal: null, vertical: null }, wrap: { kind: 'tight', polygon: null },
    } },
  ] }, 'invented reading paragraph') as unknown as ParagraphAcquisitionInput;
}
function layout(): AcquiredParagraphLayoutInput {
  return { kind: 'paragraph', id: 'paragraph', source, flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 30, yPt: 100, widthPt: 80, heightPt: 18 }, inkBounds: { xPt: 30, yPt: 100, widthPt: 50, heightPt: 12 },
    spacing: { beforePt: 0, afterPt: 6 }, lines: [{ range: { start: 0, end: 15 }, bounds: { xPt: 30, yPt: 100, widthPt: 50, heightPt: 12 }, baselinePt: 110, advancePt: 12, placements: [] }],
    resources: [], drawings: [], textBoxes: [], borders: [], events: [], exclusions: [],
  };
}
const frame = { xPt: 30, yPt: 30, widthPt: 80, heightPt: 100 };
describe('native reading paragraph production candidate', () => {
  it('reading_scene_occurrence preserves complete scene owners through real page projection and preclear admission', async () => {
    const acquired = layoutParagraph(appendNativeReadingScenes(paragraph(), layout(), frame, registry));
    const before = JSON.stringify(acquired);
    const projected = ['page-one/body', 'page-two/body'].map((occurrenceId, index) => projectBodyOccurrence(acquired, {
      occurrenceId,
      destination: { coordinateSpace: 'logical-page-points', flowDomainId: occurrenceId,
        translation: { xPt: 10 + index * 20, yPt: 15 + index * 30 } },
    }));
    for (const selected of projected) {
      // The complete acquired scene remains reachable by its page-local owner;
      // acquisition IDs and source resource keys must not be rewritten in place.
      assert.deepEqual(selected.nativeReadingRelocations, selected.drawings.map(drawing => drawing.id));
      assert.notEqual(selected.drawings[0].id, acquired.drawings[0].id);
      assert.deepEqual(selected.source, acquired.source);
      assert.deepEqual(selected.lines.map(line => line.range), acquired.lines.map(line => line.range));
      const command = selected.drawings[0].commands[0], original = acquired.drawings[0].commands[0];
      assert.equal(command.kind, 'resource'); assert.equal(original.kind, 'resource');
      if (command.kind !== 'resource' || original.kind !== 'resource') throw new Error('acquired scene changed resource class');
      assert.equal(command.resourceKey, key);
      assert.deepEqual(command.nativeImagePlan, original.nativeImagePlan);
    }
    assert.notEqual(projected[0].drawings[0].id, projected[1].drawings[0].id);
    assert.equal(JSON.stringify(acquired), before);
    const pageFor = (node: typeof acquired): LayoutPage => ({
      pageIndex: 0,
      geometry: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300, contentTopPt: 0, contentBottomPt: 300 },
      flowDomains: [],
      section: { geometry: { pageWidth: 200, pageHeight: 300, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
        headerDistance: 0, footerDistance: 0 }, columns: [{ xPt: 0, wPt: 200 }], columnSeparator: false,
        grid: { kind: 'none', linePitchPt: null, charSpacePt: null }, textDirection: 'lrTb', verticalAlignment: 'top' },
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [], columnSeparators: [], pageBorder: null, readingOrder: [],
      layers: buildPageLayers([{ layer: 'body', node }]),
    });
    const preclearRefusal = async (node: typeof acquired, reason: RegExp) => {
      const selected = pageFor(node), operations: string[] = [];
      const target = { width: 87, height: 43, getContext: () => ({
        scale: () => operations.push('scale'), fillRect: () => operations.push('fillRect'),
      }) };
      await assert.rejects(renderSelectedDocumentPage({ pages: [selected], diagnostics: [] }, selected,
        target as unknown as OffscreenCanvas, { dpr: 1, parseError: false, registry,
          rasterPaintOccurrences: [], textRuns: [] }), reason);
      assert.deepEqual([target.width, target.height], [87, 43]);
      assert.deepEqual(operations, []);
    };
    // With no byte source, a reachable image must reach the actual decoded-owner
    // guard, not a spurious missing-drawing error. This is not a decode/paint proof.
    await preclearRefusal(projected[0], /available immutable decoded ImageBitmap/);
    await preclearRefusal({ ...projected[0], drawings: [] }, /Reading page lost a complete drawing/);
    assert.equal(JSON.stringify(acquired), before);
  });

  it('keeps strict missing contour untouched and requires an explicit retained request', () => {
    const base = layout(); assert.strictEqual(appendNativeReadingScenes(paragraph(false), base, frame, registry), base);
    assert.deepEqual(nativeReadingOccurrenceIds(paragraph(false)), []);
  });
  it('retains text and exact resource in one full flow block without changing authored wrap', () => {
    const input = paragraph(), base = layout(), before = JSON.stringify(input);
    const candidate = appendNativeReadingScenes(input, base, frame, registry);
    assert.strictEqual(candidate.lines[0], base.lines[0]); assert.equal(candidate.lines.length, 2);
    assert.equal(candidate.flowBounds.heightPt, 38); assert.deepEqual(candidate.nativeReadingRelocations, ['paragraph:reading:picture']);
    assert.deepEqual(candidate.drawings[0].flowBounds, { xPt: 30, yPt: 112, widthPt: 40, heightPt: 20 });
    assert.equal(candidate.drawings[0].commands[0].kind, 'resource');
    assert.deepEqual(candidate.lines[1].range, { start: 15, end: 15 }); assert.equal(JSON.stringify(input), before);
    assert.equal(base.drawings.length, 0);
  });
  it('rejects intrinsic full-column overflow and source continuation atomically', () => {
    const base = layout(); assert.throws(() => appendNativeReadingScenes(paragraph(), base, { ...frame, heightPt: 37 }, registry), /full column/);
    assert.throws(() => appendNativeReadingScenes(paragraph(), { ...base, continuation: { lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false } }, frame, registry), /complete unclipped/);
    assert.equal(base.drawings.length, 0);
  });
  it('reading_scene_exclusion places the complete scene below chained exclusions and charges every gap', () => {
    const base = layout();
    const exclusion = (id: string, xPt: number, yPt: number, widthPt: number, heightPt: number) =>
      ({ id, wrap: 'square' as const, bounds: { xPt, yPt, widthPt, heightPt }, polygon: [] });
    const input = { ...base, exclusions: [
      exclusion('later', 40, 145, 10, 10), exclusion('first', 50, 120, 10, 10),
      exclusion('beside', 75, 100, 10, 100), exclusion('empty', 30, 113, 10, 0),
    ] };
    const before = JSON.stringify(input), strict = paragraph(false);
    assert.strictEqual(appendNativeReadingScenes(strict, input, frame, registry), input);
    const result = appendNativeReadingScenes(paragraph(), input, frame, registry);
    assert.strictEqual(result.lines[0], base.lines[0]);
    assert.deepEqual(result.drawings[0].flowBounds, { xPt: 30, yPt: 155, widthPt: 40, heightPt: 20 });
    assert.equal(result.flowBounds.heightPt, 81); // text18 + gap43 + scene20, including trailing6.
    assert.strictEqual(result.exclusions, input.exclusions);
    assert.equal(JSON.stringify(input), before);
    assert.equal(result.drawings[0].commands[0].kind, 'resource');
    assert.deepEqual(result.nativeReadingRelocations, ['paragraph:reading:picture']);
  });
  it('reading_scene_exclusion lets the real indivisible selector defer and reacquire the whole paragraph', () => {
    const base = layout(), input = { ...base, exclusions: [{ id: 'prior-float', wrap: 'square' as const,
      bounds: { xPt: 50, yPt: 120, widthPt: 10, heightPt: 10 }, polygon: [] }] };
    const acquired = layoutParagraph(appendNativeReadingScenes(paragraph(), input, frame, registry));
    const cursor = { boundary: null } as const;
    const decision = selectParagraphFragment(acquired, cursor, { kind: 'indivisible' }, 30, 100, true,
      { keepLines: false, widowControl: false, authoredSpaceAfterPt: 6 });
    assert.equal(decision.requiresFreshFlowRegion, true);assert.equal(decision.fragment, null);
    const freshBase = { ...base, flowBounds: { ...base.flowBounds, yPt: 30 },
      inkBounds: { ...base.inkBounds, yPt: 30 },
      lines: base.lines.map(line=>({ ...line, bounds: { ...line.bounds, yPt: 30 }, baselinePt: 40 })) };
    const fresh = layoutParagraph(appendNativeReadingScenes(paragraph(), freshBase, frame, registry));
    const admitted = selectParagraphFragment(fresh, cursor, { kind: 'indivisible' }, 100, 100, false,
      { keepLines: false, widowControl: false, authoredSpaceAfterPt: 6 });
    assert.strictEqual(admitted.fragment, fresh);assert.equal(admitted.requiresFreshFlowRegion, false);
    assert.equal(fresh.drawings[0].flowBounds.yPt, 42);
    assert.deepEqual(fresh.lines[0].range, acquired.lines[0].range);
    assert.deepEqual(fresh.nativeReadingRelocations, acquired.nativeReadingRelocations);
    assert.equal(fresh.drawings[0].commands[0].kind, 'resource');
  });
  it('reading_scene_exclusion keeps malformed, over-budget and unplaceable candidates atomic', () => {
    const base = layout();
    const blocked = { ...base, exclusions: [{ id: 'large', wrap: 'square' as const,
      bounds: { xPt: 30, yPt: 120, widthPt: 40, heightPt: 100 }, polygon: [] }] };
    assert.throws(()=>appendNativeReadingScenes(paragraph(), blocked, frame, registry), /full column/);
    const invalid = { ...base, exclusions: [{ ...blocked.exclusions[0], bounds: { ...blocked.exclusions[0].bounds, yPt: Infinity } }] };
    assert.throws(()=>appendNativeReadingScenes(paragraph(), invalid, frame, registry), /exclusion/);
    const many = { ...base, exclusions: Array.from({length:100},(_,i)=>({ ...blocked.exclusions[0], id:String(i) })) };
    assert.throws(()=>appendNativeReadingScenes(paragraph(), many, frame, registry, 20), /allowance/);
    assert.equal(base.drawings.length,0);assert.equal(blocked.drawings.length,0);
  });
  it('rejects a later unsupported scene instead of publishing the first scene or changing input', () => {
    const raw = JSON.parse(JSON.stringify(paragraph())); raw.runs.push({ type: 'anchorHost', anchorOccurrenceId: 'unsupported' });
    raw.runs.push({ type: 'unavailableDrawing', anchorAcquisitionInput: { ...raw.runs[2].anchorAcquisitionInput, occurrenceId: 'unsupported' } });
    const input = snapshotPlainData(raw, 'invented later scene') as ParagraphAcquisitionInput, base = layout();
    assert.throws(() => appendNativeReadingScenes(input, base, frame, registry)); assert.equal(base.drawings.length, 0);
  });
});
