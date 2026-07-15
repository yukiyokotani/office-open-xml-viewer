import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import { deepFreezeDocumentLayout, layoutFingerprint } from './invariants.js';
import type { DocumentLayout } from './types.js';

describe('DocumentLayout data boundary', () => {
  it('is deeply immutable and structured-clone safe without platform objects', () => {
    const layout: DocumentLayout = {
      pages: [{
        pageIndex: 0,
        geometry: {
          xPt: 0,
          yPt: 0,
          widthPt: 612,
          heightPt: 792,
          contentTopPt: 72,
          contentBottomPt: 720,
        },
        flowDomains: [{
          id: 'body',
          kind: 'body',
          bounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
        }],
        section: {
          geometry: {
            pageWidth: 100,
            pageHeight: 200,
            marginTop: 10,
            marginRight: 10,
            marginBottom: 10,
            marginLeft: 10,
            headerDistance: 5,
            footerDistance: 5,
          },
          columns: [{ xPt: 10, wPt: 80 }],
          grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
          textDirection: 'lrTb',
          verticalAlignment: 'top',
        } satisfies SectionLayoutContext,
        sectionOccurrenceId: 'section:0',
        parityBlank: false,
        bookmarkStarts: [],
        pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
        sectionRegions: [{
          id: 'region:0', sectionOccurrenceId: 'section:0',
          coordinateSpace: { writingMode: 'horizontal-tb', logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
          blockStartPt: 10, blockEndPt: 190, flowDomainIds: ['body'],
          section: {} as SectionLayoutContext,
        }],
        layers: {
          paintOrder: [{ layer: 'front', nodeId: 'drawing-1', coordinateSpace: 'physical-page-points' }],
          background: [],
          behindText: [],
          header: [],
          body: [],
          notes: [],
          front: [{
            kind: 'drawing',
            id: 'drawing-1',
            source: { story: 'body', storyInstance: 'body', path: [0, 1] },
            flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            inkBounds: { xPt: 9, yPt: 19, widthPt: 32, heightPt: 42 },
            clipBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            advancePt: 0,
            ordinaryFlow: false,
            flowDomainId: 'body',
            transform: { a: 1, b: 0, c: 0, d: 1, e: 10, f: 20 },
            clip: {
              kind: 'polygon',
              points: [{ xPt: 0, yPt: 0 }, { xPt: 30, yPt: 0 }, { xPt: 30, yPt: 40 }],
            },
            commands: [{
              kind: 'fill-rect',
              rect: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
              fill: '#123456',
            }],
          }],
          footer: [],
        },
        readingOrder: ['drawing-1'],
      }],
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        message: 'synthetic diagnostic',
      }],
    };

    const frozen = deepFreezeDocumentLayout(layout);
    const cloned = structuredClone(frozen);
    const clonedNode = cloned.pages[0]?.layers.front[0];
    if (!clonedNode || clonedNode.kind !== 'drawing') throw new Error('drawing clone missing');

    expect(Object.isFrozen(frozen)).toBe(true);
    expect(Object.isFrozen(frozen.pages)).toBe(true);
    expect(Object.isFrozen(frozen.pages[0]?.layers.front[0]?.source.path)).toBe(true);
    expect(Object.getPrototypeOf(clonedNode.transform)).toBe(Object.prototype);
    expect(layoutFingerprint(cloned)).toBe(layoutFingerprint(frozen));
    expect(() => (frozen as unknown as { pages: unknown[] }).pages.push({})).toThrow();
  });

  it.each([
    ['Date', new Date(0)],
    ['Map', new Map()],
    ['WeakMap', new WeakMap()],
    ['function', () => undefined],
    ['symbol', Symbol('layout')],
  ])('rejects non-plain %s values before freeze or fingerprint', (_name, invalid) => {
    const layout = {
      pages: [],
      diagnostics: [],
      invalid,
    } as unknown as DocumentLayout;

    expect(() => deepFreezeDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFingerprint(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it('rejects cyclic records', () => {
    const layout = { pages: [], diagnostics: [] } as DocumentLayout & { self?: unknown };
    layout.self = layout;

    expect(() => deepFreezeDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFingerprint(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it('rejects array properties that JSON fingerprints would omit', () => {
    const pages: unknown[] & { hidden?: string } = [];
    pages.hidden = 'not retained by JSON';
    const layout = { pages, diagnostics: [] } as unknown as DocumentLayout;

    expect(() => deepFreezeDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFingerprint(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it.each(['accessor index', 'non-enumerable property'])('rejects arrays with an %s', (kind) => {
    const pages: unknown[] = [];
    if (kind === 'accessor index') {
      Object.defineProperty(pages, '0', { enumerable: true, get: () => ({}) });
    } else {
      Object.defineProperty(pages, 'hidden', { value: 'hidden', enumerable: false });
    }
    const layout = { pages, diagnostics: [] } as unknown as DocumentLayout;

    expect(() => deepFreezeDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFingerprint(layout)).toThrow(/INVALID_GEOMETRY/);
  });
});
