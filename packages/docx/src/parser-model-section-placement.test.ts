import { describe, expect, it } from 'vitest';
import {
  bodySectionIndexInput,
  normalizeInternalDocumentModel,
  sectionPlacementInputFrom,
  sectionPlacementInputFromBody,
} from './parser-model.js';
import type { BodyElement, DocxDocumentModel, LineNumbering, SectionProps } from './types.js';

interface PrivateSectionPlacementWire {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: LineNumbering | null;
}

type PrivateSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement: PrivateSectionPlacementWire;
};

function sectionBreak(sectionId: string): BodyElement {
  return {
    type: 'sectionBreak', kind: 'nextPage', columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
    __sectionPlacement: {
      sectionId,
      vAlign: 'center',
      lineNumbering: { countBy: 2, start: 7, distance: 12, restart: 'newSection' },
    },
  } as PrivateSectionBreak;
}

function model(): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', vAlign: 'bottom',
    lineNumbering: { countBy: 1, start: 3, distance: 8, restart: 'newPage' },
  } as SectionProps;
  return {
    section,
    body: [sectionBreak('section:0')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('parser-private section placement projection', () => {
  it('survives the worker structured-clone boundary with stable section identity', () => {
    const main = normalizeInternalDocumentModel(model()).document;
    const worker = normalizeInternalDocumentModel(structuredClone(model())).document;

    const mainEnding = sectionPlacementInputFrom(main, 0);
    const workerEnding = sectionPlacementInputFrom(worker, 0);
    const mainFinal = sectionPlacementInputFrom(main, 1);
    const workerFinal = sectionPlacementInputFrom(worker, 1);

    expect(workerEnding).toEqual(mainEnding);
    expect(workerEnding).toEqual({
      sectionId: 'section:0',
      vAlign: 'center',
      lineNumbering: { countBy: 2, start: 7, distance: 12, restart: 'newSection' },
    });
    expect(workerFinal).toEqual(mainFinal);
    expect(workerFinal.sectionId).toBe('section:1');
    expect(Object.isFrozen(workerEnding)).toBe(true);
    expect(Object.isFrozen(workerEnding.lineNumbering)).toBe(true);
  });

  it('projects distinct section occurrences and body ownership as immutable plain input', () => {
    const input = model();
    input.body.push({ type: 'paragraph', runs: [] } as unknown as BodyElement);

    const projected = bodySectionIndexInput(input);

    expect(projected.bodyLength).toBe(2);
    expect(projected.occurrences.map((occurrence) => ({
      id: occurrence.sectionOccurrenceId,
      start: occurrence.startBodyIndex,
      end: occurrence.endBodyIndex,
      final: occurrence.final,
    }))).toEqual([
      { id: 'section:0', start: 0, end: 0, final: false },
      { id: 'section:1', start: 1, end: 1, final: true },
    ]);
    expect(Object.isFrozen(projected)).toBe(true);
    expect(Object.isFrozen(projected.occurrences)).toBe(true);
    expect(Object.isFrozen(projected.occurrences[0]?.placement.lineNumbering)).toBe(true);
    expect(structuredClone(projected)).toEqual(projected);
  });

  it('snapshots placement facts without freezing or retaining caller-owned nested objects', () => {
    const input = model();
    const finalLineNumbering = input.section.lineNumbering!;
    const endingWire = (input.body[0] as PrivateSectionBreak).__sectionPlacement;
    const endingLineNumbering = endingWire.lineNumbering!;

    const normalized = normalizeInternalDocumentModel(input).document;
    const endingSnapshot = sectionPlacementInputFrom(normalized, 0);
    const finalSnapshot = sectionPlacementInputFrom(normalized, 1);

    expect(Object.isFrozen(finalLineNumbering)).toBe(false);
    expect(Object.isFrozen(endingLineNumbering)).toBe(false);
    expect(finalSnapshot.lineNumbering).not.toBe(finalLineNumbering);
    expect(endingSnapshot.lineNumbering).not.toBe(endingLineNumbering);

    finalLineNumbering.start = 101;
    endingLineNumbering.countBy = 99;
    expect(finalSnapshot.lineNumbering?.start).toBe(3);
    expect(endingSnapshot.lineNumbering?.countBy).toBe(2);
    expect(Object.isFrozen(finalSnapshot.lineNumbering)).toBe(true);
    expect(Object.isFrozen(endingSnapshot.lineNumbering)).toBe(true);
  });

  it('keys paginator sidecars by both body and final-section identity', () => {
    const body: BodyElement[] = [];
    const base = model().section;
    const top = {
      ...base,
      vAlign: 'top',
      lineNumbering: { countBy: 2, start: 1, distance: 8, restart: 'newPage' },
    } as SectionProps;
    const bottom = {
      ...base,
      vAlign: 'bottom',
      lineNumbering: { countBy: 3, start: 7, distance: 12, restart: 'newSection' },
    } as SectionProps;

    const first = sectionPlacementInputFromBody(body, top, 0);
    const second = sectionPlacementInputFromBody(body, bottom, 0);
    const switchedBack = sectionPlacementInputFromBody(body, top, 0);

    expect(first).toMatchObject({
      vAlign: 'top',
      lineNumbering: { countBy: 2, start: 1, distance: 8, restart: 'newPage' },
    });
    expect(second).toMatchObject({
      vAlign: 'bottom',
      lineNumbering: { countBy: 3, start: 7, distance: 12, restart: 'newSection' },
    });
    expect(second).not.toBe(first);
    expect(switchedBack).toBe(first);
    expect(Object.isFrozen(top.lineNumbering)).toBe(false);
    expect(Object.isFrozen(bottom.lineNumbering)).toBe(false);
  });
});
