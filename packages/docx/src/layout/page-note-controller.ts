import type { PageNoteFlow, PageNoteSource } from './page-note-area.js';

/** Internal control transfer to the live page's immutable source checkpoint. */
export class PageNoteReplay extends Error {}

type NoteAreaBuilder = (sources: readonly PageNoteSource[], capacityPt: number, balanced: boolean,
  retainAllHeads?: boolean) => PageNoteFlow;

interface NoteAreaSearch {
  readonly baseSources: readonly PageNoteSource[];
  readonly baseRequired: readonly string[];
  readonly baseArea: PageNoteFlow;
  readonly addedIds: readonly string[];
  lowerPt: number;
  upperPt: number;
  capacityPt: number;
  best?: PageNoteFlow;
  finalizing: boolean;
  readonly outcomes: Map<number, boolean>;
}

/** A single live page's note-space transaction. No preceding page or body
 * prefix is retained here. Trial acceptance is grounded by actual committed
 * reference occurrences, never speculative shaping or a source flag. Trials
 * stop as soon as their last required reference commits; rejection happens
 * before advancing the physical page. The caller replays that page through
 * its ordinary body paginator and shares its acquisition/work ledgers.
 *
 * A later ordinary reference is deferred while an earlier reference's trial
 * is unresolved. Mandatory keep/widow-unit references instead join one trial:
 * rejecting that group relocates the same unit the body selector requires.
 */
export class PageNoteController {
  private sources: readonly PageNoteSource[];
  private required: readonly string[] = [];
  private readonly actual = new Set<string>();
  private readonly blocked = new Set<string>();
  private readonly incoming: ReadonlySet<string>;
  private search: NoteAreaSearch | undefined;
  private evaluations = 0;
  private retained: PageNoteFlow;

  constructor(
    sources: readonly PageNoteSource[],
    private readonly maximumPt: number,
    private readonly build: NoteAreaBuilder,
    private readonly allowContinuation: boolean,
  ) {
    this.sources = sources;
    this.incoming = new Set(sources.map(source => source.id));
    this.retained = this.evaluate(maximumPt, true);
    this.resetActual();
  }

  get area(): PageNoteFlow { return this.retained; }

  /** Reset only the page's real-reference journal before replay. Incoming
   * fragments own this page without requiring another body reference. */
  resetActual(): readonly string[] {
    this.actual.clear();
    for (const chunk of this.retained.chunks) {
      if (this.incoming.has(chunk.layout.source.storyInstance)) this.actual.add(chunk.layout.source.storyInstance);
    }
    return Object.freeze([...this.actual]);
  }

  /** False asks the ordinary selector to relocate this reference. New plans
   * throw PageNoteReplay so every column sees their common reserve. */
  discover(ids: readonly string[], mandatoryUnit: boolean, allowPartial?: boolean): boolean {
    if (ids.some(id => this.blocked.has(id))) return false;
    const known = new Set(this.sources.map(source => source.id));
    const added = [...new Set(ids)].filter(id => !known.has(id));
    if (added.length === 0) return true;
    if (this.search && !mandatoryUnit) return false;
    const prior = this.search;
    const baseSources = prior?.baseSources ?? this.sources;
    const baseRequired = prior?.baseRequired ?? this.required;
    const baseArea = prior?.baseArea ?? this.retained;
    this.sources = Object.freeze([...this.sources,
      ...added.map((id, index) => Object.freeze({ id, cursor: null,
        allowPartial: allowPartial === undefined || (allowPartial && index === added.length - 1),
      }))]);
    this.required = Object.freeze([...this.required, ...added]);
    this.search = {
      baseSources, baseRequired, baseArea,
      addedIds: Object.freeze([...(prior?.addedIds ?? []), ...added]),
      lowerPt: 0, upperPt: this.maximumPt, capacityPt: this.maximumPt, finalizing: false, outcomes: new Map(),
    };
    const first = this.evaluate(this.maximumPt, false);
    // A source-ordered prefix unable to reach every required head at full
    // capacity cannot coexist with this reference unit on this page.
    if (!this.containsRequired(first)) return this.rejectGroup();
    this.retained = first;
    throw new PageNoteReplay();
  }

  committed(ids: readonly string[]): void {
    if (ids.length === 0) return;
    // A fresh-region body selector can force an indivisible reference unit
    // after ordinary admission rejects it. That is permission for overflow,
    // never permission to commit references without their note content.
    // Replay with every required head retained; subsequent pages own tails.
    const heads = new Set(this.retained.chunks.map(chunk => chunk.layout.source.storyInstance));
    if (ids.some(id => !heads.has(id))) {
      if (!ids.every(id => heads.has(id) || this.blocked.has(id))) {
        throw new Error('A committed reference has no retained note head');
      }
      const known = new Set(this.sources.map(source => source.id));
      this.sources = Object.freeze([...this.sources,
        ...ids.filter(id => !known.has(id)).map(id => Object.freeze({ id, cursor: null }))]);
      this.required = Object.freeze([...new Set([...this.required, ...ids])]);
      this.retained = this.build(this.sources, this.maximumPt, false, true);
      const retainedHeads = new Set(this.retained.chunks.map(chunk => chunk.layout.source.storyInstance));
      if (!this.required.every(id => retainedHeads.has(id))) throw new Error('Overflow omitted a required note head');
      ids.forEach(id => this.blocked.delete(id));
      throw new PageNoteReplay();
    }
    ids.forEach(id => this.actual.add(id));
    const search = this.search;
    if (!search || !this.required.every(id => this.actual.has(id))) return;
    if (search.finalizing) {
      this.search = undefined;
      return;
    }
    search.outcomes.set(this.retained.heightPt, true);
    search.best = this.retained;
    search.lowerPt = search.capacityPt;
    if (search.capacityPt === this.maximumPt) this.finalizeSearch();
    this.nextTrial();
  }

  beforePageClose(): void {
    const search = this.search;
    if (!search) return;
    if (search.finalizing) throw new Error('Balanced page note area lost an admitted reference');
    search.outcomes.set(this.retained.heightPt, false);
    search.upperPt = search.capacityPt;
    this.nextTrial();
  }

  private containsRequired(area: PageNoteFlow): boolean {
    if (!this.allowContinuation && area.remaining.length > 0) return false;
    const heads = new Set(area.chunks.map(chunk => chunk.layout.source.storyInstance));
    return this.required.every(id => heads.has(id));
  }

  private evaluate(capacityPt: number, balanced: boolean): PageNoteFlow {
    // Resource policy, shared by all reference groups on this one page. Work
    // credits are not restored when body/source state rolls back.
    if (++this.evaluations > 1024) throw new Error('Page note planning exceeded its operational evaluation limit');
    return this.build(this.sources, capacityPt, balanced);
  }

  private nextTrial(): never {
    const search = this.search!;
    for (;;) {
      const middlePt = search.lowerPt + (search.upperPt - search.lowerPt) / 2;
      if (middlePt === search.lowerPt || middlePt === search.upperPt) this.finalizeSearch();
      search.capacityPt = middlePt;
      const trial = this.evaluate(middlePt, false);
      if (!this.containsRequired(trial)) {
        search.lowerPt = middlePt;
        continue;
      }
      // Body acquisition observes only the common reserve, not how this
      // source prefix is distributed inside the independent note story. An
      // exactly repeated reserve reuses its actual-reference verdict; no
      // nearby height or different required-reference group is inferred.
      const known = search.outcomes.get(trial.heightPt);
      if (known !== undefined) {
        if (known) { search.best = trial; search.lowerPt = middlePt; }
        else search.upperPt = middlePt;
        continue;
      }
      this.retained = trial;
      throw new PageNoteReplay();
    }
  }

  private finalizeSearch(): never {
    const search = this.search!;
    if (!search.best) return this.rejectGroup();
    this.retained = this.evaluate(search.best.heightPt, true);
    if (!this.containsRequired(this.retained)) throw new Error('Balanced page note area lost a required note head');
    search.finalizing = true;
    throw new PageNoteReplay();
  }

  private rejectGroup(): never {
    const search = this.search!;
    search.addedIds.forEach(id => this.blocked.add(id));
    this.sources = search.baseSources;
    this.required = search.baseRequired;
    this.retained = search.baseArea;
    this.search = undefined;
    throw new PageNoteReplay();
  }
}
