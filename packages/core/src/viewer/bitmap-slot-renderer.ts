import { StaticCanvasRenderDispatcher } from '../internal/canvas-viewer-mechanics';

export interface BitmapSlot {
  canvas: HTMLCanvasElement;
  dispatcher: StaticCanvasRenderDispatcher;
  renderedScale: number;
}

export interface BitmapSlotHooks<Slot extends BitmapSlot, Run> {
  slots(): ReadonlyMap<number, Slot>;
  inFlight(): Set<number>;
  epoch(): number;
  destroyed(): boolean;
  scale(): number;
  width(unit: number): number;
  dpr(): number;
  slotIndex(slot: Slot): number;
  token(slot: Slot): number;
  nextToken(slot: Slot): number;
  canRetry(slot: Slot): boolean;
  wantRuns(slot: Slot): boolean;
  render(unit: number, canvas: HTMLCanvasElement, width: number, dpr: number,
    onRun: ((run: Run) => void) | undefined): Promise<ImageBitmap>;
  commitBitmap(unit: number, slot: Slot, dispatcher: StaticCanvasRenderDispatcher,
    generation: number, bitmap: ImageBitmap, width: number): boolean;
  commitRuns(unit: number, slot: Slot, runs: Run[], width: number, wantedRuns: boolean): void;
  reportError(error: unknown): void;
}

/** Coalesces worker bitmap dispatches and follows a superseded dispatch until
 * the live slot receives current pixels. A recycled slot can return to the same
 * unit while its old-scale bitmap is in flight; epoch, slot, canvas and format
 * token guards jointly close that orphan. A plain render failure never retries
 * because it could otherwise create an unbounded reject/redispatch loop. */
export class BitmapSlotRenderer<Slot extends BitmapSlot, Run> {
  /** The in-flight paint of each unit, including the dispatches it follows. */
  private readonly pending = new Map<number, Promise<void>>();

  constructor(private readonly hooks: BitmapSlotHooks<Slot, Run>) {}

  /** Paint `unit` into `slot`. A request for a unit already in flight is
   *  coalesced into that dispatch, which follows a superseded dispatch until
   *  the live slot is painted; the returned promise settles with that final
   *  paint, so a caller can await the pixels it asked for. */
  render(
    unit: number,
    slot: Slot,
    width: number,
    dpr: number,
    scale: number,
    token = this.hooks.nextToken(slot),
    dispatcher = slot.dispatcher,
    generation = dispatcher.begin(),
    reportErrors = true,
  ): Promise<void> {
    if (this.hooks.inFlight().has(unit)) {
      const pending = this.pending.get(unit) ?? Promise.resolve();
      // The coalesced dispatch reports its own failure to its first caller.
      return pending.catch(() => {});
    }
    if (this.hooks.slots().get(unit) !== slot) return Promise.resolve();
    const paint = this.paint(unit, slot, width, dpr, scale, token, dispatcher, generation, reportErrors);
    this.pending.set(unit, paint);
    const forget = () => { if (this.pending.get(unit) === paint) this.pending.delete(unit); };
    paint.then(forget, forget);
    return paint;
  }

  private async paint(
    unit: number,
    slot: Slot,
    width: number,
    dpr: number,
    scale: number,
    token: number,
    dispatcher: StaticCanvasRenderDispatcher,
    generation: number,
    reportErrors: boolean,
  ): Promise<void> {
    const inFlight = this.hooks.inFlight();
    const epoch = this.hooks.epoch();
    const canvas = slot.canvas;
    inFlight.add(unit);
    let painted = false;
    const runs: Run[] = [];
    const wantedRuns = this.hooks.wantRuns(slot);
    try {
      const bitmap = await this.hooks.render(
        unit, canvas, width, dpr,
        wantedRuns ? (run) => runs.push(run) : undefined,
      );
      if (!this.isCurrent(unit, slot, token, dispatcher, generation, canvas, epoch)) {
        bitmap.close();
        return;
      }
      if (!this.hooks.commitBitmap(unit, slot, dispatcher, generation, bitmap, width)) return;
      slot.renderedScale = scale;
      this.hooks.commitRuns(unit, slot, runs, width, wantedRuns);
      painted = true;
    } catch (error) {
      if (this.isCurrent(unit, slot, token, dispatcher, generation, canvas, epoch)) {
        if (reportErrors) this.hooks.reportError(error);
        else throw error;
      }
    } finally {
      inFlight.delete(unit);
      const live = this.hooks.slots().get(unit);
      if (!painted && live && this.hooks.canRetry(live) &&
          (live !== slot || epoch !== this.hooks.epoch() ||
            token !== this.hooks.token(live) || !dispatcher.isCurrent(generation)) &&
          !inFlight.has(unit) && !this.hooks.destroyed()) {
        const nextDispatcher = live.dispatcher;
        await this.render(
          unit, live, this.hooks.width(unit), this.hooks.dpr(), this.hooks.scale(),
          this.hooks.nextToken(live), nextDispatcher, nextDispatcher.begin(), reportErrors,
        );
      }
    }
  }

  private isCurrent(
    unit: number, slot: Slot, token: number, dispatcher: StaticCanvasRenderDispatcher,
    generation: number, canvas: HTMLCanvasElement, epoch: number,
  ): boolean {
    return token === this.hooks.token(slot) && dispatcher.isCurrent(generation) &&
      canvas === slot.canvas && epoch === this.hooks.epoch() &&
      this.hooks.slots().get(unit) === slot && this.hooks.slotIndex(slot) === unit;
  }
}
