/** Retained device-space path for [MS-EMF] 2.3.10 path brackets.
 * Canvas's current path is not retained across unrelated beginPath calls.
 * Store geometry only; the brush, pen and filling mode are chosen at playback.
 */
type Command =
  | ['moveTo' | 'lineTo', number, number]
  | ['bezierCurveTo', number, number, number, number, number, number]
  | ['closePath'];
type PathContext = Pick<CanvasRenderingContext2D,
  'beginPath' | 'moveTo' | 'lineTo' | 'bezierCurveTo' | 'closePath'>;
interface Node {
  readonly command: Command;
  readonly previous: Node | null;
}

/** Shared across all paths/DC snapshots in one playback; allocation work limit,
 * not a format restriction. Snapshot roots share immutable command nodes. */
export interface EmfPathBudget {
  remaining: number;
  replayRemaining: number;
}
export const createEmfPathBudget = (): EmfPathBudget => ({ remaining: 0x40000, replayRemaining: 0x100000 });

// Resource policy, not an EMF limit: bound retained commands across records.
// On overflow discard the entire path, never paint a truncated outline.
const MAX_PATH_COMMANDS = 0x10000;

export class EmfPath {
  private tail: Node | null = null;
  private length = 0;
  private valid = true;
  private needsMove = true;

  constructor(private readonly budget = createEmfPathBudget()) {}

  /** [MS-EMF] 3.1.1.2 / 3.1.1.2.4: Path and PathBracket are DC state.
   * Constant-time snapshot; subsequent appends cannot mutate saved geometry. */
  snapshot(): EmfPath {
    const copy = new EmfPath(this.budget);
    copy.tail = this.tail;
    copy.length = this.length;
    copy.valid = this.valid;
    copy.needsMove = this.needsMove;
    return copy;
  }

  invalidate(): void {
    this.valid = false;
    this.tail = null;
    this.length = 0;
  }

  private append(command: Command): void {
    if (!this.valid) return;
    if (this.length >= MAX_PATH_COMMANDS || this.budget.remaining <= 0) {
      this.invalidate();
      return;
    }
    for (let i = 1; i < command.length; i++) {
      if (!Number.isFinite(command[i])) {
        this.invalidate();
        return;
      }
    }
    this.budget.remaining--;
    this.tail = { command, previous: this.tail };
    this.length++;
  }

  moveTo(x: number, y: number): void {
    this.append(['moveTo', x, y]);
    this.needsMove = false;
  }
  /** TO records start from the DC current position only at a new figure. */
  continueFrom(x: number, y: number): void {
    if (this.needsMove) this.moveTo(x, y);
  }
  lineTo(x: number, y: number): void {
    this.append(['lineTo', x, y]);
  }
  bezierCurveTo(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.append(['bezierCurveTo', a, b, c, d, e, f]);
  }
  closePath(): void {
    if (!this.needsMove) this.append(['closePath']);
    this.needsMove = true;
  }

  replay(ctx: PathContext, closeFigures = false): boolean {
    if (!this.valid || this.length === 0) return false;
    // SaveDC/RestoreDC can replay a large shared path many times without
    // allocating geometry. Bound that work separately, before any Canvas call.
    if (this.budget.replayRemaining < this.length) return false;
    this.budget.replayRemaining -= this.length;
    // At most MAX_PATH_COMMANDS references, released after this synchronous
    // replay. The DC stack never duplicates the coordinate arrays.
    const commands: Command[] = [];
    for (let node = this.tail; node; node = node.previous) commands.push(node.command);
    ctx.beginPath();
    let open = false;
    for (let i = commands.length - 1; i >= 0; i--) {
      const command = commands[i];
      switch (command[0]) {
        case 'moveTo':
          if (closeFigures && open) ctx.closePath();
          ctx.moveTo(command[1], command[2]);
          open = true;
          break;
        case 'lineTo': ctx.lineTo(command[1], command[2]); break;
        case 'bezierCurveTo':
          ctx.bezierCurveTo(command[1], command[2], command[3], command[4], command[5], command[6]);
          break;
        case 'closePath': ctx.closePath(); open = false; break;
      }
    }
    if (closeFigures && open) ctx.closePath();
    return true;
  }
}
