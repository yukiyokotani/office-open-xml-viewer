/** Numeric geometry, constructed without a renderer or paint callback. */
export interface GeometryPathSink {
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  quadraticCurveTo(x1: number, y1: number, x: number, y: number): void;
  bezierCurveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void;
  closePath(): void;
  rect(x: number, y: number, w: number, h: number): void;
  roundRect(x: number, y: number, w: number, h: number, radius: number | readonly Readonly<{ x: number; y: number }>[]): void;
  arc(x: number, y: number, radius: number, start: number, end: number, ccw?: boolean): void;
  ellipse(x: number, y: number, rx: number, ry: number, rotation: number, start: number, end: number, ccw?: boolean): void;
}
export type GeometryPathOperation = keyof GeometryPathSink;
export interface GeometryPathCommand {
  readonly op: GeometryPathOperation;
  /** arc/ellipse's final direction is represented by numeric 0/1. */
  readonly args: readonly number[];
}
export type GeometryPath = readonly GeometryPathCommand[];
export class GeometryWorkBudgetError extends RangeError {}

export function createGeometryPath(charge: (units: number) => void = () => {}): {
  readonly sink: GeometryPathSink;
  readonly commands: GeometryPathCommand[];
} {
  const commands: GeometryPathCommand[] = [];
  const add = (op: GeometryPathOperation, args: number[]) => {
    charge(1 + args.length);
    if (!args.every(Number.isFinite)) throw new RangeError('Non-finite resolved DrawingML path');
    if ((op === 'arc' && args[2] < 0) || (op === 'ellipse' && (args[2] < 0 || args[3] < 0))) {
      throw new RangeError('Negative resolved DrawingML ellipse radius');
    }
    commands.push(Object.freeze({ op, args: Object.freeze(args) }));
  };
  return { commands, sink: {
    moveTo: (x, y) => add('moveTo', [x, y]),
    lineTo: (x, y) => add('lineTo', [x, y]),
    quadraticCurveTo: (x1, y1, x, y) => add('quadraticCurveTo', [x1, y1, x, y]),
    bezierCurveTo: (x1, y1, x2, y2, x, y) => add('bezierCurveTo', [x1, y1, x2, y2, x, y]),
    closePath: () => add('closePath', []),
    rect: (x, y, w, h) => add('rect', [x, y, w, h]),
    roundRect: (x, y, w, h, radius) => add('roundRect', typeof radius === 'number' ? [x, y, w, h, 0, radius] : [x, y, w, h, 1, ...radius.flatMap(r => [r.x, r.y])]),
    arc: (x, y, radius, start, end, ccw = false) => add('arc', [x, y, radius, start, end, Number(ccw)]),
    ellipse: (x, y, rx, ry, rotation, start, end, ccw = false) => add('ellipse', [x, y, rx, ry, rotation, start, end, Number(ccw)]),
  } };
}

/** Validate transported numeric records before canonical consumption. */
export function assertGeometryPath(path: GeometryPath, charge: (units: number) => void = () => {}): void {
  const arities: Partial<Record<GeometryPathOperation, number>> = {
    moveTo: 2, lineTo: 2, quadraticCurveTo: 4, bezierCurveTo: 6, closePath: 0, rect: 4, arc: 6, ellipse: 8,
  };
  if (!Array.isArray(path)) throw new RangeError('Invalid retained geometry path');
  for (const command of path) {
    const a = command.args;
    if (!Array.isArray(a)) throw new RangeError('Invalid retained geometry arguments');
    charge(1 + a.length);
    if (!a.every(Number.isFinite)) throw new RangeError('Non-finite retained geometry path');
    if (command.op === 'roundRect') {
      if (!((a[4] === 0 && a.length === 6) || (a[4] === 1 && a.length >= 7 && (a.length - 5) % 2 === 0))) throw new RangeError('Invalid retained round rectangle');
    } else if (!(command.op in arities) || a.length !== arities[command.op as GeometryPathOperation]) throw new RangeError('Invalid retained geometry operation');
    if ((command.op === 'arc' && (a[2] < 0 || (a[5] !== 0 && a[5] !== 1)))
      || (command.op === 'ellipse' && (a[2] < 0 || a[3] < 0 || (a[7] !== 0 && a[7] !== 1)))) throw new RangeError('Invalid retained ellipse');
  }
}

/** Replay already-resolved coordinates. Translating their origin numerically
 * keeps the caller's brush/gradient CTM unchanged. */
export function appendGeometryPath(sink: GeometryPathSink, path: GeometryPath, x = 0, y = 0): void {
  for (const { op, args: a } of path) {
    switch (op) {
      case 'moveTo': sink.moveTo(x + a[0], y + a[1]); break;
      case 'lineTo': sink.lineTo(x + a[0], y + a[1]); break;
      case 'quadraticCurveTo': sink.quadraticCurveTo(x + a[0], y + a[1], x + a[2], y + a[3]); break;
      case 'bezierCurveTo': sink.bezierCurveTo(x + a[0], y + a[1], x + a[2], y + a[3], x + a[4], y + a[5]); break;
      case 'closePath': sink.closePath(); break;
      case 'rect': sink.rect(x + a[0], y + a[1], a[2], a[3]); break;
      case 'roundRect': {
        const radii = a[4] === 0 ? a[5] : Array.from({ length: (a.length - 5) / 2 }, (_, i) => ({ x: a[5 + i * 2], y: a[6 + i * 2] }));
        sink.roundRect(x + a[0], y + a[1], a[2], a[3], radii); break;
      }
      case 'arc': sink.arc(x + a[0], y + a[1], a[2], a[3], a[4], Boolean(a[5])); break;
      case 'ellipse': sink.ellipse(x + a[0], y + a[1], a[2], a[3], a[4], a[5], a[6], Boolean(a[7])); break;
    }
  }
}
