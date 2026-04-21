/**
 * ECMA-376 DrawingML preset-geometry formula evaluator (§19.5.31.3).
 *
 * Each gdLst entry is a postfix expression whose operands are
 * previously-defined guide names, built-in names (w, h, ss, cd4, …), adjust
 * names (adj1..adj8), or numeric literals. We evaluate them in declaration
 * order into a name→value map.
 *
 * Angle-bearing operators (sin/cos/tan/at2/sat2/cat2) treat their angle
 * arguments as 60 000-ths of a degree — the OOXML convention shared by
 * the `cd4`/`3cd4`/… built-ins and by `<arcTo>`'s stAng/swAng attributes.
 */

// 60 000-ths of a degree per full revolution (2π radians).
const CD = 21600000;
const DEG60K_TO_RAD = (Math.PI * 2) / CD;

export interface EvalInputs {
  /** Path-space width (usually the shape's canvas-px width). */
  w: number;
  /** Path-space height. */
  h: number;
  /** Adjust values (index 0 = adj/adj1, 1 = adj2, …). `null` / missing → use preset default. */
  adj: (number | null | undefined)[];
}

export interface Evaluator {
  /** Get a value by name. Throws if unknown. */
  v(name: string): number;
  /** Evaluate a raw formula expression. */
  fmla(expr: string): number;
  /** Evaluate a token that is either a guide name, built-in, or literal number. */
  resolve(token: string): number;
}

export function createEvaluator(
  inputs: EvalInputs,
  adjDefaults: [string, string][],
  gdList: [string, string][],
): Evaluator {
  const { w, h, adj } = inputs;
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);

  const env: Record<string, number> = Object.create(null);

  // ── Built-ins (ECMA-376 §20.1.9.2) ──────────────────────────────────────
  Object.assign(env, {
    w, h,
    l: 0, t: 0, r: w, b: h,
    hc: w / 2, vc: h / 2,
    wd2: w / 2, wd3: w / 3, wd4: w / 4, wd5: w / 5, wd6: w / 6,
    wd8: w / 8, wd10: w / 10, wd12: w / 12, wd16: w / 16, wd32: w / 32,
    hd2: h / 2, hd3: h / 3, hd4: h / 4, hd5: h / 5, hd6: h / 6,
    hd8: h / 8, hd10: h / 10, hd12: h / 12, hd16: h / 16, hd32: h / 32,
    ss, ssd2: ss / 2, ssd4: ss / 4, ssd6: ss / 6, ssd8: ss / 8,
    ssd16: ss / 16, ssd32: ss / 32,
    ls, lsd2: ls / 2, lsd4: ls / 4, lsd6: ls / 6, lsd8: ls / 8,
    lsd16: ls / 16, lsd32: ls / 32,
    cd: CD,
    cd2: CD / 2, cd4: CD / 4, cd8: CD / 8,
    '3cd4': (3 * CD) / 4, '3cd8': (3 * CD) / 8,
    '5cd8': (5 * CD) / 8, '7cd8': (7 * CD) / 8,
  });

  // ── Adjust resolution: caller-supplied (non-null) overrides preset default ─
  adjDefaults.forEach(([name, fmla], i) => {
    const supplied = adj[i];
    env[name] = typeof supplied === 'number' ? supplied : evaluateFormula(fmla);
    // Many shapes name the first adjust "adj" but reference it as "adj1".
    if (name === 'adj')  env.adj1 = env.adj;
    if (name === 'adj1') env.adj  = env.adj1;
  });

  // ── Guide list (evaluated in declaration order) ─────────────────────────
  for (const [name, fmla] of gdList) {
    env[name] = evaluateFormula(fmla);
  }

  return {
    v: (name) => {
      if (name in env) return env[name];
      throw new Error(`preset-shape: unknown name "${name}"`);
    },
    fmla: evaluateFormula,
    resolve,
  };

  function resolve(token: string): number {
    if (token in env) return env[token];
    const n = Number(token);
    if (Number.isFinite(n)) return n;
    throw new Error(`preset-shape: cannot resolve "${token}"`);
  }

  function evaluateFormula(expr: string): number {
    const parts = expr.trim().split(/\s+/);
    const op = parts[0];
    const args = parts.slice(1).map(resolve);
    return applyOp(op, args, expr);
  }

  function applyOp(op: string, a: number[], original: string): number {
    switch (op) {
      case 'val': return a[0];
      case '*/':  return (a[0] * a[1]) / a[2];
      case '+-':  return a[0] + a[1] - a[2];
      case '+/':  return (a[0] + a[1]) / a[2];
      case '?:':  return a[0] > 0 ? a[1] : a[2];
      case 'abs': return Math.abs(a[0]);
      case 'min': return Math.min(a[0], a[1]);
      case 'max': return Math.max(a[0], a[1]);
      case 'pin': return a[1] < a[0] ? a[0] : a[1] > a[2] ? a[2] : a[1];
      case 'sqrt': return Math.sqrt(Math.max(0, a[0]));
      case 'mod': return Math.sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
      case 'sin': return a[0] * Math.sin(a[1] * DEG60K_TO_RAD);
      case 'cos': return a[0] * Math.cos(a[1] * DEG60K_TO_RAD);
      case 'tan': return a[0] * Math.tan(a[1] * DEG60K_TO_RAD);
      // at2 returns an angle in 60 000-ths of a degree.
      case 'at2': return Math.atan2(a[1], a[0]) / DEG60K_TO_RAD;
      // cat2 y z: x * cos(atan2(z, y)) — used when projecting a point on an ellipse.
      case 'cat2': return a[0] * Math.cos(Math.atan2(a[2], a[1]));
      case 'sat2': return a[0] * Math.sin(Math.atan2(a[2], a[1]));
      default:
        throw new Error(`preset-shape: unknown operator "${op}" in "${original}"`);
    }
  }
}
