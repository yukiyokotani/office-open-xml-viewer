import { parseA1 } from './a1.js';
import type { DefinedName } from './types.js';

/** Concrete destination for a SpreadsheetML internal hyperlink. */
export interface XlsxInternalHyperlinkDestination {
  readonly sheetIndex: number;
  /** First cell of the referenced cell or range, retaining authored `$` markers. */
  readonly cellRef: string;
}

function normalizeExpression(value: string): string {
  let normalized = value.trim();
  if (normalized.startsWith('#')) normalized = normalized.slice(1).trim();
  if (normalized.startsWith('=')) normalized = normalized.slice(1).trim();
  return normalized;
}

/** Find the last `!` outside a quoted sheet name. Doubled apostrophes inside a
 * quoted name are escaped content, not quote delimiters. */
function sheetSeparator(value: string): number {
  let quoted = false;
  let separator = -1;
  for (let i = 0; i < value.length; i++) {
    if (value[i] === "'") {
      if (quoted && value[i + 1] === "'") {
        i++;
      } else {
        quoted = !quoted;
      }
    } else if (value[i] === '!' && !quoted) {
      separator = i;
    }
  }
  return quoted ? -1 : separator;
}

function unquoteSheetName(value: string): string {
  const name = value.trim();
  return name.length >= 2 && name.startsWith("'") && name.endsWith("'")
    ? name.slice(1, -1).replace(/''/g, "'")
    : name;
}

function directCellDestination(
  expression: string,
  currentSheetIndex: number,
  sheetNames: readonly string[],
): XlsxInternalHyperlinkDestination | null {
  const separator = sheetSeparator(expression);
  const sheetName = separator >= 0
    ? unquoteSheetName(expression.slice(0, separator))
    : undefined;
  const reference = (separator >= 0 ? expression.slice(separator + 1) : expression)
    .split(':', 1)[0]
    ?.trim() ?? '';
  if (!parseA1(reference)) return null;

  let sheetIndex = currentSheetIndex;
  if (sheetName !== undefined) {
    const folded = sheetName.toLocaleLowerCase('en-US');
    sheetIndex = sheetNames.findIndex(
      (candidate) => candidate.toLocaleLowerCase('en-US') === folded,
    );
    if (sheetIndex < 0) return null;
  }
  if (sheetIndex < 0 || sheetIndex >= sheetNames.length) return null;
  return { sheetIndex, cellRef: reference };
}

/**
 * Resolve an internal SpreadsheetML hyperlink (`hyperlink@location`,
 * ECMA-376 §18.3.1.47) to a sheet and the first cell of its target.
 *
 * A location can be a direct cell/range reference or a defined name. Defined
 * names can themselves refer to another in-scope name (§18.2.5), so resolution
 * is finite and cycle-safe. The worksheet parser supplies workbook-global plus
 * current-sheet-local definitions, locals last; later definitions win, so a
 * local name shadows its global twin as in the formula/conditional-format
 * maps and chart references.
 */
export function resolveXlsxInternalHyperlink(
  location: string,
  currentSheetIndex: number,
  sheetNames: readonly string[],
  definedNames: readonly DefinedName[],
): XlsxInternalHyperlinkDestination | null {
  const names = new Map<string, DefinedName>();
  for (const definition of definedNames) {
    names.set(definition.name.toLocaleLowerCase('en-US'), definition);
  }

  let expression = normalizeExpression(location);
  const seen = new Set<string>();
  for (;;) {
    const direct = directCellDestination(expression, currentSheetIndex, sheetNames);
    if (direct) return direct;

    const key = expression.toLocaleLowerCase('en-US');
    if (!key || seen.has(key)) return null;
    seen.add(key);
    const definition = names.get(key);
    if (!definition) return null;
    expression = normalizeExpression(definition.formula);
  }
}
