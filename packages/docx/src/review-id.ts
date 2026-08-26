/** ECMA-376 ST_DecimalNumber is xsd:integer: XML Schema whitespace collapses
 * and lexical variants such as `+1`, `01`, and `1` share one value. */
export function decimalReviewIdKey(id: string | undefined): string | undefined {
  if (id === undefined) return undefined;
  // xsd:whiteSpace="collapse" is defined over XML Schema whitespace only
  // (#x9, #xA, #xD, #x20). JavaScript trim() would also remove NBSP and other
  // Unicode separators, incorrectly accepting a non-integer lexeme.
  const collapsed = id.replace(/^[\t\n\r ]+|[\t\n\r ]+$/g, '');
  if (!/^[+-]?\d+$/.test(collapsed)) return undefined;
  try {
    return BigInt(collapsed).toString();
  } catch {
    return undefined;
  }
}
