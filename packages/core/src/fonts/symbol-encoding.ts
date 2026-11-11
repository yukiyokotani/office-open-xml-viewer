/*! Symbol encoding table derived from "Adobe Symbol Encoding to Unicode"
 * (table version 1.0, 2011 July 12), published by Unicode, Inc. at
 * https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
 * Its notice, which travels with this data (also in THIRD_PARTY_NOTICES.md):
 *
 * Copyright (c) 1991-2011 Unicode, Inc. All Rights reserved.
 *
 * This file is provided as-is by Unicode, Inc. (The Unicode Consortium). No
 * claims are made as to fitness for any particular purpose. No warranties of
 * any kind are expressed or implied. The recipient agrees to determine
 * applicability of information provided. If this file has been provided on
 * magnetic media by Unicode, Inc., the sole remedy for any claim will be
 * exchange of defective media within 90 days of receipt.
 *
 * Unicode, Inc. hereby grants the right to freely use the information
 * supplied in this file in the creation of products supporting the
 * Unicode Standard, and to make copies of this file in any form for
 * internal or external distribution as long as this notice remains
 * attached.
 */
// The Symbol font's own 8-bit encoding → Unicode, shared by run text
// (symbol-font.ts) and GDI metafile text (metafile-symbol-encoding.ts).
//
// SOURCE: the table above (field 1 = Unicode, field 2 = Symbol code); only
// its mapping data is reproduced here, with the deviations listed below.
// The Microsoft Symbol font (symbol.ttf) carries the same repertoire through a
// symbol cmap (platform 3 encoding 0) at U+F020..U+F0FF, so a code is stored
// either bare (0x20..0xFF) or F0-shifted:
//   https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
//
// Deviations from the raw table:
//   - Where the table lists two targets for one code (0x20, 0x44, 0x57, 0x6D,
//     0xA4) the Greek/primary entry is used.
//   - 0xA0 is the euro sign in both the 2011 table and Microsoft's Symbol.
//   - Phi is the one deliberate exchange. The table predates Unicode 3.0, which
//     swapped the representative glyphs of U+03C6 (now the open, looped phi)
//     and U+03D5 (now the stroked phi). Office evidence: a Windows Word PDF of
//     a MathType equation drawing Symbol 0x6A ("phi1") shows the open, looped
//     phi, and Times New Roman (the Windows serif fallback) draws U+03C6 looped
//     and U+03D5 stroked. Mapping 0x6A to U+03D5 would therefore paint the
//     other glyph, so 0x6A maps to U+03C6 and its stroked counterpart 0x66 to
//     U+03D5.
//   - The table's 29 Corporate-Use targets (U+F6D9..U+F6DB, U+F8E5..U+F8FE)
//     are not standard characters; see SYMBOL_CORPORATE_USE_UNICODE.
//   - 0xF0 is unassigned in both the table and Microsoft's Symbol (Apple's Mac
//     Symbol puts its logo there, at Apple's private U+F8FF), and 0x7F..0x9F
//     are unassigned; those codes have no Unicode target.
// Storage: codes 0x20..0xFE in order as comma-separated tokens; an empty
// token is an unassigned code, any other the target code point minus the
// previous target's (signed, base 36). symbol-font.test.ts pins the decoded
// table (entry count and an order-sensitive hash of the literal table).
export const ADOBE_SYMBOL_UNICODE: Readonly<Record<number, number>> = decodeDeltaTable(0x20,
  'w,1,6ov,-6ot,6ow,-6ou,1,6p1,-6oz,1,6pa,-6p8,1,6p2,-6p0,1,1,1,1,1,1,1,1,1,'
  + '1,1,1,1,1,1,1,1,6py,-62c,1,l,-j,1,h,-j,4,2,1k,-1j,1,1,1,2,1,-8,9,2,1,1,t,'
  + '-p,-b,a,-i,-mz,6op,-6on,6rs,-6rq,,nm,1,l,-j,1,w,-y,4,2,d,-c,1,1,1,2,1,-8,'
  + '9,2,1,1,h,-d,-b,a,-i,-mz,1,1,6nz,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,-b4,'
  + '-5p6,5ls,fm,-f4,d6,-6fg,79t,3,-1,-5,-y4,-4,1,1,1,-6hv,1,682,fm,-6mm,6km,'
  + '-r,-dc,-65n,6ll,1,-p,-f6,,,b3,-3k,-10,b,-4,an,-2,-40,10,1,2h,4,-3,-2,4,'
  + '-3i,1,n,-p,,,,8,b,4r,-6qh,6m3,1,-2c,-4,1,1,1,s7,-ip,,,,-7s,,,,,,,,,,,,7t,'
  + '-73,6t,,1,,,,,,,,,');

/** Decode a table of base-36 code point deltas starting at `firstCode`. */
export function decodeDeltaTable(firstCode: number, encoding: string): Readonly<Record<number, number>> {
  const out: Record<number, number> = {};
  let code = firstCode;
  let cp = 0;
  for (const token of encoding.split(',')) {
    if (token) {
      cp += Number.parseInt(token, 36);
      out[code] = cp;
    }
    code++;
  }
  return Object.freeze(out);
}

// The codes symbol.txt assigns to Corporate-Use code points, mapped to the
// standard characters Unicode later encoded for exactly these glyphs, judged
// by the table's PostScript glyph names:
//   - 0xD2..0xD4 registerserif/copyrightserif/trademarkserif and 0xE2..0xE4
//     their sans counterparts: the serif/sans distinction is a style variant
//     of U+00AE, U+00A9 and U+2122.
//   - 0xBD arrowvertex / 0xBE arrowhorizex: U+23D0 VERTICAL LINE EXTENSION and
//     U+23AF HORIZONTAL LINE EXTENSION.
//   - 0xE6..0xEF, 0xF6..0xFE parenthesis/bracket/brace pieces and 0xF4
//     integralex: the U+239B..U+23AE bracket-piece block, added in Unicode 3.2
//     for these Symbol-font pieces (name for name: parenlefttp = LEFT
//     PARENTHESIS UPPER HOOK, braceleftmid = LEFT CURLY BRACKET MIDDLE PIECE).
//   - 0x60 radicalex: the bar that extends a radical sign over its operand,
//     drawn at the top of the em box. Unicode has no radical-extension piece;
//     U+203E OVERLINE is the standard character of that form, and a Windows
//     PowerPoint PDF of Symbol U+F060 shows exactly such a top bar.
export const SYMBOL_CORPORATE_USE_UNICODE: Readonly<Record<number, number>> = {
  0x60: 0x203e,
  0xbd: 0x23d0, 0xbe: 0x23af,
  0xd2: 0x00ae, 0xd3: 0x00a9, 0xd4: 0x2122,
  0xe2: 0x00ae, 0xe3: 0x00a9, 0xe4: 0x2122,
  0xe6: 0x239b, 0xe7: 0x239c, 0xe8: 0x239d,
  0xe9: 0x23a1, 0xea: 0x23a2, 0xeb: 0x23a3,
  0xec: 0x23a7, 0xed: 0x23a8, 0xee: 0x23a9, 0xef: 0x23aa,
  0xf4: 0x23ae,
  0xf6: 0x239e, 0xf7: 0x239f, 0xf8: 0x23a0,
  0xf9: 0x23a4, 0xfa: 0x23a5, 0xfb: 0x23a6,
  0xfc: 0x23ab, 0xfd: 0x23ac, 0xfe: 0x23ad,
};
