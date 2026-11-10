// Wingdings, Wingdings 2, Wingdings 3 and Webdings font codes → Unicode.
//
// These fonts use a symbol cmap (OpenType platform 3 encoding 0), so a code is
// stored either bare (0x20..0xFF) or F0-shifted (U+F020..U+F0FF). Unicode 6.0
// and 7.0 encoded their whole repertoire; the correspondence is the source
// reference of that encoding:
//   Michel Suignard, "Status of encoding of Wingdings and Webdings Symbols",
//   ISO/IEC JTC1/SC2/WG2 N4384 = UTC L2/12-368 (2012-11-06), section 7
//   ("Source reference", pp. 68-71), cross-checked against the section 6
//   tables (pp. 20-67): https://www.unicode.org/wg2/docs/n4384.pdf
// Every target is an assigned Unicode 7.0 character (UnicodeData.txt 7.0.0).
// Only the code → code-point correspondences (facts) are reproduced here; no
// text, glyph image or font data from the document or the fonts is copied.
// The document gives no outbound licence of its own; it is published on
// unicode.org under the Unicode Terms of Use.
//
// N4384 notes: Wingdings 0x26 is U+1F56E BOOK in sections 6 and 7 (section 3
// says U+1F4D6; sections 6 and 7 agree, so U+1F56E is used). Codes the source
// leaves without a Unicode character keep the authored face: the space 0x20
// and 0x7F in every font, Wingdings 0xFF (Windows logo, not encoded),
// Wingdings 2 0xFA..0xFF and Wingdings 3 0xF1..0xFF (outside the fonts'
// populated range) — 30 codes in all.

/** Wingdings: codes 0x20..0xFF, one code point each ('-' = no target). */
export const WINGDINGS_ENCODING =
  '-\u{1F589}\u2702\u2701\u{1F453}\u{1F56D}\u{1F56E}\u{1F56F}\u{1F57F}\u2706\u{1F582}\u{1F583}\u{1F4EA}\u{1F4EB}\u{1F4EC}\u{1F4ED}'
  + '\u{1F5C0}\u{1F5C1}\u{1F5CE}\u{1F5CF}\u{1F5D0}\u{1F5C4}\u23F3\u{1F5AE}\u{1F5B0}\u{1F5B2}\u{1F5B3}\u{1F5B4}\u{1F5AB}\u{1F5AC}\u2707\u270D'
  + '\u{1F58E}\u270C\u{1F58F}\u{1F44D}\u{1F44E}\u261C\u261E\u261D\u{1F597}\u{1F590}\u263A\u{1F610}\u2639\u{1F4A3}\u{1F571}\u{1F3F3}'
  + '\u{1F3F1}\u2708\u263C\u{1F322}\u2744\u{1F546}\u271E\u{1F548}\u2720\u2721\u262A\u262F\u{1F549}\u2638\u2648\u2649'
  + '\u264A\u264B\u264C\u264D\u264E\u264F\u2650\u2651\u2652\u2653\u{1F670}\u{1F675}\u26AB\u{1F53E}\u25FC\u{1F78F}'
  + '\u{1F790}\u2751\u2752\u{1F79F}\u29EB\u25C6\u2756\u{1F799}\u2327\u2BB9\u2318\u{1F3F5}\u{1F3F6}\u{1F676}\u{1F677}-'
  + '\u{1F10B}\u2780\u2781\u2782\u2783\u2784\u2785\u2786\u2787\u2788\u2789\u{1F10C}\u278A\u278B\u278C\u278D'
  + '\u278E\u278F\u2790\u2791\u2792\u2793\u{1F662}\u{1F660}\u{1F661}\u{1F663}\u{1F65E}\u{1F65C}\u{1F65D}\u{1F65F}\u2219\u2022'
  + '\u2B1D\u2B58\u{1F786}\u{1F788}\u{1F78A}\u{1F78B}\u{1F53F}\u25AA\u{1F78E}\u{1F7C1}\u{1F7C5}\u2605\u{1F7CB}\u{1F7CF}\u{1F7D3}\u{1F7D1}'
  + '\u2BD0\u2316\u2BCE\u2BCF\u2BD1\u272A\u2730\u{1F550}\u{1F551}\u{1F552}\u{1F553}\u{1F554}\u{1F555}\u{1F556}\u{1F557}\u{1F558}'
  + '\u{1F559}\u{1F55A}\u{1F55B}\u2BB0\u2BB1\u2BB2\u2BB3\u2BB4\u2BB5\u2BB6\u2BB7\u{1F66A}\u{1F66B}\u{1F655}\u{1F654}\u{1F657}'
  + '\u{1F656}\u{1F650}\u{1F651}\u{1F652}\u{1F653}\u232B\u2326\u2B98\u2B9A\u2B99\u2B9B\u2B88\u2B8A\u2B89\u2B8B\u{1F868}'
  + '\u{1F86A}\u{1F869}\u{1F86B}\u{1F86C}\u{1F86D}\u{1F86F}\u{1F86E}\u{1F878}\u{1F87A}\u{1F879}\u{1F87B}\u{1F87C}\u{1F87D}\u{1F87F}\u{1F87E}\u21E6'
  + '\u21E8\u21E7\u21E9\u2B04\u21F3\u2B01\u2B00\u2B03\u2B02\u{1F8AC}\u{1F8AD}\u{1F5F6}\u2713\u{1F5F7}\u{1F5F9}-';

/** Wingdings 2: codes 0x20..0xFF, one code point each ('-' = no target). */
export const WINGDINGS_2_ENCODING =
  '-\u{1F58A}\u{1F58B}\u{1F58C}\u{1F58D}\u2704\u2700\u{1F57E}\u{1F57D}\u{1F5C5}\u{1F5C6}\u{1F5C7}\u{1F5C8}\u{1F5C9}\u{1F5CA}\u{1F5CB}'
  + '\u{1F5CC}\u{1F5CD}\u{1F4CB}\u{1F5D1}\u{1F5D4}\u{1F5B5}\u{1F5B6}\u{1F5B7}\u{1F5B8}\u{1F5AD}\u{1F5AF}\u{1F5B1}\u{1F592}\u{1F593}\u{1F598}\u{1F599}'
  + '\u{1F59A}\u{1F59B}\u{1F448}\u{1F449}\u{1F59C}\u{1F59D}\u{1F59E}\u{1F59F}\u{1F5A0}\u{1F5A1}\u{1F446}\u{1F447}\u{1F5A2}\u{1F5A3}\u{1F591}\u{1F5F4}'
  + '\u{1F5F8}\u{1F5F5}\u2611\u2BBD\u2612\u2BBE\u2BBF\u{1F6C7}\u29B8\u{1F671}\u{1F674}\u{1F672}\u{1F673}\u203D\u{1F679}\u{1F67A}'
  + '\u{1F67B}\u{1F666}\u{1F664}\u{1F665}\u{1F667}\u{1F65A}\u{1F658}\u{1F659}\u{1F65B}\u24EA\u2460\u2461\u2462\u2463\u2464\u2465'
  + '\u2466\u2467\u2468\u2469\u24FF\u2776\u2777\u2778\u2779\u277A\u277B\u277C\u277D\u277E\u277F-'
  + '\u2609\u{1F315}\u263D\u263E\u2E3F\u271D\u{1F547}\u{1F55C}\u{1F55D}\u{1F55E}\u{1F55F}\u{1F560}\u{1F561}\u{1F562}\u{1F563}\u{1F564}'
  + '\u{1F565}\u{1F566}\u{1F567}\u{1F668}\u{1F669}\u22C5\u{1F784}\u2981\u25CF\u25CB\u{1F785}\u{1F787}\u{1F789}\u2299\u29BF\u{1F78C}'
  + '\u{1F78D}\u25FE\u25A0\u25A1\u{1F791}\u{1F792}\u{1F793}\u{1F794}\u25A3\u{1F795}\u{1F796}\u{1F797}\u{1F798}\u2B29\u2B25\u25C7'
  + '\u{1F79A}\u25C8\u{1F79B}\u{1F79C}\u{1F79D}\u{1F79E}\u2B2A\u2B27\u25CA\u{1F7A0}\u25D6\u25D7\u2BCA\u2BCB\u2BC0\u2BC1'
  + '\u2B1F\u2BC2\u2B23\u2B22\u2BC3\u2BC4\u{1F7A1}\u{1F7A2}\u{1F7A3}\u{1F7A4}\u{1F7A5}\u{1F7A6}\u{1F7A7}\u{1F7A8}\u{1F7A9}\u{1F7AA}'
  + '\u{1F7AB}\u{1F7AC}\u{1F7AD}\u{1F7AE}\u{1F7AF}\u{1F7B0}\u{1F7B1}\u{1F7B2}\u{1F7B3}\u{1F7B4}\u{1F7B5}\u{1F7B6}\u{1F7B7}\u{1F7B8}\u{1F7B9}\u{1F7BA}'
  + '\u{1F7BB}\u{1F7BC}\u{1F7BD}\u{1F7BE}\u{1F7BF}\u{1F7C0}\u{1F7C2}\u{1F7C4}\u{1F7C6}\u{1F7C9}\u{1F7CA}\u2736\u{1F7CC}\u{1F7CE}\u{1F7D0}\u{1F7D2}'
  + '\u2739\u{1F7C3}\u{1F7C7}\u272F\u{1F7CD}\u{1F7D4}\u2BCC\u2BCD\u203B\u2042------';

/** Wingdings 3: codes 0x20..0xFF, one code point each ('-' = no target). */
export const WINGDINGS_3_ENCODING =
  '-\u2B60\u2B62\u2B61\u2B63\u2B66\u2B67\u2B69\u2B68\u2B70\u2B72\u2B71\u2B73\u2B76\u2B78\u2B7B'
  + '\u2B7D\u2B64\u2B65\u2B6A\u2B6C\u2B6B\u2B6D\u2B4D\u2BA0\u2BA1\u2BA2\u2BA3\u2BA4\u2BA5\u2BA6\u2BA7'
  + '\u2B90\u2B91\u2B92\u2B93\u2B80\u2B83\u2B7E\u2B7F\u2B84\u2B86\u2B85\u2B87\u2B8F\u2B8D\u2B8E\u2B8C'
  + '\u2B6E\u2B6F\u238B\u2324\u2303\u2325\u2423\u237D\u21EA\u2BB8\u{1F8A0}\u{1F8A1}\u{1F8A2}\u{1F8A3}\u{1F8A4}\u{1F8A5}'
  + '\u{1F8A6}\u{1F8A7}\u{1F8A8}\u{1F8A9}\u{1F8AA}\u{1F8AB}\u{1F850}\u{1F852}\u{1F851}\u{1F853}\u{1F854}\u{1F855}\u{1F857}\u{1F856}\u{1F858}\u{1F859}'
  + '\u25B2\u25BC\u25B3\u25BD\u25C0\u25B6\u25C1\u25B7\u25E3\u25E2\u25E4\u25E5\u{1F780}\u{1F782}\u{1F781}-'
  + '\u{1F783}\u2BC5\u2BC6\u2BC7\u2BC8\u2B9C\u2B9E\u2B9D\u2B9F\u{1F810}\u{1F812}\u{1F811}\u{1F813}\u{1F814}\u{1F816}\u{1F815}'
  + '\u{1F817}\u{1F818}\u{1F81A}\u{1F819}\u{1F81B}\u{1F81C}\u{1F81E}\u{1F81D}\u{1F81F}\u{1F800}\u{1F802}\u{1F801}\u{1F803}\u{1F804}\u{1F806}\u{1F805}'
  + '\u{1F807}\u{1F808}\u{1F80A}\u{1F809}\u{1F80B}\u{1F820}\u{1F822}\u{1F824}\u{1F826}\u{1F828}\u{1F82A}\u{1F82C}\u{1F89C}\u{1F89D}\u{1F89E}\u{1F89F}'
  + '\u{1F82E}\u{1F830}\u{1F832}\u{1F834}\u{1F836}\u{1F838}\u{1F83A}\u{1F839}\u{1F83B}\u{1F898}\u{1F89A}\u{1F899}\u{1F89B}\u{1F83C}\u{1F83E}\u{1F83D}'
  + '\u{1F83F}\u{1F840}\u{1F842}\u{1F841}\u{1F843}\u{1F844}\u{1F846}\u{1F845}\u{1F847}\u2BA8\u2BA9\u2BAA\u2BAB\u2BAC\u2BAD\u2BAE'
  + '\u2BAF\u{1F860}\u{1F862}\u{1F861}\u{1F863}\u{1F864}\u{1F865}\u{1F867}\u{1F866}\u{1F870}\u{1F872}\u{1F871}\u{1F873}\u{1F874}\u{1F875}\u{1F877}'
  + '\u{1F876}\u{1F880}\u{1F882}\u{1F881}\u{1F883}\u{1F884}\u{1F885}\u{1F887}\u{1F886}\u{1F890}\u{1F892}\u{1F891}\u{1F893}\u{1F894}\u{1F896}\u{1F895}'
  + '\u{1F897}---------------';

/** Webdings: codes 0x20..0xFF, one code point each ('-' = no target). */
export const WEBDINGS_ENCODING =
  '-\u{1F577}\u{1F578}\u{1F572}\u{1F576}\u{1F3C6}\u{1F396}\u{1F587}\u{1F5E8}\u{1F5E9}\u{1F5F0}\u{1F5F1}\u{1F336}\u{1F397}\u{1F67E}\u{1F67C}'
  + '\u{1F5D5}\u{1F5D6}\u{1F5D7}\u23F4\u23F5\u23F6\u23F7\u23EA\u23E9\u23EE\u23ED\u23F8\u23F9\u23FA\u{1F5DA}\u{1F5F3}'
  + '\u{1F6E0}\u{1F3D7}\u{1F3D8}\u{1F3D9}\u{1F3DA}\u{1F3DC}\u{1F3ED}\u{1F3DB}\u{1F3E0}\u{1F3D6}\u{1F3DD}\u{1F6E3}\u{1F50D}\u{1F3D4}\u{1F441}\u{1F442}'
  + '\u{1F3DE}\u{1F3D5}\u{1F6E4}\u{1F3DF}\u{1F6F3}\u{1F56C}\u{1F56B}\u{1F568}\u{1F508}\u{1F394}\u{1F395}\u{1F5EC}\u{1F67D}\u{1F5ED}\u{1F5EA}\u{1F5EB}'
  + '\u2B94\u2714\u{1F6B2}\u2B1C\u{1F6E1}\u{1F4E6}\u{1F6F1}\u2B1B\u{1F691}\u{1F6C8}\u{1F6E9}\u{1F6F0}\u{1F7C8}\u{1F574}\u2B24\u{1F6E5}'
  + '\u{1F694}\u{1F5D8}\u{1F5D9}\u2753\u{1F6F2}\u{1F687}\u{1F68D}\u26F3\u29B8\u2296\u{1F6AD}\u{1F5EE}\u23D0\u{1F5EF}\u{1F5F2}-'
  + '\u{1F6B9}\u{1F6BA}\u{1F6C9}\u{1F6CA}\u{1F6BC}\u{1F47D}\u{1F3CB}\u26F7\u{1F3C2}\u{1F3CC}\u{1F3CA}\u{1F3C4}\u{1F3CD}\u{1F3CE}\u{1F698}\u{1F5E0}'
  + '\u{1F6E2}\u{1F4B0}\u{1F3F7}\u{1F4B3}\u{1F46A}\u{1F5E1}\u{1F5E2}\u{1F5E3}\u272F\u{1F584}\u{1F585}\u{1F583}\u{1F586}\u{1F5B9}\u{1F5BA}\u{1F5BB}'
  + '\u{1F575}\u{1F570}\u{1F5BD}\u{1F5BE}\u{1F4CB}\u{1F5D2}\u{1F5D3}\u{1F56E}\u{1F4DA}\u{1F5DE}\u{1F5DF}\u{1F5C3}\u{1F5C2}\u{1F5BC}\u{1F3AD}\u{1F39C}'
  + '\u{1F398}\u{1F399}\u{1F3A7}\u{1F4BF}\u{1F39E}\u{1F4F7}\u{1F39F}\u{1F3AC}\u{1F4FD}\u{1F4F9}\u{1F4FE}\u{1F4FB}\u{1F39A}\u{1F39B}\u{1F4FA}\u{1F4BB}'
  + '\u{1F5A5}\u{1F5A6}\u{1F5A7}\u{1F579}\u{1F3AE}\u{1F57B}\u{1F57C}\u{1F4DF}\u{1F581}\u{1F580}\u{1F5A8}\u{1F5A9}\u{1F5BF}\u{1F5AA}\u{1F5DC}\u{1F512}'
  + '\u{1F513}\u{1F5DD}\u{1F4E5}\u{1F4E4}\u{1F573}\u{1F323}\u{1F324}\u{1F325}\u{1F326}\u2601\u{1F328}\u{1F327}\u{1F329}\u{1F32A}\u{1F32C}\u{1F32B}'
  + '\u{1F31C}\u{1F321}\u{1F6CB}\u{1F6CF}\u{1F37D}\u{1F378}\u{1F6CE}\u{1F6CD}\u24C5\u267F\u{1F6C6}\u{1F588}\u{1F393}\u{1F5E4}\u{1F5E5}\u{1F5E6}'
  + '\u{1F5E7}\u{1F6EA}\u{1F43F}\u{1F426}\u{1F41F}\u{1F415}\u{1F408}\u{1F66C}\u{1F66E}\u{1F66D}\u{1F66F}\u{1F5FA}\u{1F30D}\u{1F30F}\u{1F30E}\u{1F54A}';
