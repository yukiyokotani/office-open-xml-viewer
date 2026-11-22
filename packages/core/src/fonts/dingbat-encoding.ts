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
//
// Storage: each table lists codes 0x20..0xFF in order as comma-separated
// tokens; an empty token is a code without a target, any other token is the
// target's code point minus the previous target's (signed, base 36; the
// first is relative to 0). Most targets are astral pictographs in a few
// blocks, so deltas are short; as literal characters the four tables cost
// almost twice the gzip bytes, and they ship twice (page and worker bundle).
// symbol-font.ts `dingbatTable` decodes them; symbol-font.test.ts pins
// spot values and the 866-code total.

/** wingdings: codes 0x20..0xFF. */
export const WINGDINGS_ENCODING =
  ',2r2h,-2jd3,-1,2j4i,7u,1,1,g,-2jcp,2jcs,1,-49,1,1,1,5v,1,d,1,1,-c,-2k0h,'
  + '2jzv,2,2,1,1,-9,1,-2jdx,6,2jcx,-2jcy,2jcz,-8y,1,-2jaq,2,-1,2jju,-7,-2jiu,'
  + '2jme,-2jmf,2jca,5q,-am,-2,-2j1l,-5o,2j1i,-2iu6,2j9e,-2jag,2jai,-2jag,1,-6v,'
  + '5,2jh6,-2jgx,g,1,1,1,1,1,1,1,1,1,1,1,2jod,5,-2jm2,2jdf,-2jia,2jyr,1,-2jpb,1,'
  + '2jpp,-2j78,-th,b4,2jpf,-2kj6,1oy,-1pd,2jtp,1,hs,1,,-12k,-2idn,1,1,1,1,1,1,1,'
  + '1,1,2idf,-2ide,1,1,1,1,1,1,1,1,1,2jf3,-2,1,2,-5,-2,1,2,-2khy,-dz,263,1n,'
  + '2iwe,2,2,1,-gc,-2jkl,2k10,1f,4,-2k00,2k06,4,4,-2,-2iv5,-1q2,1q0,1,2,-x3,6,'
  + '2ja8,1,1,1,1,1,1,1,1,1,1,1,-2iej,1,1,1,1,1,1,1,2ilv,1,-m,-1,3,-1,-6,1,1,1,'
  + '-2ka0,-5,1o2,2,-1,2,-j,2,-1,2,2j19,2,-1,2,1,1,2,-1,a,2,-1,2,1,1,2,-1,-2kyg,'
  + '2,-1,2,1sr,-1sh,1se,-1,3,-1,2j6y,1,-jb,-2jfn,2jfo,2,';

/** wingdings 2: codes 0x20..0xFF. */
export const WINGDINGS_2_ENCODING =
  ',2r2i,1,1,1,-2jd5,-4,2jcu,-1,20,1,1,1,1,1,1,1,1,-76,7a,3,-v,1,1,1,-b,2,2,-v,'
  + '1,5,1,1,1,-9f,1,9f,1,1,1,1,1,-9n,1,9n,1,-i,2r,4,-3,-2jms,14c,-14b,14c,1,'
  + '2io8,-2j2n,2j09,3,-2,1,-2kvq,2kvw,1,1,-l,-2,1,2,-d,-2,1,2,-2jxt,-3u,1,1,1,1,'
  + '1,1,1,1,1,46,hj,1,1,1,1,1,1,1,1,1,,-ae,2j2k,-2j14,1,1kx,-1eq,2jai,l,1,1,1,1,'
  + '1,1,1,1,1,1,1,75,1,-2kdg,2klb,-2j9f,-qa,-4,2jzu,2,2,-2kmo,1eu,2j7x,1,-2jyn,'
  + '-2m,1,2k1c,1,1,1,-2k1d,2k1e,1,1,1,-2iy7,-4,-126,2k0j,-2k0i,2k0j,1,1,1,-2iyc,'
  + '-3,-125,2k0m,-2k0a,1,16b,1,-b,1,-4i,4j,-4f,-1,4h,1,2iu5,1,1,1,1,1,1,1,1,1,1,'
  + '1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,2,2,2,3,1,-2jro,2jrq,2,2,2,-2jrt,'
  + '2jre,4,-2jrs,2jry,7,-2ivc,1,-2aa,7,,,,,,';

/** wingdings 3: codes 0x20..0xFF. */
export const WINGDINGS_3_ENCODING =
  ',8kg,2,-1,2,3,1,2,-1,8,2,-1,2,3,2,3,2,-p,1,5,2,-1,2,-w,2b,1,1,1,1,1,1,1,-n,'
  + '1,1,1,-j,3,-5,1,5,2,-1,2,8,-2,1,-2,-u,1,-1k4,-2v,-x,y,72,-4m,-b7,1xq,2j1k,1,'
  + '1,1,1,1,1,1,1,1,1,1,-2j,2,-1,2,1,1,2,-1,2,1,-2k6f,a,-9,a,3,-a,b,-a,18,-1,2,'
  + '1,2jyz,2,-1,,2,-2ita,1,1,1,-18,2,-1,2,2iy9,2,-1,2,1,2,-1,2,1,2,-1,2,1,2,-1,'
  + '2,-v,2,-1,2,1,2,-1,2,1,2,-1,2,l,2,2,2,2,2,2,34,1,1,1,-35,2,2,2,2,2,2,-1,2,'
  + '2l,2,-1,2,-2n,2,-1,2,1,2,-1,2,1,2,-1,2,-2izj,1,1,1,1,1,1,1,2j01,2,-1,2,1,1,'
  + '2,-1,a,2,-1,2,1,1,2,-1,a,2,-1,2,1,1,2,-1,a,2,-1,2,1,2,-1,2,,,,,,,,,,,,,,,';

/** webdings: codes 0x20..0xFF. */
export const WEBDINGS_ENCODING =
  ',2r1z,1,-6,4,-c0,-1c,dt,2p,1,7,1,-jf,2p,kn,-2,-4n,1,1,-2k0z,1,1,1,-d,-1,5,'
  + '-1,b,1,1,2k0w,p,6l,-ll,1,1,1,2,h,-i,5,-a,7,li,-d2,-8p,31,1,-2s,-9,lr,-lh,lw,'
  + '-av,-1,-3,-2o,-ac,1,gn,41,-40,-3,1,-2ijb,-w0,2jku,-2is6,2ith,-e3,ej,-2ity,'
  + '2ira,1j,x,7,60,-gk,-2ij4,2itd,-29,-58,1,-2jd2,2jkv,-2z,6,-2jkq,jp,-1eq,2kgn,'
  + '-5b,-2k2m,2k2n,3,,5j,1,f,1,-e,-fz,-4y,-2j10,2j0r,a,-2,-6,9,1,ju,-54,76,-fm,'
  + '-55,58,-21,af,1,1,-2jec,2jbp,1,-2,3,1f,1,1,-1y,-5,25,1,-6r,7b,1,-2t,-44,78,'
  + '1,-s,-1,-6,-en,-h,-4,1,e,7s,-81,9l,-9k,d,9d,-4,5,-3,-9t,1,9r,-1r,6i,1,1,-1a,'
  + '-cr,ct,1,-4d,4i,-1,14,1,m,-l,1e,-5m,1,5m,-6w,-1,3z,-gg,1,1,1,-2j39,2j3b,-1,'
  + '2,1,2,-1,-f,5,q2,4,-nm,-5,nq,-1,-2k20,ca,2jpj,-8u,-dx,gh,1,1,1,77,-iz,-p,-7,'
  + '-a,-d,h0,2,-1,2,-39,-kt,2,-1,fw';
