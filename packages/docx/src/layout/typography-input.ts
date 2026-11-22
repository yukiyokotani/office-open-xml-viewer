import type { DocParagraph, DocxTextRun, FieldRun } from '../types.js';
import { snapshotPlainData } from './plain-data.js';

export type TypographyValueStatus = 'missing' | 'invalid' | 'valid';

/** Parser values retain their authored lexical form because a missing required
 * value and a malformed value are diagnostics, not permission to invent an
 * Office-looking fallback. */
export interface TypographyValueInput<T> {
  readonly status: TypographyValueStatus;
  readonly raw: string | null;
  readonly value: T | null;
}

export interface CtBorderTypographyInput {
  readonly val: TypographyValueInput<string>;
  readonly color: TypographyValueInput<string>;
  readonly themeColor: TypographyValueInput<string>;
  readonly themeTint: TypographyValueInput<string>;
  readonly themeShade: TypographyValueInput<string>;
  readonly sizePt: TypographyValueInput<number>;
  readonly spacePt: TypographyValueInput<number>;
  readonly shadow: TypographyValueInput<boolean>;
  readonly frame: TypographyValueInput<boolean>;
}

export interface InternalRunTypographyWire {
  readonly underline?: Readonly<{
    val: TypographyValueInput<string>;
    color: TypographyValueInput<string>;
    themeColor: TypographyValueInput<string>;
    themeTint: TypographyValueInput<string>;
    themeShade: TypographyValueInput<string>;
  }>;
  readonly strike: boolean;
  readonly doubleStrike: boolean;
  readonly caps: boolean;
  readonly smallCaps: boolean;
  readonly colorAuto: boolean;
  readonly verticalAlign: TypographyValueInput<'super' | 'sub'>;
  readonly positionPt: TypographyValueInput<number>;
  readonly snapToGrid: boolean | null;
  readonly characterSpacingPt: number | null;
  readonly characterScale: number | null;
  readonly fitText?: Readonly<{ valTwips: number; id: string | null }>;
  readonly kerningThresholdPt: number | null;
  readonly emphasis: TypographyValueInput<string>;
  readonly languages: Readonly<{ eastAsia: string | null; bidi: string | null }>;
  readonly eastAsianLayout: Readonly<{
    vert: boolean | null;
    vertCompress: boolean | null;
    combine: boolean | null;
    combineBrackets: TypographyValueInput<string>;
  }>;
  readonly border?: CtBorderTypographyInput;
  readonly ruby?: Readonly<{
    align: TypographyValueInput<string>;
    baseFontSizePt: TypographyValueInput<number>;
    raisePt: TypographyValueInput<number>;
    language: TypographyValueInput<string>;
    guideRuns: readonly Readonly<{
      text: string;
      fontFamily: string | null;
      fontSizePt: number | null;
      bold: boolean;
      italic: boolean;
      color: string | null;
      language: string | null;
    }>[];
  }>;
  readonly revision?: Readonly<{
    kind: 'insertion' | 'deletion';
    id: TypographyValueInput<string>;
    author: string | null;
    date: string | null;
  }>;
}

export interface InternalParagraphTypographyWire {
  readonly borders: Readonly<{
    top?: CtBorderTypographyInput;
    right?: CtBorderTypographyInput;
    bottom?: CtBorderTypographyInput;
    left?: CtBorderTypographyInput;
    between?: CtBorderTypographyInput;
    bar?: CtBorderTypographyInput;
  }>;
}

type InternalTypographyRun = (DocxTextRun | FieldRun) & Readonly<{
  __typographyAcquisition?: InternalRunTypographyWire;
}>;

type InternalTypographyParagraph = DocParagraph & Readonly<{
  __paragraphTypographyAcquisition?: InternalParagraphTypographyWire;
}>;

export type RunTypographyAcquisitionInput = InternalRunTypographyWire & Readonly<{
  sourceText: string;
}>;

/** ECMA-376 §§17.3.2 and 17.3.3 define the formatting semantics independently
 * of whether glyph content came from w:t or a field result. Project both arms
 * through this one immutable boundary so FieldRun cannot silently lose axes. */
export function runTypographyAcquisitionInput(
  run: Readonly<DocxTextRun | FieldRun>,
): Readonly<RunTypographyAcquisitionInput> | undefined {
  const wire = (run as Readonly<InternalTypographyRun>).__typographyAcquisition;
  if (wire === undefined) return undefined;
  const sourceText = 'text' in run ? run.text : run.fallbackText;
  return snapshotPlainData({ sourceText, ...wire }, 'DOCX run typography acquisition input');
}

/** CT_PBdr contains `bar` as well as the five historically public edges. The
 * complete set stays private so retained layout can implement §17.3.1.24
 * without changing the stable consumer-facing DocParagraph type. */
export function paragraphTypographyAcquisitionInput(
  paragraph: Readonly<DocParagraph>,
): Readonly<InternalParagraphTypographyWire> | undefined {
  const wire = (paragraph as Readonly<InternalTypographyParagraph>)
    .__paragraphTypographyAcquisition;
  if (wire === undefined) return undefined;
  return snapshotPlainData(wire, 'DOCX paragraph typography acquisition input');
}
