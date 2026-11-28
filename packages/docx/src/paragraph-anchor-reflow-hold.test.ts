import { beforeAll, describe, expect, it } from 'vitest';
import {
  anchoredRectangle,
  layOutPages,
  loadParser,
  P,
  paragraph,
  type Direction,
  type Stories,
} from './test-support/vertical-section-upright.test-support.js';

// The anchor reflow of a paragraph (AnchorReflowMoves, layout/paragraph.ts):
// a wrapped drawing positioned from its character follows text that its own
// exclusion moves, so its positions can repeat without settling. The reflow
// then ends in the hold pass, in whichever story the paragraph lies, and the
// document lays out the same way every time.

beforeAll(loadParser);

/** A 140-character paragraph whose last run anchors a 200 x 50 pt
 * square-wrapped rectangle `x` pt from its character and `y` pt down its
 * paragraph, followed by a 40-character paragraph. */
const unsettled = (x: number, y: number) =>
  paragraph('あ'.repeat(140)).replace(
    '</w:p>', `${anchoredRectangle(x, y, 200, 50, 'paragraph', 'character')}</w:p>`,
  ) + paragraph('あ'.repeat(40));

const cellTable = (content: string) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
  + '<w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>'
  + `${content}</w:tc></w:tr></w:tbl>`;

/** A paragraph anchoring a 300 x 400 pt text box, not wrapped, whose story
 * is `content`; `vert` is its body's text direction (a:bodyPr@vert). */
const textBox = (content: string, vert?: string) =>
  '<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
  + ' distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/>'
  + '<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
  + '<wp:positionV relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionV>'
  + '<wp:extent cx="3810000" cy="5080000"/><wp:wrapNone/><wp:docPr id="90" name="Box"/>'
  + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  + '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
  + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr>'
  + '<a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="5080000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
  + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
  + `<wps:bodyPr${vert ? ` vert="${vert}"` : ''} lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr>`
  + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';

describe('anchor reflow of a paragraph whose drawing keeps moving its own text', () => {
  const cases: ReadonlyArray<readonly [string, readonly string[], Stories | undefined]> = [
    ['the body', [unsettled(-44, 39)], undefined],
    ['a table cell', [P, cellTable(unsettled(59, 61))], undefined],
    ['a header', [P], { header: unsettled(-44, 39) }],
    ['a footer', [P], { footer: unsettled(-44, 39) }],
    ['a table cell of a header', [P], { header: cellTable(unsettled(59, 61)) }],
    ['a text box', [textBox(unsettled(-44, 39)), P], undefined],
    ['a vertical text box', [textBox(unsettled(-44, 39), 'eaVert'), P], undefined],
  ];
  for (const direction of ['tbRl', 'tbLrV'] as const satisfies readonly Direction[]) {
    it.each(cases)(`ends in a ${direction} section, in %s`, (_story, body, stories) => {
      const layOut = () => layOutPages(direction, body, 1, 360, undefined, undefined, stories);
      let pages: ReturnType<typeof layOut> = [];
      expect(() => { pages = layOut(); }).not.toThrow();
      expect(pages.length).toBeGreaterThan(0);
      expect(layOut()).toEqual(pages);
    });
  }
});
