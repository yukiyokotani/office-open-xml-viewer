import { describe, expect, it } from 'vitest';
import { paintedDocument, PICTURE_PART } from './test-support/painted-glyphs.test-support.js';

describe('a floating picture anchored in the last line of a vertical page', () => {
  // A tbRl section whose 468pt of columns hold a 444pt line and a 25pt one:
  // the last line's ink fits in the 24pt left, so the text ends on page one.
  const exactLine = (twips: number) => `<w:pPr><w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/></w:pPr>`;
  const text = (value: string) => `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${value}</w:t></w:r>`;
  const floating = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
    + 'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
    + '<wp:positionH relativeFrom="character"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    + '<wp:positionV relativeFrom="line"><wp:posOffset>508000</wp:posOffset></wp:positionV>'
    + '<wp:extent cx="381000" cy="508000"/><wp:wrapNone/><wp:docPr id="1" name="Picture"/>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
    + '<pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip r:embed="rIdP"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="508000"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
    + '</wp:anchor></w:drawing></w:r>';
  const pages = async (anchor: string): Promise<number> => (await paintedDocument(
    `<w:p>${exactLine(8880)}${text('filler')}</w:p><w:p>${exactLine(500)}${text('last')}${anchor}</w:p>`,
    { vertical: true, ...PICTURE_PART })).layout.pages.length;

  it('keeps the line on the page it fits on without the picture', async () => {
    expect(await pages('')).toBe(1);
    expect(await pages(floating)).toBe(1);
  });
});
