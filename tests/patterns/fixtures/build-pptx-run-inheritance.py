"""Build a minimal OPC deck that exercises the production PPTX parser."""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def shape(i, y, text, field=False, no_fill=False, highlight=False,
          decorated=False, warp=False, gradient=False):
    local = ('<a:ln w="63500"><a:solidFill><a:srgbClr val="008844"/>'
             '</a:solidFill></a:ln>' if decorated else '')
    local += '<a:noFill/>' if no_fill else ''
    if highlight:
        local += '<a:highlight><a:srgbClr val="F5DD22"/></a:highlight>'
    if decorated:
        local += ('<a:uFill><a:gradFill><a:gsLst>'
                  '<a:gs pos="0"><a:srgbClr val="00DDDD"/></a:gs>'
                  '<a:gs pos="100000"><a:srgbClr val="DD00DD"/></a:gs>'
                  '</a:gsLst><a:lin ang="0"/></a:gradFill></a:uFill>' if gradient else
                  '<a:uFill><a:pattFill prst="pct50">'
                  '<a:fgClr><a:srgbClr val="0033EE"/></a:fgClr>'
                  '<a:bgClr><a:srgbClr val="FFDD00"/></a:bgClr>'
                  '</a:pattFill></a:uFill>')
    run = ('<a:fld id="{0FBB9679-570D-4509-9600-E07789CF24BB}" type="datetime">'
           if field else '<a:r>')
    underline_attr = ' u="sng"' if decorated else ''
    run += f'<a:rPr{underline_attr}>{local}</a:rPr><a:t>{text}</a:t>'
    run += '</a:fld>' if field else '</a:r>'
    style = ('<a:lstStyle><a:lvl1pPr><a:defRPr sz="4800"'
             ' kumimoji="1" lang="ja-JP" altLang="en-US" kern="1200"'
             ' cap="none" normalizeH="1" noProof="1" smtClean="0" smtId="17" bmk="sample">'
             '<a:solidFill><a:srgbClr val="D21D54"/></a:solidFill>'
             '<a:latin typeface="Arial" panose="020B0604020202020204"/>'
             '<a:cs typeface="Amiri"/><a:sym typeface="Symbol"/>'
             '<a:hlinkMouseOver r:id="rId4" tooltip="hover"/>'
             '</a:defRPr></a:lvl1pPr></a:lstStyle>')
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Inheritance {i}"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="600000" y="{y}"/>'
            '<a:ext cx="9000000" cy="1300000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0">'
            + ('<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>' if warp else '')
            + '</a:bodyPr>'
            f'{style}<a:p><a:pPr><a:defRPr b="1"/></a:pPr>{run}</a:p></p:txBody></p:sp>')

shapes = (shape(2, 500000, 'RUN INHERITS RED')
          + shape(3, 2300000, 'FIELD INHERITS RED', field=True)
          + shape(4, 4100000, 'INVISIBLE GLYPHS', no_fill=True, highlight=True)
          + shape(5, 5550000, 'OUTLINE AND UNDERLINE', no_fill=True, decorated=True))
def make_slide(shapes):
    return (f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree>'
         '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
         f'<p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>')
slide = make_slide(shapes)
warp_slide = make_slide(
    shape(2, 500000, 'WARP OUTLINE UNDERLINE',
          no_fill=True, decorated=True, warp=True)
    + shape(3, 3500000, 'GRADIENT UNDERLINE',
            no_fill=True, decorated=True, gradient=True)
    + shape(4, 5200000, 'WARP GRADIENT LINE',
            no_fill=True, decorated=True, gradient=True, warp=True))
def break_shape(i, x, break_size):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Break {break_size}"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="500000"/>'
            '<a:ext cx="5000000" cy="2100000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/>'
            '<a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>BEFORE</a:t></a:r>'
            f'<a:br><a:rPr sz="{break_size}"/></a:br>'
            '<a:r><a:rPr sz="2400"/><a:t>AFTER</a:t></a:r>'
            '<a:endParaRPr sz="9600"/></a:p></p:txBody></p:sp>')

def table_cell():
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="600000" y="3600000"/>'
            '<a:ext cx="10500000" cy="1800000"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="10500000"/></a:tblGrid>'
            '<a:tr h="1800000"><a:tc><a:txBody><a:bodyPr/>'
            '<a:lstStyle><a:defPPr><a:defRPr lang="ja-JP" sz="2800">'
            '<a:solidFill><a:srgbClr val="D21D54"/></a:solidFill>'
            '<a:cs typeface="Amiri"/></a:defRPr></a:defPPr></a:lstStyle>'
            '<a:p><a:pPr><a:defRPr b="1"/></a:pPr>'
            '<a:r><a:rPr/><a:t>CELL TEXT</a:t></a:r>'
            '<a:fld id="{0FBB9679-570D-4509-9600-E07789CF24BB}" type="datetime">'
            '<a:rPr/><a:t> FIELD TEXT</a:t></a:fld></a:p></a:txBody><a:tcPr/>'
            '</a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>')

break_table_slide = make_slide(break_shape(2, 600000, 2400)
                               + break_shape(3, 6200000, 7200)
                               + table_cell())
parts = {
    '[Content_Types].xml': f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
    'ppt/presentation.xml': f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="{R}/slide" Target="slides/slide2.xml"/><Relationship Id="rId3" Type="{R}/slide" Target="slides/slide3.xml"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
    'ppt/slides/_rels/slide1.xml.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId4" Type="{R}/hyperlink" Target="https://example.test/hover" TargetMode="External"/></Relationships>',
    'ppt/slides/slide2.xml': warp_slide,
    'ppt/slides/slide3.xml': break_table_slide,
}
out = Path('tests/patterns/fixtures/pptx-run-inheritance.pptx')
with ZipFile(out, 'w', ZIP_DEFLATED) as package:
    for path, content in sorted(parts.items()):
        item = ZipInfo(path, (2020, 1, 1, 0, 0, 0))
        item.compress_type = ZIP_DEFLATED
        package.writestr(item, content)
print(out, out.stat().st_size)
