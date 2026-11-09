"""Generate the synthetic font-source test fonts (design §12.1).

Writes these fonts under tests/fixtures/fonts/:

- ooxml-fixture.ttf: printable ASCII, U+00A0 and U+00B7;
- ooxml-fixture-jp.ttf: U+3000-U+3005, U+3041-U+3046 and U+4E00-U+4E05;
- ooxml-fixture-gap.ttf: the Latin set with a non-zero line gap (hhea and
  OS/2 typo lineGap 300) and a declared Latin 1 code page, so line-metric tests
  can tell a face's own normal line from its gap-less font box;
- ooxml-fixture-gap.woff / .woff2: the same font as WOFF 1.0 and WOFF 2.0;
- ooxml-fixture-gap-var.ttf: a variable version (wght 100-900) whose Black
  master has ascender 900 and lineGap 500, so MVAR varies hasc/hlgp (and the
  OS/2 typo/win values) with the weight.

Every glyph is a synthetic rectangle (no outline is derived from any existing
font, so the files carry no third-party copyright). Advances are fixed per
code point and heights vary with the code point, so a page painted with the
fixture differs visibly from any host fallback and is deterministic across
platforms. Run with:

    uv run --no-project --with fonttools==4.67.0 --with brotli==1.1.0 python scripts/generate-fixture-font.py

The generated files are committed so browser tests never need Python.
"""

from pathlib import Path

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'tests/fixtures/fonts'
UPM = 1000
ASCENT = 800
DESCENT = 200


def rectangle(width: int, height: int, inset: int):
    pen = TTGlyphPen(None)
    pen.moveTo((inset, 0))
    pen.lineTo((inset, height))
    pen.lineTo((width - inset, height))
    pen.lineTo((width - inset, 0))
    pen.closePath()
    return pen.glyph()


def empty():
    return TTGlyphPen(None).glyph()


def build(path: Path | None, family: str, code_points: list[int], advance: int, line_gap: int = 0,
          ascent: int = ASCENT, weight: int = 400, style_name: str = 'Regular') -> FontBuilder:
    glyph_order = ['.notdef']
    glyphs = {'.notdef': rectangle(advance, ASCENT, 40)}
    metrics = {'.notdef': (advance, 40)}
    cmap = {}
    for cp in code_points:
        name = f'uni{cp:04X}'
        glyph_order.append(name)
        if cp in (0x20, 0xA0, 0x3000):
            glyphs[name] = empty()
            metrics[name] = (advance, 0)
        else:
            # Height encodes the code point, so different text paints differently.
            height = 200 + (cp * 37) % 560
            glyphs[name] = rectangle(advance, height, 60)
            metrics[name] = (advance, 60)
        cmap[cp] = name
    builder = FontBuilder(UPM, isTTF=True)
    builder.setupGlyphOrder(glyph_order)
    builder.setupCharacterMap(cmap)
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics(metrics)
    builder.setupHorizontalHeader(ascent=ascent, descent=-DESCENT, lineGap=line_gap)
    builder.setupNameTable({
        'familyName': family,
        'styleName': style_name,
        'uniqueFontIdentifier': f'{family} {style_name} synthetic test font',
        'fullName': f'{family} {style_name}',
        'psName': family.replace(' ', '') + '-' + style_name,
        'version': 'Version 1.000',
        'licenseDescription': 'Synthetic test font of the office-open-xml-viewer project (MIT).',
    })
    builder.setupOS2(
        sTypoAscender=ascent, sTypoDescender=-DESCENT, sTypoLineGap=line_gap,
        usWinAscent=ascent, usWinDescent=DESCENT, achVendID='OOXV',
        fsType=0, usWeightClass=weight,
        **({'ulCodePageRange1': 1} if line_gap else {}),
    )
    builder.setupPost()
    builder.setupHead(unitsPerEm=UPM, created=0, modified=0)
    if path is not None:
        path.parent.mkdir(parents=True, exist_ok=True)
        builder.save(str(path))
    return builder


def build_variable(path: Path, family: str, code_points: list[int]) -> None:
    from fontTools.designspaceLib import AxisDescriptor, DesignSpaceDocument, SourceDescriptor
    from fontTools.varLib import build as build_varfont

    doc = DesignSpaceDocument()
    axis = AxisDescriptor()
    axis.tag, axis.name, axis.minimum, axis.default, axis.maximum = 'wght', 'Weight', 100, 400, 900
    doc.addAxis(axis)
    masters = [(100, 'Thin', 800, 300), (400, 'Regular', 800, 300), (900, 'Black', 900, 500)]
    for weight, style, ascent, gap in masters:
        source = SourceDescriptor()
        source.font = build(None, family, code_points, 600, line_gap=gap, ascent=ascent,
                            weight=weight, style_name=style).font
        source.location = {'Weight': weight}
        source.styleName = style
        doc.addSource(source)
    font, _, _ = build_varfont(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    font.save(str(path))


def save_flavors(path: Path) -> None:
    from fontTools.ttLib import TTFont

    for flavor in ('woff', 'woff2'):
        font = TTFont(str(path))
        font.flavor = flavor
        font.save(str(path.with_suffix('.' + flavor)))


def main() -> None:
    latin = list(range(0x20, 0x7F)) + [0xA0, 0xB7]
    jp = list(range(0x3000, 0x3006)) + list(range(0x3041, 0x3047)) + list(range(0x4E00, 0x4E06))
    build(OUT / 'ooxml-fixture.ttf', 'Ooxml Fixture', latin, 600)
    build(OUT / 'ooxml-fixture-jp.ttf', 'Ooxml Fixture JP', jp, 1000)
    build(OUT / 'ooxml-fixture-gap.ttf', 'Ooxml Fixture Gap', latin, 600, line_gap=300)
    save_flavors(OUT / 'ooxml-fixture-gap.ttf')
    build_variable(OUT / 'ooxml-fixture-gap-var.ttf', 'Ooxml Fixture Gap Var', latin)


if __name__ == '__main__':
    main()
