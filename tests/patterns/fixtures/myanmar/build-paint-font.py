"""Rebuild the OFL Myanmar Canvas fixture from the pinned upstream TTF.

Requires fonttools[woff] 4.60.1. Pass the downloaded TTF as the first argument.
"""
from pathlib import Path
import hashlib
import sys
from fontTools import subset
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

source = Path(sys.argv[1])
assert hashlib.sha256(source.read_bytes()).hexdigest() == "7abbbfbe2514105d7ce94937aee3feb2ba89b73a256c8b77b5866bd9b83e32ec"
font = instantiateVariableFont(TTFont(source), {"wght": 400, "wdth": 100}, inplace=True)
options = subset.Options()
options.layout_features = ["*"]
options.name_IDs = ["*"]
options.name_legacy = True
options.name_languages = ["*"]
subsetter = subset.Subsetter(options=options)
subsetter.populate(unicodes=[0x20, 0x1000, 0x109D, 0xA9E5, 0xAA7B, 0xAA7C, 0xAA7D, 0x25CC])
subsetter.subset(font)
for record in font["name"].names:
    if record.nameID in [1, 3, 4, 6, 16]:
        name = "MyanmarPaintFixture" if record.nameID == 6 else "Myanmar Paint Fixture"
        record.string = name.encode(record.getEncoding())
font.recalcTimestamp = False
font.flavor = "woff2"
font.save(Path(__file__).with_name("paint.woff2"))
