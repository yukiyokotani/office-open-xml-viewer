#!/usr/bin/env python3
"""Generate metadata-only reference font metrics for deterministic layout fallbacks.

The generated profiles are reference facts, not proof of the font face selected by
Canvas, the operating system, or Office. No outlines, glyph IDs, or per-glyph
advances are written to the output. Glyph coverage records cmap presence only.
OS/2 xAvgCharWidth is a scalar font metric, not a shaped text advance.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import plistlib
import re
import subprocess
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from fontTools.ttLib import TTCollection, TTFont


FONT_SUFFIXES = {".otc", ".otf", ".ttc", ".ttf"}
OFFICE_ROOT = Path("/Applications/Microsoft Word.app/Contents/Resources")
MACOS_ROOT = Path("/System/Library/Fonts/Supplemental")
MACOS_PRIMARY_ROOT = Path("/System/Library/Fonts")
DATA_OUTPUT = Path("packages/core/src/fonts/reference-font-metrics-data.json")
PROVENANCE_OUTPUT = Path("scripts/reference-font-metrics-provenance.json")

# Bounded symbol domain from the #1653/#1689 slot sweeps, not all Unicode:
# Latin-1 symbols; General Punctuation; Letterlike/Number Forms, Arrows, Math,
# Misc Technical; Enclosed, Box/Block/Geometric, Misc Symbols and Dingbats.
# Presence is a font fact, independent of which language routes it to ea.
SYMBOL_COVERAGE_RANGES = ((0x00A0, 0x00FF), (0x2000, 0x206F),
                          (0x2100, 0x23FF), (0x2460, 0x27BF))

# Script-bearing CJK, punctuation and width forms used by the empty-ea path.
# A bounded scalar domain, including supplementary Han/kana; outside it the
# runtime reports unknown rather than inferring coverage from the family.
CJK_COVERAGE_RANGES = ((0x1100, 0x11FF), (0x2E80, 0x33FF), (0x3400, 0x9FFF),
                       (0xA960, 0xA97F), (0xAC00, 0xD7FF), (0xF900, 0xFAFF),
                       (0xFE30, 0xFE4F), (0xFF00, 0xFFEF), (0x16FE0, 0x18DFF),
                       (0x1AFF0, 0x1B16F), (0x20000, 0x323AF))


def bounded_coverage(ranges: list[list[int]] | None, domain: tuple[tuple[int, int], ...]) -> list[int] | None:
    """Project production all-map presence/possible facts into fixed domains."""
    if ranges is None:
        return None
    result: list[int] = []
    for lo, hi in ranges:
        for start, end in domain:
            a, b = max(lo, start), min(hi, end)
            if a <= b:
                result.extend((a, b))
    return result


def packed_cjk_bitmap(ranges: tuple[int, ...]) -> str:
    """PackBits bitmap in domain order; trailing zero bytes are implicit.

    Dense CJK repertoires have many one-glyph holes. A bitmap avoids thousands
    of numeric endpoints; PackBits suppresses long absent/full byte runs without
    needing a runtime compression dependency. 128 is the PackBits packet limit.
    """
    bits = bytearray((sum(end - start + 1 for start, end in CJK_COVERAGE_RANGES) + 7) // 8)
    offset = 0
    for start, end in CJK_COVERAGE_RANGES:
        for lo, hi in zip(ranges[::2], ranges[1::2]):
            for code in range(max(start, lo), min(end, hi) + 1):
                index = offset + code - start
                bits[index >> 3] |= 1 << (index & 7)
        offset += end - start + 1
    while bits and bits[-1] == 0:
        bits.pop()
    packed = bytearray()
    index = 0
    while index < len(bits):
        run = 1
        while index + run < len(bits) and run < 128 and bits[index + run] == bits[index]:
            run += 1
        if run >= 3:
            packed.extend((257 - run, bits[index]))
            index += run
            continue
        start = index
        index += run
        while index < len(bits) and index - start < 128:
            run = 1
            while index + run < len(bits) and run < 128 and bits[index + run] == bits[index]:
                run += 1
            if run >= 3:
                break
            index += min(run, 128 - (index - start))
        packed.append(index - start - 1)
        packed.extend(bits[start:index])
    return base64.b64encode(packed).decode("ascii")


@dataclass(frozen=True)
class Source:
    id: str
    roots: tuple[tuple[str, Path], ...]
    version: str


def normalized_name(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def unique_names(font: TTFont, ids: set[int]) -> list[str]:
    values: dict[str, str] = {}
    if "name" not in font:
        return []
    for record in font["name"].names:
        if record.nameID not in ids:
            continue
        try:
            value = " ".join(record.toUnicode().split())
        except (UnicodeDecodeError, AttributeError):
            continue
        key = normalized_name(value)
        if value and key not in values:
            values[key] = value
    return [values[key] for key in sorted(values)]


def preferred_name(font: TTFont, ids: Iterable[int]) -> str | None:
    if "name" not in font:
        return None
    records = font["name"].names
    for name_id in ids:
        candidates = []
        for record in records:
            if record.nameID != name_id:
                continue
            try:
                value = " ".join(record.toUnicode().split())
            except (UnicodeDecodeError, AttributeError):
                continue
            if not value:
                continue
            english = record.langID in {0, 0x0409} or record.langID == 0x8000
            windows = record.platformID == 3
            candidates.append((not english, not windows, record.langID, value))
        if candidates:
            return min(candidates)[3]
    return None


def integer(table: Any, field: str) -> int | None:
    value = getattr(table, field, None)
    return int(value) if value is not None else None


def slug(value: str) -> str:
    compact = re.sub(r"[^a-z0-9]+", "-", normalized_name(value)).strip("-")
    return compact[:48] or "unnamed"


def read_plist_version(path: Path, keys: tuple[str, ...]) -> str:
    try:
        payload = plistlib.loads(path.read_bytes())
    except (OSError, plistlib.InvalidFileException):
        return "unknown"
    return " (".join(str(payload[key]) for key in keys if payload.get(key)) + (")" if all(payload.get(key) for key in keys) and len(keys) > 1 else "")


def font_files(source: Source) -> list[tuple[str, Path]]:
    files = []
    for label, root in source.roots:
        for path in root.iterdir():
            if path.is_file() and path.suffix.lower() in FONT_SUFFIXES:
                files.append((f"{label}/{path.name}", path))
    return sorted(files, key=lambda item: (item[0].casefold(), item[0]))


def face_profile(font: TTFont, source_id: str) -> dict[str, Any] | None:
    if "head" not in font or "hhea" not in font:
        return None
    head = font["head"]
    hhea = font["hhea"]
    os2 = font["OS/2"] if "OS/2" in font else None
    family_aliases = unique_names(font, {1, 16})
    family = preferred_name(font, (16, 1)) or (family_aliases[0] if family_aliases else "Unknown")
    postscript = preferred_name(font, (6,))
    aliases_by_key = {normalized_name(value): value for value in family_aliases + unique_names(font, {6})}
    aliases = [aliases_by_key[key] for key in sorted(aliases_by_key)]
    fs_selection = integer(os2, "fsSelection") if os2 else 0
    os2_version = integer(os2, "version") if os2 else None
    italic = bool((fs_selection or 0) & 0x01 or integer(head, "macStyle") & 0x02)
    profile = {
        "source": source_id,
        "family": family,
        "aliases": aliases,
        "weight": integer(os2, "usWeightClass") if os2 else 400,
        "style": "italic" if italic else "normal",
        "unitsPerEm": integer(head, "unitsPerEm"),
        "xAvgCharWidth": integer(os2, "xAvgCharWidth") if os2 else None,
        "hhea": [integer(hhea, "ascent"), integer(hhea, "descent"), integer(hhea, "lineGap")],
    }
    provenance_code_page_range1 = (
        integer(os2, "ulCodePageRange1")
        if os2 is not None and (os2_version or 0) >= 1
        else None
    )
    # Word for Mac's controlled auto-line tests classify bits 17–20 as the
    # Far-East allocation class. Keep only that derived class at runtime;
    # absent OS/2 code-page data is unknown, not evidence of the Latin class.
    profile["farEastCodePage"] = (
        None if provenance_code_page_range1 is None
        else bool(provenance_code_page_range1 & 0x001E0000)
    )
    # Excel's DrawingML shape-text line box follows the OS/2 usWin extent
    # (issue #1604 controls: Yu Gothic, whose hhea and usWin boxes differ).
    # Null means the face has no OS/2 table.
    profile["win"] = (
        None if os2 is None
        else [integer(os2, "usWinAscent"), integer(os2, "usWinDescent")]
    )
    # PANOSE family kind and serif style (OS/2 panose bytes 1-2). PowerPoint
    # picks its application-default East Asian face from the Latin face's
    # serif class (issue #1627 controls: serif styles 2-10 took MS Mincho,
    # 11-15 MS Gothic). Null means the face has no OS/2 table.
    profile["panose"] = (
        None if os2 is None
        else [int(os2.panose.bFamilyType), int(os2.panose.bSerifStyle)]
    )
    # Assigned below from the production presence/possible projection. The
    # Far-East base-CJK branch requires known presence or known absence; the
    # preferred FontTools cmap is not authority when selectable maps disagree.
    profile["cjkUnifiedIdeographs"] = None
    profile["symbolCoverage"] = None
    profile["cjkCoverage"] = None
    # Honor fsSelection USE_TYPO_METRICS (bit 7), like the resource parser.
    # Although introduced in OS/2 v4, installed v3 resources also set it;
    # #1689 exported symbol resources confirm their typo+gap line metrics.
    # Requiring v4 silently discards the font's declared selection.
    if os2 is not None and (fs_selection or 0) & 0x80:
        profile["typoMetrics"] = [integer(os2, "sTypoAscender"), integer(os2, "sTypoDescender"),
                                  integer(os2, "sTypoLineGap")]
    # Keep provenance identifiers stable when a source fact stops shipping in
    # the runtime profile. The identity still covers that raw OS/2 value.
    identity_profile = {
        **{key: value for key, value in profile.items()
           if key not in {"farEastCodePage", "win", "typoMetrics", "panose", "cjkUnifiedIdeographs", "symbolCoverage", "cjkCoverage"}},
        "os2": None if os2 is None else {
            "codePageRange1": provenance_code_page_range1,
        },
    }
    canonical = json.dumps(identity_profile, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    profile_id = f"{source_id}:{slug(postscript or family)}:{hashlib.sha256(canonical.encode()).hexdigest()[:12]}"
    profile["_provenanceId"] = profile_id
    profile["_provenanceMetrics"] = None if os2 is None else {
        "typo": [integer(os2, "sTypoAscender"), integer(os2, "sTypoDescender"), integer(os2, "sTypoLineGap")],
        "win": [integer(os2, "usWinAscent"), integer(os2, "usWinDescent")],
        "useTypoMetrics": bool((fs_selection or 0) & 0x80),
    }
    profile["_provenanceCodePageRange1"] = provenance_code_page_range1
    return profile


def packed_ranges(ranges: Iterable[int]) -> str:
    """Lossless unsigned varint (gap, length) pairs; no repertoire reduction."""
    packed = bytearray()
    previous = -1
    def append(value: int) -> None:
        while value >= 128:
            packed.append((value & 127) | 128)
            value >>= 7
        packed.append(value)
    values = list(ranges)
    for lo, hi in zip(values[::2], values[1::2]):
        append(lo - previous - 1)
        append(hi - lo)
        previous = hi
    return base64.b64encode(packed).decode("ascii")


def split_catalogue(profiles: list[dict[str, Any]], coverages: dict[str, Any], routes: Any) -> tuple[dict, dict, dict]:
    """Preserve physical row order even when projected metrics become equal.

    Only PPTX rendering owns per-cut support; shared metrics and parse/preload
    routing must not import it. A generated identity binds every companion.
    Empty, false, missing and unknown remain distinct in fixed tuples.
    """
    resource_fields = {"symbolCoverage", "symbolPossibleCoverage", "cjkCoverage", "cjkPossibleCoverage", "supportFacts", "cjkUnifiedIdeographs"}
    metrics = [{k: v for k, v in p.items() if k not in resource_fields} for p in profiles]
    generation = hashlib.sha256(json.dumps(profiles, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    templates, template_ids, erasures, erasure_ids, rows = [], {}, [], {}, []
    ternary = lambda value: 0 if value is None else 2 if value else 1
    dispositions = [None, "identity", "active-open-type", "profile-inactive-major"]
    reasons = [None, "glyph-domain", "unsupported", "malformed", "budget", "cycle"]
    for profile in profiles:
        support = profile.get("supportFacts")
        template_id = None
        if support is not None:
            erased = support.get("erasureSafeRanges")
            erasure_id = None
            if erased is not None:
                packed = packed_ranges(n for pair in erased for n in pair)
                if packed not in erasure_ids:
                    erasure_ids[packed] = len(erasures); erasures.append(packed)
                erasure_id = erasure_ids[packed]
            template = [ternary(support.get(k)) for k in ("nonzeroPreserved", "missingIsolated", "noErasure", "anyIndic3ScriptPresent")]
            template += [support.get("gsubLookupCount"), dispositions.index(support.get("gsubDisposition")), reasons.index(support.get("reason")), erasure_id, int(support.get("profile") is not None)]
            key = tuple(template)
            if key not in template_ids:
                template_ids[key] = len(templates); templates.append(template)
            template_id = template_ids[key]
        rows.append([profile.get(k) for k in ("symbolCoverage", "symbolPossibleCoverage", "cjkCoverage", "cjkPossibleCoverage")]
                    + [template_id, support.get("glyphCount") if support else None, ternary(profile.get("cjkUnifiedIdeographs"))])
    # Six unsigned 16-bit fields plus one ternary byte per physical row.
    # Index zero denotes missing; actual glyph counts stay unshifted. Fixed
    # records avoid per-row emitted JavaScript syntax without changing identity.
    encoded_rows = bytearray()
    for row in rows:
        for index, value in enumerate(row[:6]):
            encoded = 0 if value is None else value + (index < 5)
            if not 0 <= encoded <= 65535:
                raise ValueError("catalogue field exceeds fixed record domain")
            encoded_rows.extend(encoded.to_bytes(2, "big"))
        encoded_rows.append(row[6])
    identity = {"schemaVersion": 1, "generation": generation}
    return ({**identity, "profiles": metrics},
            {**identity, "supportSchema": "ot-definedness-1", "supportProfile": "canonical-static-v1", "unicode": "17.0.0",
             "symbolCoverageRanges": SYMBOL_COVERAGE_RANGES, "cjkCoverageRanges": CJK_COVERAGE_RANGES,
             "symbolCoverages": [packed_ranges(r) for r in coverages["symbolCoverages"]], "cjkCoverages": coverages["cjkCoverages"],
             "templates": templates, "erasureRanges": erasures, "rowCount": len(rows), "rowsEncoded": base64.b64encode(encoded_rows).decode("ascii")},
            {**identity, "routes": routes})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--office-root", type=Path, default=OFFICE_ROOT)
    parser.add_argument("--macos-root", type=Path, default=MACOS_ROOT)
    parser.add_argument("--macos-primary-root", type=Path, default=MACOS_PRIMARY_ROOT)
    parser.add_argument("--data-output", type=Path, default=DATA_OUTPUT)
    parser.add_argument("--provenance-output", type=Path, default=PROVENANCE_OUTPUT)
    parser.add_argument("--resource-output", type=Path, default=Path("packages/pptx/src/font-resource-catalogue-data.json"))
    parser.add_argument("--route-output", type=Path, default=Path("packages/pptx/src/font-route-data.json"))
    args = parser.parse_args()

    office_version = read_plist_version(args.office_root.parent / "Info.plist", ("CFBundleShortVersionString", "CFBundleVersion"))
    system_version = read_plist_version(Path("/System/Library/CoreServices/SystemVersion.plist"), ("ProductVersion", "ProductBuildVersion"))
    sources = (
        Source("office-mac", (("DFonts", args.office_root / "DFonts"), ("OtherFonts", args.office_root / "OtherFonts")), office_version),
        Source("macos-system", (("Primary", args.macos_primary_root),), system_version),
        Source("macos-supplemental", (("Supplemental", args.macos_root),), system_version),
    )

    support_process = subprocess.Popen(['node', 'scripts/generate-font-support-facts.mjs'],
                                       stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    profiles_by_canonical: dict[str, dict[str, Any]] = {}
    provenance: list[dict[str, Any]] = []
    exclusions: list[dict[str, Any]] = []
    for source in sources:
        for relative_path, path in font_files(source):
            file_hash = hashlib.sha256(path.read_bytes()).hexdigest()
            collection = TTCollection(path, lazy=True) if path.suffix.lower() in {".ttc", ".otc"} else None
            fonts = collection.fonts if collection else [TTFont(path, lazy=True)]
            try:
                for face_index, font in enumerate(fonts):
                    if "fvar" in font:
                        exclusions.append({"source": source.id, "file": relative_path, "faceIndex": face_index, "reason": "variable-font"})
                        continue
                    profile = face_profile(font, source.id)
                    if profile is None:
                        exclusions.append({"source": source.id, "file": relative_path, "faceIndex": face_index, "reason": "missing-head-or-hhea"})
                        continue
                    support_process.stdin.write(json.dumps({"path": str(path), "faceIndex": face_index}) + "\n")
                    support_process.stdin.flush()
                    resource = json.loads(support_process.stdout.readline())
                    profile["supportFacts"] = resource["supportFacts"]
                    known, possible = resource["unicodeRanges"], resource["unicodePossibleRanges"]
                    basic = lambda ranges: any(lo <= 0x9FFF and hi >= 0x4E00 for lo, hi in ranges)
                    profile["cjkUnifiedIdeographs"] = (True if known is not None and basic(known)
                        else False if possible is not None and not basic(possible) else None)

                    # All eligible-map intersection proves presence; their union
                    # bounds possible presence. A preferred cmap cannot certify a
                    # resource when browser-selectable maps disagree.
                    for field, domain in [("symbolCoverage", SYMBOL_COVERAGE_RANGES), ("cjkCoverage", CJK_COVERAGE_RANGES)]:
                        profile[field] = bounded_coverage(resource["unicodeRanges"], domain)
                        profile[field.replace("Coverage", "PossibleCoverage")] = bounded_coverage(resource["unicodePossibleRanges"], domain)
                    profile_id = profile.pop("_provenanceId")
                    provenance_metrics = profile.pop("_provenanceMetrics")
                    provenance_code_page_range1 = profile.pop("_provenanceCodePageRange1")
                    canonical = json.dumps(profile, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                    profiles_by_canonical.setdefault(canonical, profile)
                    provenance.append({
                        "profileId": profile_id,
                        "source": source.id,
                        "file": relative_path,
                        "faceIndex": face_index,
                        "version": preferred_name(font, (5,)),
                        "sha256": file_hash,
                        "rawOs2VerticalMetrics": provenance_metrics,
                        "rawOs2CodePageRange1": provenance_code_page_range1,
                    })
            finally:
                if collection:
                    collection.close()
                else:
                    fonts[0].close()

    source_order = {source.id: index for index, source in enumerate(sources)}
    profiles = sorted(
        profiles_by_canonical.values(),
        key=lambda item: (
            source_order[item["source"]], normalized_name(item["family"]),
            item["weight"], item["style"],
            json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
        ),
    )
    provenance.sort(key=lambda item: (item["source"], item["file"].casefold(), item["file"], item["faceIndex"]))
    exclusions.sort(key=lambda item: (item["source"], item["file"].casefold(), item["file"], item["faceIndex"]))
    # Intern identical repertoires across cuts and sources, including empty
    # cmaps. Runtime shares the frozen arrays; no per-glyph cache is needed.
    coverages = {}
    for field in ("symbolCoverage", "cjkCoverage"):
        possible_field = field.replace("Coverage", "PossibleCoverage")
        repertoires = sorted({tuple(p[k]) for p in profiles for k in (field, possible_field) if p[k] is not None})
        coverage_ids = {r: i for i, r in enumerate(repertoires)}
        for profile in profiles:
            for k in (field, possible_field):
                coverage = profile[k]
                profile[k] = None if coverage is None else coverage_ids[tuple(coverage)]
        coverages[field + "s"] = ([packed_cjk_bitmap(r) for r in repertoires]
                                   if field == "cjkCoverage" else repertoires)
    # The same JavaScript alias/source reducer serves generation and tests.
    support_process.stdin.write(json.dumps({"profiles": profiles}) + "\n")
    support_process.stdin.flush()
    routes = json.loads(support_process.stdout.readline())
    support_process.stdin.close()
    if support_process.wait() != 0:
        raise RuntimeError('Font support certificate generation failed')
    data, resources, routing = split_catalogue(profiles, coverages, routes)
    manifest = {
        "schemaVersion": 2,
        "notice": "Development provenance only; this file is not imported by the runtime metrics lookup.",
        "sources": [{"id": source.id, "version": source.version} for source in sources],
        "faces": provenance,
        "excludedFaces": exclusions,
    }
    args.data_output.parent.mkdir(parents=True, exist_ok=True)
    args.provenance_output.parent.mkdir(parents=True, exist_ok=True)
    for output, payload in [(args.data_output, data), (args.resource_output, resources), (args.route_output, routing)]:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
    args.provenance_output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")


if __name__ == "__main__":
    main()
