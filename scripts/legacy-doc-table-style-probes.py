#!/usr/bin/env python3
"""Generate and passively mutate bounded Word table-style probe documents.

The DOCX source contains eight tables backed by seven named table styles.
``prepare`` accepts the Word 97-2003 DOC produced from that source,
derives all CFB/FKP/STSH offsets, removes only flattened direct color and
alignment/font/size properties, and emits reviewed same-length
connected/negative counterfactuals. ``mutate`` additionally accepts a bounded
``legacy-doc-table-style-recipe/v1`` JSON object: full TAPX/PAPX/CHPX payload
hex by logical style name, optional TIstd connection, and TTlp flags by T01-T08.
PAPX recipe payloads include the embedded istd. The tool never starts Office
or changes filesystem modes.

This is a compatibility probe, not a general DOC editor.  It intentionally
accepts only this fixed design and delegates PAPX/Data provenance checks to
``legacy-doc-papx-probes.py``.
"""

from __future__ import annotations

import argparse
from bisect import bisect_right
from dataclasses import dataclass
from hashlib import sha256
import importlib.util
from io import BytesIO
import json
import os
from pathlib import Path
import stat
import sys
from typing import Mapping
import zipfile


SCHEMA = "legacy-doc-table-style-probes/v1"
RECIPE_SCHEMA = "legacy-doc-table-style-recipe/v1"
MAX_JSON_BYTES = 8 * 1024 * 1024
MAX_REWRITE_BYTES = 1024 * 1024
MAX_PROBE_CFB_BYTES = 16 * 1024 * 1024
MAX_PROBE_STORY_UNITS = 4096
TABLE_COUNT = 8
DIRECT_COLOR_TABLE = 6

STYLE_DEFINITIONS = (
    ("Base", None, "FF0000", "left", 20, "Arial", "Arial"),
    ("Child", "Base", "0000FF", "center", 28, "Times New Roman", "Times New Roman"),
    ("Empty", "Base", None, None, None, None, None),
    ("Grand", "Child", "008000", "right", 36, "Courier New", "Courier New"),
    ("ReverseBase", None, "008000", "right", 36, "Courier New", "Courier New"),
    ("ReverseChild", "ReverseBase", "FF0000", "left", 20, "Arial", "Arial"),
    ("Plain", None, "000000", "left", 20, "Arial", "Arial"),
)
TABLE_STYLES = (
    "Base",
    "Child",
    "Empty",
    "Grand",
    "ReverseChild",
    "Base",
    "Plain",
    "Plain",
)
ROW_STYLES = (
    "Base", "Child", "Empty", "Grand", "ReverseChild", "Base",
    "Plain", "Plain", "Plain", "Plain",
)
ROW_TABLES = ("T01", "T02", "T03", "T04", "T05", "T06",
              "T07", "T07", "T07", "T08")
MARKERS = tuple(
    [f"T0{table}R1C1 Ω" for table in range(1, 7)]
    + [f"T07R{row}C{column} Ω" for row in range(1, 4) for column in range(1, 4)]
    + ["T08R1C1 Ω"]
)

CI_CO = 0x2A42
C_CV = 0x6870
P_JC_80 = 0x2403
P_JC = 0x2461
TI_STD = 0x563A
P_ITAP = 0x6649
PF_TTP = 0x2417
SUPPORTED_CHPX = frozenset((CI_CO, C_CV, 0x4A43, 0x4A4F, 0x4A51))
SUPPORTED_PAPX = frozenset((P_JC_80, P_JC))

FC_STSHF = 0xA2
LCB_STSHF = 0xA6
FC_PLCF_BTE_CHPX = 0xFA
LCB_PLCF_BTE_CHPX = 0xFE


class ProbeError(ValueError):
    """A source does not match the fixed probe or a mutation is unsafe."""


_PAPX = None


def _papx_module():
    global _PAPX
    if _PAPX is not None:
        return _PAPX
    path = Path(__file__).with_name("legacy-doc-papx-probes.py")
    spec = importlib.util.spec_from_file_location("legacy_doc_papx_probes", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    _PAPX = module
    return _PAPX


def _load_probe_document(path):
    return _papx_module().load_document(
        path, max_bytes=MAX_PROBE_CFB_BYTES,
        max_aggregate_bytes=MAX_PROBE_CFB_BYTES,
    )


def _xml(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _style_xml(style_id, based_on, color, alignment, half_points, ascii_font, hansi_font):
    based = "" if based_on is None else f'<w:basedOn w:val="{_xml(based_on)}"/>'
    tabs = "".join(
        f'<w:tab w:val="left" w:pos="{100 + index * 80}"/>' for index in range(40)
    )
    jc = "" if alignment is None else f'<w:jc w:val="{alignment}"/>'
    paragraph = f'<w:pPr>{jc}<w:tabs>{tabs}</w:tabs></w:pPr>'
    run_values = []
    if ascii_font is not None or hansi_font is not None:
        run_values.append(
            f'<w:rFonts w:ascii="{_xml(ascii_font)}" w:hAnsi="{_xml(hansi_font)}"/>'
        )
    if color is not None:
        run_values.append(f'<w:color w:val="{color}"/>')
    if half_points is not None:
        run_values.append(f'<w:sz w:val="{half_points}"/>')
    run = "" if not run_values else '<w:rPr>' + "".join(run_values) + '</w:rPr>'
    return (
        f'<w:style w:type="table" w:customStyle="1" w:styleId="{_xml(style_id)}">'
        f'<w:name w:val="{_xml(style_id)}"/>{based}{paragraph}{run}</w:style>'
    )


def build_docx():
    """Return a deterministic passive eight-table DOCX source."""
    styles = "".join(_style_xml(*definition) for definition in STYLE_DEFINITIONS)
    direct = (
        '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>'
        '<w:color w:val="0000FF"/><w:sz w:val="24"/></w:rPr>'
    )

    def table_xml(ordinal, style_id, rows):
        columns = len(rows[0])
        width = 9000 // columns
        grid = "".join(f'<w:gridCol w:w="{width}"/>' for _ in range(columns))
        row_xml = []
        marker_index = 0
        for row in rows:
            cells = []
            for marker in row:
                marker_index += 1
                run = direct if ordinal == DIRECT_COLOR_TABLE else ""
                paragraph = '<w:p>'
                if ordinal == DIRECT_COLOR_TABLE:
                    paragraph += '<w:pPr><w:jc w:val="center"/></w:pPr>'
                paragraph += f'<w:r>{run}<w:t>{marker}</w:t></w:r></w:p>'
                cells.append(
                    f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
                    f'{paragraph}</w:tc>'
                )
            row_xml.append('<w:tr>' + "".join(cells) + '</w:tr>')
        return (
            '<w:tbl><w:tblPr>'
            f'<w:tblStyle w:val="{style_id}"/><w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/>'
            '</w:tblPr><w:tblGrid>' + grid + '</w:tblGrid>'
            + "".join(row_xml) + '</w:tbl>'
        )

    tables = []
    marker = 0
    for ordinal, style_id in enumerate(TABLE_STYLES, 1):
        if ordinal == 7:
            rows = []
            for _ in range(3):
                rows.append(MARKERS[marker:marker + 3])
                marker += 3
        else:
            rows = ((MARKERS[marker],),)
            marker += 1
        label = (
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
            f'<w:r><w:t>T{ordinal:02d}</w:t></w:r></w:p>'
        )
        tables.append(label + table_xml(ordinal, style_id, rows))
    if marker != len(MARKERS):
        raise AssertionError("probe marker inventory mismatch")
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>' + "".join(tables) +
        '<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>'
        '</w:sectPr></w:body></w:document>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr>'
        '</w:rPrDefault></w:docDefaults>' + styles + '</w:styles>'
    )
    entries = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>'
        ),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        "word/document.xml": document,
        "word/styles.xml": styles_xml,
    }
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(entries):
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, entries[name].encode("utf-8"))
    return output.getvalue()


def _write_new(path, value):
    path = Path(path)
    if path.exists():
        raise ProbeError(f"refusing to replace existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as output:
        output.write(value)


def _u16(data, offset, message="truncated 16-bit value"):
    if offset < 0 or offset + 2 > len(data):
        raise ProbeError(message)
    return int.from_bytes(data[offset:offset + 2], "little")


def _u32(data, offset, message="truncated 32-bit value"):
    if offset < 0 or offset + 4 > len(data):
        raise ProbeError(message)
    return int.from_bytes(data[offset:offset + 4], "little")


def _slice(data, start, end, message):
    if start < 0 or end < start or end > len(data):
        raise ProbeError(message)
    return data[start:end]


@dataclass(frozen=True)
class StyleRecord:
    istd: int
    name: str
    base: int
    kind: int
    start: int
    end: int
    base_size: int
    name_range: tuple
    property_ranges: tuple


def parse_styles(streams):
    """Parse just enough STSH/STD framing to identify the seven probe styles."""
    papx = _papx_module()
    checked = papx._streams(streams)
    fib = papx._fib(checked)
    word = checked["WordDocument"]
    table = checked[fib["table_stream"]]
    start = _u32(word, FC_STSHF)
    size = _u32(word, LCB_STSHF)
    stylesheet = _slice(table, start, start + size, "stylesheet outside table stream")
    header_size = _u16(stylesheet, 0, "truncated stylesheet header")
    header = _slice(stylesheet, 2, 2 + header_size, "truncated stylesheet header")
    count = _u16(header, 0)
    base_size = _u16(header, 2)
    if count < 1 or count > 4094 or base_size not in (10, 18):
        raise ProbeError("invalid stylesheet header")
    records = []
    position = 2 + header_size
    for istd in range(count):
        record_size = _u16(stylesheet, position, "truncated STD size")
        if record_size > 0x7FFF:
            raise ProbeError("negative STD size")
        record_start = position + 2
        record_end = record_start + record_size
        record = _slice(stylesheet, record_start, record_end, "STD outside stylesheet")
        position = record_end + record_size % 2
        if not record:
            continue
        if len(record) < base_size + 2:
            raise ProbeError("truncated STD")
        if _u16(record, 6) != record_size:
            raise ProbeError("STD bchUpe does not equal cbStd")
        kind_base = _u16(record, 2)
        kind = kind_base & 0xF
        base = kind_base >> 4
        property_count = _u16(record, 4) & 0xF
        name_units = _u16(record, base_size)
        name_start = base_size + 2
        name_end = name_start + name_units * 2
        raw_name = _slice(record, name_start, name_end, "STD name outside record")
        try:
            name = raw_name.decode("utf-16le")
        except UnicodeDecodeError as error:
            raise ProbeError("invalid STD UTF-16 name") from error
        cursor = name_end
        if record[cursor:cursor + 2] != b"\0\0":
            raise ProbeError("STD style name is not null-terminated")
        cursor += 2
        xstz_end = cursor
        ranges = []
        for _ in range(property_count):
            property_size = _u16(record, cursor, "truncated STD property size")
            property_start = cursor + 2
            property_end = property_start + property_size
            _slice(record, property_start, property_end, "STD property outside record")
            ranges.append((start + record_start + property_start,
                           start + record_start + property_end))
            if property_size % 2 and record[property_end:property_end + 1] != b"\0":
                raise ProbeError("STD UPX padding is not zero")
            cursor = property_end + property_size % 2
        if cursor > len(record):
            raise ProbeError("STD properties exceed record")
        if cursor != len(record):
            raise ProbeError("STD contains unexplained trailing bytes")
        records.append(StyleRecord(
            istd, name, base, kind, start + record_start,
            start + record_end, base_size,
            (start + record_start + base_size,
             start + record_start + xstz_end),
            tuple(ranges),
        ))
    if position > len(stylesheet):
        raise ProbeError("stylesheet records exceed stylesheet")
    return fib["table_stream"], tuple(records), (start, start + size)


def probe_styles(streams):
    table_stream, records, stylesheet_range = parse_styles(streams)
    expected_names = {definition[0] for definition in STYLE_DEFINITIONS}
    selected = [record for record in records if record.name in expected_names]
    by_name = {record.name: record for record in selected}
    if set(by_name) != expected_names:
        missing = sorted(expected_names - set(by_name))
        raise ProbeError(f"missing probe table styles: {', '.join(missing)}")
    if len(selected) != len(STYLE_DEFINITIONS):
        raise ProbeError("duplicate probe table style names")
    for name, based_on, _color, _alignment, _size, _ascii, _hansi in STYLE_DEFINITIONS:
        record = by_name[name]
        if record.kind != 3 or len(record.property_ranges) != 3:
            raise ProbeError(f"probe style {name} is not a three-UPX table style")
        expected_base = 0xFFF if based_on is None else by_name[based_on].istd
        if record.base != expected_base:
            raise ProbeError(f"probe style {name} has an unexpected base style")
    return table_stream, by_name, stylesheet_range


@dataclass(frozen=True)
class FormattingRun:
    fc_start: int
    fc_end: int
    header_offset: int | None
    offset: int | None
    end: int | None
    encoding: str


def _character_runs(streams, fib):
    word = streams["WordDocument"]
    table = streams[fib["table_stream"]]
    start = _u32(word, FC_PLCF_BTE_CHPX)
    size = _u32(word, LCB_PLCF_BTE_CHPX)
    plc = _slice(table, start, start + size, "CHPX page table outside selected table stream")
    if not plc:
        return ()
    if len(plc) < 4 or (len(plc) - 4) % 8:
        raise ProbeError("invalid CHPX page table")
    page_count = (len(plc) - 4) // 8
    if page_count > 65_536:
        raise ProbeError("CHPX page count exceeds policy")
    result = []
    for page_index in range(page_count):
        lower = _u32(plc, page_index * 4)
        upper = _u32(plc, page_index * 4 + 4)
        if lower >= upper:
            raise ProbeError("unordered CHPX page range")
        pn = _u32(plc, (page_count + 1) * 4 + page_index * 4) & 0x003FFFFF
        page_base = pn * 512
        page = _slice(word, page_base, page_base + 512, "CHP FKP outside WordDocument")
        count = page[511]
        if count < 1 or count > 101 or len(result) + count > 1_000_000:
            raise ProbeError("invalid or excessive CHPX run count")
        pointers = (count + 1) * 4
        payload_start = pointers + count
        for index in range(count):
            fc_start = _u32(page, index * 4)
            fc_end = _u32(page, index * 4 + 4)
            if fc_start >= fc_end or fc_start < lower or fc_end > upper:
                raise ProbeError("invalid CHPX physical range")
            payload = page[pointers + index] * 2
            if payload == 0:
                result.append(FormattingRun(fc_start, fc_end, None, None, None, "character"))
                continue
            if payload < payload_start or payload >= 511:
                raise ProbeError("CHPX payload overlaps FKP index")
            length = page[payload]
            end = payload + 1 + length
            if end > 511:
                raise ProbeError("CHPX payload exceeds FKP payload area")
            result.append(FormattingRun(
                fc_start, fc_end, page_base + payload, page_base + payload + 1,
                page_base + end, "character",
            ))
    return tuple(result)


def _main_characters(streams, fib, pieces):
    if fib["ccp_text"] > MAX_PROBE_STORY_UNITS:
        raise ProbeError("probe main story exceeds the character policy")
    word = streams["WordDocument"]
    result = []
    for piece in pieces:
        start = piece["cp_start"]
        end = min(piece["cp_end"], fib["ccp_text"])
        if start >= end:
            continue
        width = 1 if piece["compressed"] else 2
        for cp in range(start, end):
            fc = piece["fc_start"] + (cp - start) * width
            raw = _slice(word, fc, fc + width, "piece character outside WordDocument")
            try:
                character = raw.decode("cp1252" if width == 1 else "utf-16le")
            except UnicodeDecodeError as error:
                raise ProbeError("probe text uses an unsupported encoded character") from error
            result.append({"cp": cp, "fc": fc, "width": width, "character": character})
    if len(result) != fib["ccp_text"]:
        raise ProbeError("CLX does not map the complete probe main story")
    return result


def _find_marker_spans(characters):
    text = "".join(item["character"] for item in characters)
    spans = []
    for marker in MARKERS:
        start = text.find(marker)
        if start < 0 or text.find(marker, start + 1) >= 0:
            raise ProbeError(f"probe marker is missing or duplicated: {marker}")
        end = start + len(marker)
        spans.append((start, end))
    if any(left[1] > right[0] for left, right in zip(sorted(spans), sorted(spans)[1:])):
        raise ProbeError("probe marker ranges overlap")
    return tuple(spans)


def _active(trace, code):
    matches = [item for item in trace if item.get("kind") == "prl"
               and item.get("applied") and int(item["code"], 16) == code]
    return None if not matches else matches[-1]


def _run_at(runs, fc, label):
    selected = [run for run in runs if run.fc_start <= fc < run.fc_end]
    if len(selected) != 1:
        raise ProbeError(f"{label} does not select exactly one formatting run")
    return selected[0]


def _run_lookup(runs, label, start_key=lambda run: run.fc_start,
                end_key=lambda run: run.fc_end):
    ordered = sorted(runs, key=start_key)
    if any(end_key(left) > start_key(right) for left, right in zip(ordered, ordered[1:])):
        raise ProbeError(f"overlapping {label} physical ranges")
    starts = [start_key(run) for run in ordered]

    def lookup(fc):
        index = bisect_right(starts, fc) - 1
        if index < 0 or not (start_key(ordered[index]) <= fc < end_key(ordered[index])):
            raise ProbeError(f"{label} does not select exactly one formatting run")
        return ordered[index]

    return lookup


def _papx_run_at(runs, fc):
    selected = [run for run in runs if run["fc_start"] <= fc < run["fc_end"]]
    if len(selected) != 1 or selected[0]["papx"] is None:
        raise ProbeError("probe marker does not select exactly one PAPX payload")
    return selected[0]


def _probe_layout(loaded):
    papx = _papx_module()
    streams, fib, papx_runs, _prcs, pieces = papx._document_parts(loaded)
    inspected = papx.inspect_document(loaded)
    if len(inspected["papx_runs"]) != len(papx_runs):
        raise ProbeError("inconsistent PAPX inspection")
    characters = _main_characters(streams, fib, pieces)
    spans = _find_marker_spans(characters)
    traces = inspected["property_traces"]
    papx_at = _run_lookup(
        inspected["papx_runs"], "PAPX",
        start_key=lambda run: run["fc_start"], end_key=lambda run: run["fc_end"],
    )
    ttp = []
    for marker in (item for item in characters if item["character"] == "\x07"):
        run = papx_at(marker["fc"])
        trace_id = run.get("direct_trace_id")
        if trace_id is None:
            continue
        trace = traces[trace_id]["entries"]
        if (_active(trace, P_ITAP) or {}).get("operand") != "01000000":
            continue
        if (_active(trace, PF_TTP) or {}).get("operand") != "01":
            continue
        style = _active(trace, TI_STD)
        if style is None or len(bytes.fromhex(style["operand"])) != 2:
            raise ProbeError("TTP lacks one effective TIstd")
        ttp.append({"cp": marker["cp"], "fc": marker["fc"],
                    "run": run, "trace": trace, "tistd": style})
    ttp.sort(key=lambda item: item["cp"])
    if len(ttp) != len(ROW_STYLES):
        raise ProbeError("probe must contain exactly ten top-level TTP rows")

    row_marker_groups = ((0,), (1,), (2,), (3,), (4,), (5,),
                         (6, 7, 8), (9, 10, 11), (12, 13, 14), (15,))
    for index, (row, marker_group) in enumerate(zip(ttp, row_marker_groups)):
        first = min(spans[marker][0] for marker in marker_group)
        last = max(spans[marker][1] for marker in marker_group)
        next_first = (len(characters) if index + 1 == len(ttp)
                      else min(spans[marker][0] for marker in row_marker_groups[index + 1]))
        if not (last <= row["cp"] < next_first and first < row["cp"]):
            raise ProbeError("TTP order does not match the fixed table matrix")

    chpx_runs = _character_runs(streams, fib)
    chpx_at = _run_lookup(chpx_runs, "CHPX")
    marker_targets = []
    for ordinal, (start, end) in enumerate(spans, 1):
        selected_chpx = set()
        selected_chpx_runs = set()
        for character in characters[start:end]:
            run = chpx_at(character["fc"])
            if run.offset is not None:
                selected_chpx.add((run.header_offset, run.offset, run.end, run.encoding))
                selected_chpx_runs.add((run.fc_start, run.fc_end, run.header_offset,
                                        run.offset, run.end, run.encoding))
        pap_run = papx_at(characters[start]["fc"])
        if any(not (pap_run["fc_start"] <= characters[position]["fc"] < pap_run["fc_end"])
               for position in range(start, end)):
            raise ProbeError("probe marker crosses PAPX runs")
        paragraph_marks = [item for item in characters
                           if pap_run["fc_start"] <= item["fc"] < pap_run["fc_end"]
                           and item["character"] == "\r"]
        if len(paragraph_marks) != 1:
            raise ProbeError("probe marker PAPX must own exactly one paragraph mark")
        mark_run = chpx_at(paragraph_marks[0]["fc"])
        mark_chpx = None
        mark_chpx_run = None
        if mark_run.offset is not None:
            mark_chpx = (mark_run.header_offset, mark_run.offset,
                         mark_run.end, mark_run.encoding)
            mark_chpx_run = (mark_run.fc_start, mark_run.fc_end, mark_run.header_offset,
                             mark_run.offset, mark_run.end, mark_run.encoding)
        marker_targets.append({
            "ordinal": ordinal,
            "span": (start, end),
            "chpx": tuple(sorted(selected_chpx)),
            "chpx_runs": tuple(sorted(selected_chpx_runs)),
            "mark_chpx": mark_chpx,
            "mark_chpx_run": mark_chpx_run,
            "papx": pap_run["papx"],
            "papx_run": (pap_run["fc_start"], pap_run["fc_end"]),
            "preserve_direct": ordinal == DIRECT_COLOR_TABLE,
        })
    return {
        "streams": streams,
        "fib": fib,
        "characters": characters,
        "spans": spans,
        "ttp": tuple(ttp),
        "chpx_runs": chpx_runs,
        "papx_runs": inspected["papx_runs"],
        "markers": tuple(marker_targets),
    }


def _prl_segments(streams, stream, start, end):
    papx = _papx_module()
    trace = papx.trace_properties(streams, stream, start, end)
    if any(item.get("kind") != "prl" or item.get("followed") for item in trace):
        raise ProbeError("probe direct formatting unexpectedly uses indirection")
    if trace and (trace[0]["offset"] != start or trace[-1]["end"] != end):
        raise ProbeError("property trace does not cover the payload")
    if any(left["end"] != right["offset"] for left, right in zip(trace, trace[1:])):
        raise ProbeError("property trace is not contiguous")
    return trace


def _rewrite_property_payload(streams, descriptor, removed_codes):
    stream = descriptor["stream"]
    header = descriptor["header_offset"]
    start = descriptor["offset"]
    end = descriptor["end"]
    encoding = descriptor["encoding"]
    data = streams[stream]
    prefix = data[start:start + 2] if encoding != "character" else b""
    group_start = start + len(prefix)
    trace = _prl_segments(streams, stream, group_start, end)
    removed = [item for item in trace if int(item["code"], 16) in removed_codes]
    kept = [item for item in trace if int(item["code"], 16) not in removed_codes]
    if not removed:
        return None
    raw = prefix + b"".join(data[item["offset"]:item["end"]] for item in kept)
    capacity = end - start
    if len(raw) > capacity or capacity > MAX_REWRITE_BYTES:
        raise ProbeError("rewritten formatting payload exceeds policy")
    if encoding == "character":
        if len(raw) > 255:
            raise ProbeError("rewritten CHPX is too large")
        header_bytes = bytes((len(raw),))
    elif encoding in ("short", "extended"):
        if not raw or len(raw) > 510:
            raise ProbeError("rewritten PAPX has an invalid size")
        if len(raw) % 2:
            if len(raw) > 509:
                raise ProbeError("rewritten short PAPX is too large")
            header_bytes = bytes(((len(raw) + 1) // 2,))
        else:
            header_bytes = bytes((0, len(raw) // 2))
    else:
        raise ProbeError("unknown formatting payload encoding")
    replacement = header_bytes + raw + bytes(end - (header + len(header_bytes) + len(raw)))
    before = data[header:end]
    if len(replacement) != len(before):
        raise ProbeError("rewritten formatting payload changes its physical range")
    return stream, header, before, replacement, tuple(int(item["code"], 16) for item in removed)


def _descriptor(run):
    return {"stream": "WordDocument", "header_offset": run.header_offset,
            "offset": run.offset, "end": run.end, "encoding": run.encoding}


def _assert_unshared(layout, selected, all_runs, label):
    owners = {}
    for run in all_runs:
        if isinstance(run, FormattingRun):
            if run.offset is None:
                continue
            key = (run.header_offset, run.offset, run.end, run.encoding)
            owner = (run.fc_start, run.fc_end)
        else:
            raw = run.get("papx")
            if raw is None:
                continue
            key = (raw["header_offset"], raw["offset"], raw["end"], raw["encoding"])
            owner = (run["fc_start"], run["fc_end"])
        owners.setdefault(key, set()).add(owner)
    for key, target_owners in selected.items():
        if owners.get(key) != target_owners:
            raise ProbeError(f"{label} payload is shared with an untargeted formatting run")


def _apply(streams, edits):
    result = {name: bytearray(value) for name, value in streams.items()}
    ranges = {}
    for stream, offset, before, after, _semantic in sorted(edits, key=lambda item: (item[0], item[1])):
        if len(before) != len(after) or before == after:
            raise ProbeError("mutation edits must be changed same-length ranges")
        end = offset + len(before)
        prior = ranges.setdefault(stream, [])
        if prior and prior[-1][1] > offset:
            raise ProbeError("mutation edits overlap")
        prior.append((offset, end))
        if bytes(result[stream][offset:end]) != before:
            raise ProbeError("mutation edit does not match source bytes")
        result[stream][offset:end] = after
    return {name: bytes(value) for name, value in result.items()}


def _diff_edits(before, after):
    edits = []
    total = 0
    if set(before) != set(after):
        raise ProbeError("candidate changes the CFB stream inventory")
    for stream in sorted(before):
        if len(before[stream]) != len(after[stream]):
            raise ProbeError("candidate changes a CFB stream length")
        position = 0
        while position < len(before[stream]):
            if before[stream][position] == after[stream][position]:
                position += 1
                continue
            start = position
            while position < len(before[stream]) and before[stream][position] != after[stream][position]:
                position += 1
            total += position - start
            if total > MAX_REWRITE_BYTES:
                raise ProbeError("mutation diff exceeds policy")
            edits.append({"stream": stream, "offset": start,
                          "before": before[stream][start:position].hex(),
                          "after": after[stream][start:position].hex()})
    return edits


def _json_bytes(value):
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if len(encoded) > MAX_JSON_BYTES:
        raise ProbeError("plan exceeds JSON size policy")
    return encoded


def _payload_descriptor(raw):
    return {
        "stream": raw.get("stream", "WordDocument"),
        "header_offset": raw["header_offset"],
        "offset": raw["offset"],
        "end": raw["end"],
        "encoding": raw["encoding"],
    }


def _payload_key(descriptor):
    return (descriptor["header_offset"], descriptor["offset"],
            descriptor["end"], descriptor["encoding"])


def _normalization_edits(layout):
    selected_chpx = {}
    selected_papx = {}
    for marker in layout["markers"]:
        if (marker["preserve_direct"] and marker["mark_chpx_run"] is not None
                and marker["mark_chpx_run"] in marker["chpx_runs"]):
            raise ProbeError("T06 visible body shares its CHPX run with the paragraph mark")
        chpx_targets = [] if marker["preserve_direct"] else list(marker["chpx_runs"])
        if marker["mark_chpx_run"] is not None:
            chpx_targets.append(marker["mark_chpx_run"])
        for owner in chpx_targets:
            descriptor = {
                "stream": "WordDocument", "header_offset": owner[2],
                "offset": owner[3], "end": owner[4], "encoding": owner[5],
            }
            trace = _prl_segments(
                layout["streams"], descriptor["stream"], descriptor["offset"], descriptor["end"]
            )
            if any(int(item["code"], 16) in SUPPORTED_CHPX for item in trace):
                key = _payload_key(descriptor)
                selected_chpx.setdefault(key, {"descriptor": descriptor, "owners": set()})
                selected_chpx[key]["owners"].add((owner[0], owner[1]))
        if marker["preserve_direct"]:
            continue
        raw = marker["papx"]
        descriptor = _payload_descriptor(raw)
        trace = _prl_segments(
            layout["streams"], descriptor["stream"], descriptor["offset"] + 2,
            descriptor["end"],
        )
        if any(int(item["code"], 16) in SUPPORTED_PAPX for item in trace):
            key = _payload_key(descriptor)
            selected_papx.setdefault(key, {"descriptor": descriptor, "owners": set()})
            selected_papx[key]["owners"].add(marker["papx_run"])

    _assert_unshared(
        layout, {key: item["owners"] for key, item in selected_chpx.items()},
        layout["chpx_runs"], "CHPX"
    )
    _assert_unshared(
        layout, {key: item["owners"] for key, item in selected_papx.items()},
        layout["papx_runs"], "PAPX"
    )
    edits = []
    removed = {"chpx": [], "papx": []}
    for kind, selected, codes in (
        ("chpx", selected_chpx, SUPPORTED_CHPX),
        ("papx", selected_papx, SUPPORTED_PAPX),
    ):
        descriptors = [item["descriptor"] for item in selected.values()]
        for descriptor in sorted(descriptors, key=lambda item: item["header_offset"]):
            edit = _rewrite_property_payload(layout["streams"], descriptor, codes)
            if edit is None:
                continue
            edits.append(edit)
            owners = selected[_payload_key(descriptor)]["owners"]
            removed[kind].append({
                "offset": descriptor["header_offset"],
                "end": descriptor["end"],
                "owners": [list(owner) for owner in sorted(owners)],
                "codes": [f"{code:04x}" for code in edit[4]],
            })
    return edits, removed


def _connection_edits(streams, rows, styles):
    desired_by_range = {}
    targets = []
    edits = []
    for ordinal, (row, style_name) in enumerate(zip(rows, ROW_STYLES), 1):
        source = row["tistd"]
        if source["stream"] not in streams or source["end"] - source["offset"] != 4:
            raise ProbeError("TTP TIstd does not have direct two-byte framing")
        operand_offset = source["offset"] + 2
        before = bytes.fromhex(source["operand"])
        after = styles[style_name].istd.to_bytes(2, "little")
        identity = (source["stream"], operand_offset, operand_offset + 2)
        prior = desired_by_range.get(identity)
        if prior is not None and prior != after:
            raise ProbeError("shared TTP TIstd would require conflicting styles")
        desired_by_range[identity] = after
        if before != after:
            if prior is None:
                edits.append((source["stream"], operand_offset, before, after,
                              ("TIstd", ordinal, style_name)))
            targets.append({"fc": row["fc"], "code": "TIstd",
                            "before": before.hex(), "after": after.hex()})
    if not targets:
        raise ProbeError("source TTP rows are already connected to the probe styles")
    return edits, targets


def _style_manifest(table_stream, styles, stylesheet_range, streams):
    return {
        "stream": table_stream,
        "offset": stylesheet_range[0],
        "end": stylesheet_range[1],
        "sha256": sha256(
            streams[table_stream][stylesheet_range[0]:stylesheet_range[1]]
        ).hexdigest(),
        "styles": [
            {"name": name, "istd": record.istd, "base": record.base,
             "kind": record.kind,
             "property_ranges": [list(value) for value in record.property_ranges]}
            for name, record in sorted(styles.items())
        ],
    }


def _hex_field(value, label):
    if not isinstance(value, str) or len(value) % 2:
        raise ProbeError(f"{label} must be whole-byte hexadecimal")
    try:
        result = bytes.fromhex(value)
    except ValueError as error:
        raise ProbeError(f"{label} must be whole-byte hexadecimal") from error
    if len(result) > 65_535:
        raise ProbeError(f"{label} exceeds the UPX size limit")
    return result


def _validate_style_upx(kind, value, istd):
    papx = _papx_module()
    if kind == "papx":
        if len(value) < 2 or int.from_bytes(value[:2], "little") != istd:
            raise ProbeError("table-style PAPX must begin with its own istd")
        group = value[2:]
        expected_sgc = 1
    else:
        group = value
        expected_sgc = 5 if kind == "tapx" else 2
    trace = papx.trace_properties({"UPX": group}, "UPX", 0, len(group))
    if any(item.get("kind") != "prl" or item.get("followed") for item in trace):
        raise ProbeError(f"table-style {kind.upper()} cannot contain indirection")
    for item in trace:
        if ((int(item["code"], 16) >> 10) & 7) != expected_sgc:
            raise ProbeError(f"table-style {kind.upper()} contains the wrong SPRM class")


def _padded_name(original, units, property_bytes):
    if not original.isascii():
        raise ProbeError("probe style names must remain ASCII")
    if units < len(original):
        raise ProbeError(f"style {original} has insufficient name reserve")
    if units == len(original):
        return original
    digest = sha256(property_bytes).hexdigest()[:12]
    suffix = "_" + digest
    value = (original + suffix)[:units]
    return value + "_" * (units - len(value))


def _rewrite_style(streams, table_stream, record, replacement):
    if not isinstance(replacement, Mapping) or not replacement:
        raise ProbeError(f"style {record.name} replacement must be a nonempty object")
    unknown = set(replacement) - {"tapx", "papx", "chpx"}
    if unknown:
        raise ProbeError(f"style {record.name} replacement has unknown property sets")
    table = streams[table_stream]
    old_record = table[record.start:record.end]
    original_values = [table[left:right] for left, right in record.property_ranges]
    values = []
    for index, kind in enumerate(("tapx", "papx", "chpx")):
        value = (original_values[index] if kind not in replacement
                 else _hex_field(replacement[kind], f"style {record.name} {kind}"))
        _validate_style_upx(kind, value, record.istd)
        values.append(value)
    lp_upxes = b"".join(
        len(value).to_bytes(2, "little") + value + bytes(len(value) % 2)
        for value in values
    )
    name_bytes = len(old_record) - record.base_size - len(lp_upxes)
    if name_bytes < 4 or name_bytes % 2:
        raise ProbeError(f"style {record.name} replacement exceeds its same-length reserve")
    units = (name_bytes - 4) // 2
    name = _padded_name(record.name, units, b"".join(values))
    xstz = units.to_bytes(2, "little") + name.encode("utf-16le") + b"\0\0"
    rewritten = old_record[:record.base_size] + xstz + lp_upxes
    if len(rewritten) != len(old_record):
        raise ProbeError("rewritten STD changes cbStd")
    if rewritten[:record.base_size] != old_record[:record.base_size]:
        raise ProbeError("rewritten STD changes Stdf metadata")
    return table_stream, record.start, old_record, rewritten, {
        "logical_name": record.name, "stored_name": name, "istd": record.istd,
        "property_sha256": [sha256(value).hexdigest() for value in values],
    }


def _table_option_edits(layout, table_options):
    if table_options is None:
        return [], []
    if not isinstance(table_options, Mapping) or set(table_options) - set(ROW_TABLES):
        raise ProbeError("table_options must select only T01 through T08")
    edits = {}
    targets = []
    for row, table_name in zip(layout["ttp"], ROW_TABLES):
        if table_name not in table_options:
            continue
        flags = table_options[table_name]
        if type(flags) is not int or not 0 <= flags <= 0xFFFF:
            raise ProbeError("table option flags must be unsigned 16-bit integers")
        current = _active(row["trace"], 0x740A)
        if current is not None:
            before = bytes.fromhex(current["operand"])
            if len(before) != 4:
                raise ProbeError("TTlp has invalid framing")
            after = before[:2] + flags.to_bytes(2, "little")
            stream, offset = current["stream"], current["offset"] + 2
        else:
            revision = [item for item in row["trace"]
                        if item.get("kind") == "prl" and item.get("applied")
                        and int(item["code"], 16) == 0x6467
                        and item.get("stream") == "Data"]
            if len(revision) != 1:
                raise ProbeError("adding TTlp requires one terminal Data sprmPRsid reserve")
            source = revision[0]
            before = layout["streams"]["Data"][source["offset"]:source["end"]]
            after = (0x740A).to_bytes(2, "little") + b"\0\0" + flags.to_bytes(2, "little")
            if len(before) != len(after):
                raise ProbeError("sprmPRsid reserve does not match TTlp framing")
            stream, offset = "Data", source["offset"]
        identity = (stream, offset, len(before))
        previous = edits.get(identity)
        if previous is not None and previous[3] != after:
            raise ProbeError("shared row options would require conflicting edits")
        if before != after:
            edits[identity] = (stream, offset, before, after, ("TTlp", table_name))
            targets.append({
                "fc": row["fc"], "code": "TTlp",
                "before": None if current is None else current["operand"],
                "after": after[-4:].hex(),
            })
    return list(edits.values()), targets


def build_recipe_variant(loaded, recipe):
    """Apply one bounded declarative recipe without serializing a CFB file."""
    if not isinstance(recipe, Mapping) or recipe.get("schema") != RECIPE_SCHEMA:
        raise ProbeError("unsupported table-style recipe")
    allowed = {"schema", "strip_direct", "connect_styles", "styles", "table_options"}
    if set(recipe) - allowed:
        raise ProbeError("recipe contains unknown fields")
    if type(recipe.get("strip_direct", False)) is not bool:
        raise ProbeError("strip_direct must be boolean")
    if type(recipe.get("connect_styles", False)) is not bool:
        raise ProbeError("connect_styles must be boolean")
    layout = _probe_layout(loaded)
    table_stream, styles, stylesheet_range = probe_styles(loaded)
    edits = []
    targets = []
    removed = {"chpx": [], "papx": []}
    if recipe.get("strip_direct", False):
        normalized, removed = _normalization_edits(layout)
        edits.extend(normalized)
    if recipe.get("connect_styles", False):
        connected, style_targets = _connection_edits(layout["streams"], layout["ttp"], styles)
        edits.extend(connected)
        targets.extend(style_targets)
    replacements = recipe.get("styles", {})
    if not isinstance(replacements, Mapping) or set(replacements) - set(styles):
        raise ProbeError("recipe styles must select named probe styles")
    style_changes = []
    for name in sorted(replacements):
        edit = _rewrite_style(layout["streams"], table_stream, styles[name], replacements[name])
        edits.append((edit[0], edit[1], edit[2], edit[3], ("STD", name)))
        style_changes.append(edit[4])
    option_edits, option_targets = _table_option_edits(
        layout, recipe.get("table_options")
    )
    edits.extend(option_edits)
    targets.extend(option_targets)
    if not edits:
        raise ProbeError("recipe does not produce any changes")
    candidate = _apply(layout["streams"], edits)
    _candidate_table, candidate_records, candidate_range = parse_styles(candidate)
    if candidate_range != stylesheet_range:
        raise ProbeError("style rewrite changes the stylesheet range")
    candidate_names = [record.name for record in candidate_records]
    if len(candidate_names) != len(set(candidate_names)):
        raise ProbeError("style rewrite creates a duplicate style name")
    candidate_by_istd = {record.istd: record for record in candidate_records}
    for item in style_changes:
        record = candidate_by_istd.get(item["istd"])
        if record is None or record.name != item["stored_name"]:
            raise ProbeError("rewritten style identity does not match the plan")
        actual_hashes = [
            sha256(candidate[table_stream][left:right]).hexdigest()
            for left, right in record.property_ranges
        ]
        if actual_hashes != item["property_sha256"]:
            raise ProbeError("rewritten style properties do not match the plan")
    if recipe.get("strip_direct", False):
        candidate_layout = _probe_layout(candidate)
        for marker in candidate_layout["markers"]:
            if _mark_direct_codes(candidate_layout, marker):
                raise ProbeError("direct formatting remains on a marker paragraph mark")
            if not marker["preserve_direct"] and any(_direct_codes(candidate_layout, marker)):
                raise ProbeError("direct formatting remains on a control marker")
    plan = {
        "schema": SCHEMA, "mode": "recipe",
        "scope": "fixed passive table-style acquisition controls; not display validity",
        "source_sha256": loaded.source_sha256,
        "source_stream_sha256": {
            name: sha256(value).hexdigest() for name, value in sorted(loaded.streams.items())
        },
        "stylesheet_before": _style_manifest(
            table_stream, styles, stylesheet_range, layout["streams"]
        ),
        "style_changes": style_changes,
        "removed_direct": removed,
        "recipe": recipe,
        "edits": _diff_edits(loaded.streams, candidate),
        "targets": targets,
    }
    if targets:
        _papx_module().validate_plan(loaded, candidate, {
            "source_sha256": loaded.source_sha256,
            "edits": plan["edits"], "targets": targets,
        })
    return candidate, plan


def validate_recipe_variant(before, after_streams, plan):
    if not isinstance(plan, Mapping) or plan.get("schema") != SCHEMA or plan.get("mode") != "recipe":
        raise ProbeError("candidate does not use a table-style recipe plan")
    expected, rebuilt_plan = build_recipe_variant(before, plan.get("recipe"))
    after = _papx_module()._streams(after_streams)
    if expected != after:
        raise ProbeError("candidate streams do not match the reviewed recipe")
    for field in ("source_sha256", "source_stream_sha256", "stylesheet_before",
                  "style_changes", "removed_direct", "edits", "targets"):
        if plan.get(field) != rebuilt_plan.get(field):
            raise ProbeError(f"plan {field} does not match the reviewed recipe")
    return {"valid": True, "mode": "recipe", "source_sha256": before.source_sha256,
            "edits": len(plan["edits"]), "targets": len(plan["targets"])}


def _plan(loaded, mode, streams, styles_manifest, layout, removed, targets):
    return {
        "schema": SCHEMA,
        "mode": mode,
        "scope": "fixed passive table-style acquisition controls; not display validity",
        "source_sha256": loaded.source_sha256,
        "source_stream_sha256": {
            name: sha256(value).hexdigest() for name, value in sorted(loaded.streams.items())
        },
        "stylesheet": styles_manifest,
        "markers": [
            {"ordinal": item["ordinal"], "text": MARKERS[item["ordinal"] - 1],
             "preserve_direct": item["preserve_direct"]}
            for item in layout["markers"]
        ],
        "rows": [
            {"ordinal": ordinal, "fc": row["fc"], "style": ROW_STYLES[ordinal - 1],
             "before_tistd": row["tistd"]["operand"],
             "after_tistd": (
                 row["tistd"]["operand"] if mode == "negative"
                 else styles_manifest["styles_by_name"][ROW_STYLES[ordinal - 1]]["istd_hex"]
             )}
            for ordinal, row in enumerate(layout["ttp"], 1)
        ],
        "removed_direct": removed,
        "edits": _diff_edits(loaded.streams, streams),
        "targets": targets,
    }


def prepare_variants(loaded, validate=True):
    """Create in-memory negative and style-connected candidates and exact plans."""
    layout = _probe_layout(loaded)
    table_stream, styles, stylesheet_range = probe_styles(loaded)
    styles_manifest = _style_manifest(
        table_stream, styles, stylesheet_range, layout["streams"]
    )
    styles_manifest["styles_by_name"] = {
        name: {"istd": record.istd, "istd_hex": record.istd.to_bytes(2, "little").hex()}
        for name, record in styles.items()
    }
    normalization, removed = _normalization_edits(layout)
    negative = _apply(layout["streams"], normalization)
    connection, targets = _connection_edits(negative, layout["ttp"], styles)
    connected = _apply(negative, connection)
    negative_plan = _plan(
        loaded, "negative", negative, styles_manifest, layout, removed, []
    )
    connected_plan = _plan(
        loaded, "connected", connected, styles_manifest, layout, removed, targets
    )
    for plan in (negative_plan, connected_plan):
        plan["stylesheet"].pop("styles_by_name", None)
    if validate:
        validate_variant(loaded, negative, negative_plan)
        validate_variant(loaded, connected, connected_plan)
    return {"negative": (negative, negative_plan),
            "connected": (connected, connected_plan)}


def _validate_exact_edits(before, after, plan):
    if plan.get("source_sha256") != before.source_sha256:
        raise ProbeError("plan source hash does not match serialized source")
    if plan.get("source_stream_sha256") != {
            name: sha256(value).hexdigest() for name, value in sorted(before.streams.items())}:
        raise ProbeError("plan source stream hashes do not match source")
    if plan.get("edits") != _diff_edits(before.streams, after):
        raise ProbeError("plan is not the exact candidate stream diff")


def _direct_codes(layout, marker):
    chpx = []
    for header, offset, end, _encoding in marker["chpx"]:
        chpx.extend(int(item["code"], 16) for item in _prl_segments(
            layout["streams"], "WordDocument", offset, end
        ))
    raw = marker["papx"]
    papx = [int(item["code"], 16) for item in _prl_segments(
        layout["streams"], raw["stream"], raw["offset"] + 2, raw["end"]
    )]
    return set(chpx) & SUPPORTED_CHPX, set(papx) & SUPPORTED_PAPX


def _mark_direct_codes(layout, marker):
    raw = marker["mark_chpx"]
    if raw is None:
        return set()
    return {
        int(item["code"], 16) for item in _prl_segments(
            layout["streams"], "WordDocument", raw[1], raw[2]
        )
    } & SUPPORTED_CHPX


def validate_variant(before, after_streams, plan):
    """Validate one generated in-memory candidate against its exact plan."""
    if not isinstance(plan, Mapping) or plan.get("schema") != SCHEMA:
        raise ProbeError("unsupported table-style probe plan")
    mode = plan.get("mode")
    if mode not in ("negative", "connected"):
        raise ProbeError("invalid table-style probe mode")
    papx = _papx_module()
    after = papx._streams(after_streams)
    expected_streams, expected_plan = prepare_variants(before, validate=False)[mode]
    if after != expected_streams:
        raise ProbeError("candidate is not the deterministic fixed-probe mutation")
    if plan != expected_plan:
        raise ProbeError("plan is not the canonical fixed-probe mutation plan")
    _validate_exact_edits(before, after, plan)
    before_table, before_styles, before_range = probe_styles(before)
    after_table, after_styles, after_range = probe_styles(after)
    if (before_table, before_range, before_styles) != (after_table, after_range, after_styles):
        raise ProbeError("candidate changes stylesheet structure")
    if (before.streams[before_table][before_range[0]:before_range[1]]
            != after[after_table][after_range[0]:after_range[1]]):
        raise ProbeError("candidate changes stylesheet bytes")
    layout = _probe_layout(after)
    if len(plan.get("markers", ())) != len(MARKERS) or len(plan.get("rows", ())) != len(ROW_STYLES):
        raise ProbeError("plan inventory does not match fixed probe")
    for marker in layout["markers"]:
        chpx, papx_codes = _direct_codes(layout, marker)
        if _mark_direct_codes(layout, marker):
            raise ProbeError("flattened direct formatting remains on a paragraph mark")
        if marker["preserve_direct"]:
            if not ((chpx & {CI_CO, C_CV}) and {0x4A43, 0x4A4F, 0x4A51} <= chpx
                    and P_JC in papx_codes):
                raise ProbeError("T06 direct override is incomplete")
        elif chpx or papx_codes:
            raise ProbeError("flattened direct formatting remains on a control marker")
    for ordinal, row in enumerate(layout["ttp"], 1):
        actual = row["tistd"]["operand"]
        expected = (plan["rows"][ordinal - 1]["before_tistd"] if mode == "negative"
                    else before_styles[ROW_STYLES[ordinal - 1]].istd.to_bytes(2, "little").hex())
        if actual != expected:
            raise ProbeError("candidate TTP style connection does not match plan")
    if mode == "connected":
        papx.validate_plan(before, after, {
            "source_sha256": plan["source_sha256"],
            "edits": plan["edits"], "targets": plan["targets"],
        })
    elif plan.get("targets") != []:
        raise ProbeError("negative plan must not claim effective scalar changes")
    return {"valid": True, "mode": mode, "source_sha256": before.source_sha256,
            "edits": len(plan["edits"]), "markers": len(MARKERS), "rows": len(ROW_STYLES)}


def _write_doc_candidate(source, output, expected_streams):
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    owned = None
    try:
        with output.open("x+b") as candidate_file:
            info = os.fstat(candidate_file.fileno())
            if not stat.S_ISREG(info.st_mode):
                raise ProbeError("candidate output is not a regular file")
            owned = (info.st_dev, info.st_ino)
            candidate_file.write(source.source_bytes)
            candidate_file.flush()
            import importlib
            olefile = importlib.import_module("olefile")
            candidate_file.seek(0)
            handle = olefile.OleFileIO(candidate_file, write_mode=True)
            try:
                for name, value in expected_streams.items():
                    if value != source.streams[name]:
                        handle.write_stream(name.split("/"), value)
            finally:
                handle.close()
            candidate_file.seek(0)
            serialized = candidate_file.read(MAX_PROBE_CFB_BYTES + 1)
            if len(serialized) != len(source.source_bytes):
                raise ProbeError("candidate serialized length changed")
            current = output.lstat()
            if (current.st_dev, current.st_ino) != owned:
                raise ProbeError("candidate output path was replaced")
        actual = _papx_module().load_document_bytes(
            serialized, max_aggregate_bytes=MAX_PROBE_CFB_BYTES
        )
        if dict(actual.streams) != dict(expected_streams):
            raise ProbeError("written candidate streams do not match prepared bytes")
    except Exception:
        try:
            current = output.lstat()
            if owned is not None and (current.st_dev, current.st_ino) == owned:
                output.unlink()
        except FileNotFoundError:
            pass
        raise


def _load_plan(path):
    path = Path(path)
    initial = path.stat()
    if initial.st_size > MAX_JSON_BYTES or not stat.S_ISREG(initial.st_mode):
        raise ProbeError("plan exceeds JSON size policy")
    with path.open("rb") as source:
        value = source.read(MAX_JSON_BYTES + 1)
    final = path.stat()
    if (len(value) != initial.st_size or final.st_size != initial.st_size
            or (final.st_dev, final.st_ino) != (initial.st_dev, initial.st_ino)):
        raise ProbeError("plan changed while being read")
    try:
        return json.loads(value.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ProbeError("plan is not valid UTF-8 JSON") from error


def _command_prepare(args):
    source = _load_probe_document(args.source)
    variants = prepare_variants(source)
    output_dir = Path(args.output_dir)
    if output_dir.exists():
        raise ProbeError(f"refusing to use existing output directory: {output_dir}")
    output_dir.mkdir(parents=True)
    written = []
    try:
        for mode, (streams, plan) in variants.items():
            document = output_dir / f"{mode}.doc"
            plan_path = output_dir / f"{mode}.plan.json"
            _write_doc_candidate(source, document, streams)
            written.append(document)
            _write_new(plan_path, _json_bytes(plan))
            written.append(plan_path)
    except Exception:
        for path in written:
            path.unlink(missing_ok=True)
        try:
            output_dir.rmdir()
        except OSError:
            pass
        raise
    print(json.dumps({"source_sha256": source.source_sha256, "output_dir": str(output_dir),
                      "variants": sorted(variants)}))


def _command_validate(args):
    source = _load_probe_document(args.source)
    candidate = _load_probe_document(args.candidate)
    plan = _load_plan(args.plan)
    result = (validate_recipe_variant(source, candidate.streams, plan)
              if plan.get("mode") == "recipe"
              else validate_variant(source, candidate.streams, plan))
    print(json.dumps(result, indent=2, sort_keys=True))


def _command_mutate(args):
    source = _load_probe_document(args.source)
    recipe = _load_plan(args.recipe)
    streams, plan = build_recipe_variant(source, recipe)
    _write_doc_candidate(source, args.output, streams)
    try:
        _write_new(args.plan, _json_bytes(plan))
    except Exception:
        Path(args.output).unlink(missing_ok=True)
        raise
    print(json.dumps({"source_sha256": source.source_sha256,
                      "output": str(Path(args.output)), "plan": str(Path(args.plan)),
                      "edits": len(plan["edits"]), "targets": len(plan["targets"])}))


def _command_generate(args):
    value = build_docx()
    _write_new(args.output, value)
    print(json.dumps({"output": str(Path(args.output)), "sha256": sha256(value).hexdigest(),
                      "tables": TABLE_COUNT, "styles": len(STYLE_DEFINITIONS)}))


def _command_inspect(args):
    loaded = _load_probe_document(args.source)
    table_stream, styles, stylesheet = probe_styles(loaded)
    print(json.dumps({
        "source_sha256": loaded.source_sha256,
        "table_stream": table_stream,
        "stylesheet": {"offset": stylesheet[0], "end": stylesheet[1]},
        "styles": [{"name": name, "istd": record.istd, "base": record.base,
                    "property_ranges": record.property_ranges}
                   for name, record in sorted(styles.items())],
    }, indent=2, sort_keys=True))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    generate = commands.add_parser("generate-docx", help="write the passive eight-table DOCX")
    generate.add_argument("output")
    generate.set_defaults(run=_command_generate)
    inspect = commands.add_parser("inspect", help="inspect retained probe styles in a DOC")
    inspect.add_argument("source")
    inspect.set_defaults(run=_command_inspect)
    prepare = commands.add_parser(
        "prepare", help="write same-length negative and style-connected DOC variants"
    )
    prepare.add_argument("source")
    prepare.add_argument("output_dir")
    prepare.set_defaults(run=_command_prepare)
    validate = commands.add_parser("validate", help="validate one generated DOC and plan")
    validate.add_argument("source")
    validate.add_argument("candidate")
    validate.add_argument("plan")
    validate.set_defaults(run=_command_validate)
    mutate = commands.add_parser(
        "mutate", help="apply a bounded declarative style/row-options recipe"
    )
    mutate.add_argument("source")
    mutate.add_argument("recipe")
    mutate.add_argument("output")
    mutate.add_argument("plan")
    mutate.set_defaults(run=_command_mutate)
    args = parser.parse_args(argv)
    try:
        args.run(args)
    except (OSError, ValueError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()
