#!/usr/bin/env python3
"""Inspect and validate bounded Word 97-2003 PAPX mutation plans.

This read-only tool follows [MS-DOC] 2.4.6.1 direct paragraph formatting,
2.6.2 PHugePapx/PTableProps replacement, 2.9.174/2.9.175 PAP FKP framing,
and 2.9.210 PrcData bounds. It validates mutation plans for scalar TIstd and
TTlp targets and can separately assert acquired serialized SPRM traces.
"""

import argparse
from bisect import bisect_right
from dataclasses import dataclass, field
from hashlib import sha256
import importlib
from io import BytesIO
import json
from pathlib import Path
import re
from types import MappingProxyType
from typing import Mapping


MAX_INPUT_BYTES = 256 * 1024 * 1024
MAX_STREAM_BYTES = 256 * 1024 * 1024
MAX_AGGREGATE_STREAM_BYTES = 256 * 1024 * 1024
MAX_FKP_PAGES = 65_536
MAX_PAPX_RUNS = 1_000_000
MAX_PIECES = 1_000_000
MAX_PRCS = 32_768
MAX_PRC_GRPPRL_BYTES = 0x3FA2
MIN_INDIRECT_GRPPRL_BYTES = 10
MAX_TRACE_DEPTH = 64
MAX_TRACE_PRLS = 1_000_000
MAX_TRACE_BYTES = 4 * 1024 * 1024
MAX_INSPECT_WORK = 16 * 1024 * 1024
MAX_DIAGNOSTIC_RECORDS = 100_000
MAX_INSPECT_PAYLOAD_BYTES = 8 * 1024 * 1024
MAX_TARGETS = 4_096
MAX_TARGET_LOOKUP_WORK = 1_000_000
MAX_EDITS = 4_096
MAX_PLAN_EDIT_BYTES = 4 * 1024 * 1024
MAX_PLAN_BYTES = 8 * 1024 * 1024
TRACE_ASSERTION_SCHEMA = "legacy-doc-property-trace/v1"

FIB_IDENT = 0xA5EC
FIB_WORD_97 = 0x00C1
FIB_FLAGS_OFFSET = 0x0A
CCP_TEXT_OFFSET = 0x4C
FC_PLCF_BTE_PAPX_OFFSET = 0x102
LCB_PLCF_BTE_PAPX_OFFSET = 0x106
FC_CLX_OFFSET = 0x1A2
LCB_CLX_OFFSET = 0x1A6

PHUGE_PAPX = 0x6646
P_TABLE_PROPS = 0x646B
TARGET_CODES = {"TIstd": 0x563A, "TTlp": 0x740A}
TARGET_OPERAND_SIZES = {0x563A: 2, 0x740A: 4}

# [MS-DOC] 2.9.215. Keep this closed table separate from the general SPRM
# decoder: every Prm0 operand is one byte, and unlisted isprm values are
# reserved rather than alternate encodings that can be inferred from ispmd.
PRM0_PARAGRAPH_SPRMS = {
    0x04: 0x2602, 0x05: 0x2461, 0x07: 0x2405, 0x08: 0x2406,
    0x09: 0x2407, 0x0C: 0x260A, 0x0D: 0x2470, 0x0E: 0x240C,
    0x0F: 0x2471, 0x18: 0x2416, 0x19: 0x2417, 0x1D: 0x261B,
    0x25: 0x2423, 0x2C: 0x242A, 0x32: 0x2430, 0x33: 0x2431,
    0x35: 0x2433, 0x36: 0x2434, 0x37: 0x2435, 0x38: 0x2436,
    0x39: 0x2437, 0x3A: 0x2438, 0x78: 0x2640, 0x7E: 0x2443,
}
PRM0_CHARACTER_SPRMS = {
    0x00: 0x2879, 0x41: 0x0800, 0x42: 0x0801, 0x43: 0x0802,
    0x47: 0x0806, 0x4B: 0x080A, 0x4D: 0x2A0C, 0x4E: 0x0858,
    0x4F: 0x2859, 0x50: 0x0811, 0x51: 0x0818, 0x53: 0x2A33,
    0x55: 0x0835, 0x56: 0x0836, 0x57: 0x0837, 0x58: 0x0838,
    0x59: 0x0839, 0x5A: 0x083A, 0x5B: 0x083B, 0x5C: 0x083C,
    0x5E: 0x2A3E, 0x62: 0x2A42, 0x68: 0x2A48, 0x73: 0x2A53,
    0x74: 0x0854, 0x75: 0x0855, 0x76: 0x0856, 0x7B: 0x2A90,
    0x7C: 0x2A86,
}


class ProbeError(ValueError):
    """A malformed document, trace request, or mutation plan."""


def _checked_streams(streams):
    if not isinstance(streams, Mapping):
        raise ProbeError("streams must be a mapping")
    result = {}
    total = 0
    for name, value in streams.items():
        if not isinstance(name, str) or not name:
            raise ProbeError("invalid stream name")
        if not isinstance(value, (bytes, bytearray, memoryview)):
            raise ProbeError(f"stream {name} is not bytes")
        size = value.nbytes if isinstance(value, memoryview) else len(value)
        if size > MAX_STREAM_BYTES:
            raise ProbeError(f"stream {name} exceeds the size policy")
        total += size
        if total > MAX_AGGREGATE_STREAM_BYTES:
            raise ProbeError("aggregate streams exceed the size policy")
        result[name] = bytes(value)
    return result


@dataclass(frozen=True)
class LoadedDocument:
    """All CFB streams plus the exact serialized source used for its hash."""

    streams: Mapping[str, bytes]
    source_bytes: bytes = field(repr=False)
    source_sha256: str = field(init=False)

    def __post_init__(self):
        source = bytes(self.source_bytes)
        if len(source) > MAX_INPUT_BYTES:
            raise ProbeError("source exceeds the size policy")
        checked = _checked_streams(self.streams)
        object.__setattr__(self, "streams", MappingProxyType(checked))
        object.__setattr__(self, "source_bytes", source)
        object.__setattr__(self, "source_sha256", sha256(source).hexdigest())


def load_document_bytes(source, max_aggregate_bytes=MAX_AGGREGATE_STREAM_BYTES):
    """Load bounded CFB streams from already bounded serialized bytes."""
    source = bytes(source)
    if len(source) > MAX_INPUT_BYTES:
        raise ProbeError("source exceeds the size policy")
    if (type(max_aggregate_bytes) is not int or max_aggregate_bytes < 1
            or max_aggregate_bytes > MAX_AGGREGATE_STREAM_BYTES):
        raise ProbeError("invalid aggregate stream size policy")
    try:
        olefile = importlib.import_module("olefile")
    except ImportError as error:
        raise ProbeError("load_document requires the optional olefile package") from error
    handle = olefile.OleFileIO(BytesIO(source))
    try:
        streams = {}
        total = 0
        for parts in handle.listdir(streams=True, storages=False):
            name = "/".join(parts)
            if name in streams:
                raise ProbeError(f"duplicate CFB stream {name}")
            size = handle.get_size(parts)
            if size < 0 or size > min(MAX_STREAM_BYTES, max_aggregate_bytes):
                raise ProbeError(f"stream {name} exceeds the size policy")
            total += size
            if total > max_aggregate_bytes:
                raise ProbeError("aggregate streams exceed the size policy")
            value = handle.openstream(parts).read(size + 1)
            if len(value) != size:
                raise ProbeError(f"stream {name} length changed while reading")
            streams[name] = value
    finally:
        handle.close()
    return LoadedDocument(streams, source)


def load_document(path, max_bytes=MAX_INPUT_BYTES,
                  max_aggregate_bytes=MAX_AGGREGATE_STREAM_BYTES):
    """Load every CFB stream without making olefile an import-time dependency."""
    if type(max_bytes) is not int or max_bytes < 1 or max_bytes > MAX_INPUT_BYTES:
        raise ProbeError("invalid source size policy")
    path = Path(path)
    if path.stat().st_size > max_bytes:
        raise ProbeError("source exceeds the size policy")
    with path.open("rb") as source_file:
        source = source_file.read(max_bytes + 1)
    if len(source) > max_bytes:
        raise ProbeError("source exceeds the size policy")
    return load_document_bytes(source, max_aggregate_bytes=max_aggregate_bytes)


def _streams(value):
    if isinstance(value, LoadedDocument):
        return dict(value.streams)
    return _checked_streams(value)


def _range(data, start, size, message):
    if not isinstance(start, int) or not isinstance(size, int) or start < 0 or size < 0:
        raise ProbeError(message)
    end = start + size
    if end > len(data):
        raise ProbeError(message)
    return data[start:end]


def _u16(data, offset, message="truncated 16-bit value"):
    return int.from_bytes(_range(data, offset, 2, message), "little")


def _u32(data, offset, message="truncated 32-bit value"):
    return int.from_bytes(_range(data, offset, 4, message), "little")


class _WorkBudget:
    """Local diagnostic policy preventing input-to-output/work amplification."""

    def __init__(self, work_limit, payload_limit=0):
        self.work = 0
        self.records = 0
        self.payload_bytes = 0
        self.work_limit = work_limit
        self.payload_limit = payload_limit

    def charge(self, amount=1):
        self.work += amount
        if self.work > self.work_limit:
            raise ProbeError("diagnostic work budget exceeded")

    def record(self):
        self.records += 1
        if self.records > MAX_DIAGNOSTIC_RECORDS:
            raise ProbeError("diagnostic record budget exceeded")

    def payload(self, amount):
        self.payload_bytes += amount
        if self.payload_limit and self.payload_bytes > self.payload_limit:
            raise ProbeError("diagnostic payload budget exceeded")


class _TraceBudget:
    def __init__(self, work=None):
        self.prls = 0
        self.bytes = 0
        self.work = work

    def array(self, size):
        self.bytes += size
        if self.bytes > MAX_TRACE_BYTES:
            raise ProbeError("property trace byte budget exceeded")
        if self.work is not None:
            self.work.charge(size)

    def prl(self):
        self.prls += 1
        if self.prls > MAX_TRACE_PRLS:
            raise ProbeError("property trace SPRM budget exceeded")
        if self.work is not None:
            self.work.record()


def _read_prl(data, position, end, stream, budget):
    budget.prl()
    code = _u16(data, position, "truncated SPRM code")
    operand_start = position + 2
    spra = code >> 13
    if spra in (0, 1):
        size = 1
    elif spra in (2, 4, 5):
        size = 2
    elif spra == 3:
        size = 4
    elif spra == 7:
        size = 3
    elif code == 0xD608:
        cb = _u16(data, operand_start, "truncated TDefTable operand")
        if cb == 0:
            raise ProbeError("invalid TDefTable operand size")
        size = cb + 1
    elif code == 0xC615 and _range(data, operand_start, 1, "truncated tab operand")[0] == 0xFF:
        deleted = _range(data, operand_start + 1, 1, "truncated tab deletions")[0]
        added_offset = operand_start + 2 + deleted * 4
        added = _range(data, added_offset, 1, "truncated tab additions")[0]
        size = 3 + deleted * 4 + added * 3
    else:
        size = 1 + _range(data, operand_start, 1, "truncated variable SPRM operand")[0]
    operand_end = operand_start + size
    if operand_end > end:
        raise ProbeError("SPRM operand exceeds its property array")
    operand = data[operand_start:operand_end]
    if budget.work is not None:
        budget.work.payload(len(operand) * 2)
    return ({
        "kind": "prl",
        "stream": stream,
        "offset": position,
        "end": operand_end,
        "code": f"{code:04x}",
        "operand": operand.hex(),
        "applied": True,
    }, operand_end, code, operand)


def _trace_properties(streams, stream, start, end, budget,
                      filter_initial_paragraph=False):
    if stream not in streams:
        raise ProbeError(f"missing stream {stream}")
    _range(streams[stream], start, end - start, "property range outside stream")
    trace = []
    current_stream, current_start, current_end = stream, start, end
    visited = set()
    filter_paragraph = filter_initial_paragraph
    while True:
        data = streams[current_stream]
        budget.array(current_end - current_start)
        position = current_start
        first = True
        reference = None
        while position < current_end:
            item, next_position, code, operand = _read_prl(
                data, position, current_end, current_stream, budget
            )
            is_paragraph = ((code >> 10) & 7) == 1
            if filter_paragraph and not is_paragraph:
                item["applied"] = False
                item["reason"] = "non-paragraph-complex-PCD"
            if code == PHUGE_PAPX and not first:
                item["applied"] = False
                item["reason"] = "non-first-PHugePapx"
            trace.append(item)
            if (not filter_paragraph or is_paragraph) and (
                    code == P_TABLE_PROPS or (code == PHUGE_PAPX and first)):
                reference = _u32(operand, 0, "short paragraph data reference")
                item["followed"] = True
                if next_position < current_end:
                    if budget.work is not None:
                        budget.work.record()
                        budget.work.payload((current_end - next_position) * 2)
                    trace.append({
                        "kind": "ignored_tail",
                        "stream": current_stream,
                        "offset": next_position,
                        "end": current_end,
                        "bytes": data[next_position:current_end].hex(),
                        "reason": (
                            "after-PTableProps"
                            if code == P_TABLE_PROPS
                            else "after-first-PHugePapx"
                        ),
                    })
                break
            position = next_position
            if not filter_paragraph or is_paragraph:
                first = False
        if reference is None:
            return trace
        if reference in visited:
            raise ProbeError("cyclic paragraph data chain")
        if len(visited) >= MAX_TRACE_DEPTH:
            raise ProbeError("paragraph data chain exceeds depth policy")
        visited.add(reference)
        data = streams.get("Data")
        if data is None:
            raise ProbeError("paragraph data reference requires Data stream")
        size = _u16(data, reference, "paragraph data offset outside Data stream")
        if size < MIN_INDIRECT_GRPPRL_BYTES:
            raise ProbeError("short PrcData grpprl")
        if size > MAX_PRC_GRPPRL_BYTES:
            raise ProbeError("oversized PrcData grpprl")
        group_start = reference + 2
        group_end = group_start + size
        _range(data, group_start, size, "PrcData grpprl outside Data stream")
        if budget.work is not None:
            budget.work.record()
        trace.append({
            "kind": "prc_data",
            "stream": "Data",
            "offset": reference,
            "end": group_end,
            "grpprl_offset": group_start,
            "grpprl_end": group_end,
        })
        current_stream, current_start, current_end = "Data", group_start, group_end
        filter_paragraph = False


def _trace_piece_paragraph_properties(streams, stream, start, end, budget):
    """Trace a complex Pcd.Prm1 after its normative paragraph-SPRM filter."""
    return _trace_properties(
        streams, stream, start, end, budget, filter_initial_paragraph=True
    )


def trace_properties(streams, stream, start, end):
    """Return ordered provenance while following bounded paragraph indirection."""
    streams = _streams(streams)
    return _trace_properties(streams, stream, start, end, _TraceBudget())


def _fib(streams):
    word = streams.get("WordDocument")
    if word is None or len(word) < LCB_CLX_OFFSET + 4:
        raise ProbeError("missing or truncated WordDocument FIB")
    if _u16(word, 0) != FIB_IDENT or _u16(word, 2) < FIB_WORD_97:
        raise ProbeError("only Word 97-2003 documents are supported")
    flags = _u16(word, FIB_FLAGS_OFFSET)
    if flags & 0x8100:
        raise ProbeError("encrypted Word documents are unsupported")
    table_stream = "1Table" if flags & 0x0200 else "0Table"
    if table_stream not in streams:
        raise ProbeError(f"missing selected table stream {table_stream}")
    return {
        "table_stream": table_stream,
        "ccp_text": _u32(word, CCP_TEXT_OFFSET),
        "fc_plcf_bte_papx": _u32(word, FC_PLCF_BTE_PAPX_OFFSET),
        "lcb_plcf_bte_papx": _u32(word, LCB_PLCF_BTE_PAPX_OFFSET),
        "fc_clx": _u32(word, FC_CLX_OFFSET),
        "lcb_clx": _u32(word, LCB_CLX_OFFSET),
    }


def _papx_runs(streams, fib, work, include_hex):
    word = streams["WordDocument"]
    table = streams[fib["table_stream"]]
    plc = _range(
        table,
        fib["fc_plcf_bte_papx"],
        fib["lcb_plcf_bte_papx"],
        "PAPX page table outside selected table stream",
    )
    if not plc:
        return []
    if len(plc) < 4 or (len(plc) - 4) % 8:
        raise ProbeError("invalid PAPX page table")
    page_count = (len(plc) - 4) // 8
    if page_count > MAX_FKP_PAGES:
        raise ProbeError("PAPX page count exceeds policy")
    runs = []
    for page_index in range(page_count):
        lower = _u32(plc, page_index * 4)
        upper = _u32(plc, page_index * 4 + 4)
        if lower >= upper:
            raise ProbeError("unordered PAPX page range")
        pn_offset = (page_count + 1) * 4 + page_index * 4
        page_number = _u32(plc, pn_offset) & 0x003FFFFF
        page_base = page_number * 512
        page = _range(word, page_base, 512, "PAP FKP outside WordDocument")
        count = page[511]
        if count < 1 or count > 29 or len(runs) + count > MAX_PAPX_RUNS:
            raise ProbeError("invalid or excessive PAPX run count")
        pointers = (count + 1) * 4
        payload_start = pointers + count * 13
        for index in range(count):
            work.record()
            fc_start = _u32(page, index * 4)
            fc_end = _u32(page, index * 4 + 4)
            if fc_start >= fc_end or fc_start < lower or fc_end > upper:
                raise ProbeError("invalid PAPX physical range")
            payload = page[pointers + index * 13] * 2
            raw = None
            if payload:
                if payload < payload_start or payload >= 511:
                    raise ProbeError("PAPX payload overlaps FKP index")
                cb = page[payload]
                if cb:
                    encoding = "short"
                    header_offset = page_base + payload
                    raw_start = payload + 1
                    raw_size = cb * 2 - 1
                else:
                    encoding = "extended"
                    cb_prime = _range(page, payload + 1, 1, "truncated extended PAPX")[0]
                    if cb_prime < 1:
                        raise ProbeError("invalid extended PAPX size")
                    header_offset = page_base + payload
                    raw_start = payload + 2
                    raw_size = cb_prime * 2
                if raw_start + raw_size > 511:
                    raise ProbeError("PAPX exceeds FKP payload area")
                raw_bytes = page[raw_start:raw_start + raw_size]
                if len(raw_bytes) < 2:
                    raise ProbeError("PAPX lacks GrpPrlAndIstd style")
                absolute = page_base + raw_start
                raw = {
                    "stream": "WordDocument",
                    "offset": absolute,
                    "end": absolute + raw_size,
                    "header_offset": header_offset,
                    "encoding": encoding,
                    "istd": _u16(raw_bytes, 0),
                }
                if include_hex:
                    work.payload(raw_size * 2)
                    raw["bytes"] = raw_bytes.hex()
            runs.append({"fc_start": fc_start, "fc_end": fc_end, "papx": raw})
    return runs


def _clx(streams, fib, work, include_hex):
    table_stream = fib["table_stream"]
    table = streams[table_stream]
    clx_start = fib["fc_clx"]
    clx = _range(table, clx_start, fib["lcb_clx"], "CLX outside selected table stream")
    position = 0
    prcs = []
    while position < len(clx) and clx[position] == 0x01:
        work.record()
        size = _u16(clx, position + 1, "truncated CLX property record")
        if size > MAX_PRC_GRPPRL_BYTES or len(prcs) >= MAX_PRCS:
            raise ProbeError("invalid or excessive CLX property records")
        start = position + 3
        end = start + size
        _range(clx, start, size, "truncated CLX property record")
        prc = {
            "stream": table_stream,
            "offset": clx_start + start,
            "end": clx_start + end,
        }
        if include_hex:
            work.payload(size * 2)
            prc["bytes"] = clx[start:end].hex()
        prcs.append(prc)
        position = end
    if position >= len(clx) or clx[position] != 0x02:
        raise ProbeError("missing CLX piece table")
    plc_size = _u32(clx, position + 1, "truncated CLX piece table size")
    plc_start = position + 5
    plc = _range(clx, plc_start, plc_size, "truncated CLX piece table")
    if plc_size < 4 or (plc_size - 4) % 12:
        raise ProbeError("invalid CLX piece table size")
    count = (plc_size - 4) // 12
    if count < 1 or count > MAX_PIECES:
        raise ProbeError("invalid or excessive CLX piece count")
    cp_bytes = (count + 1) * 4
    pieces = []
    expected_cp = 0
    for index in range(count):
        work.record()
        cp_start = _u32(plc, index * 4)
        cp_end = _u32(plc, index * 4 + 4)
        if cp_start != expected_cp or cp_start > cp_end:
            raise ProbeError("non-contiguous CLX character ranges")
        expected_cp = cp_end
        pcd = cp_bytes + index * 8
        raw_fc = _u32(plc, pcd + 2)
        compressed = bool(raw_fc & 0x40000000)
        fc_start = (raw_fc & 0x3FFFFFFF) // 2 if compressed else raw_fc & 0x3FFFFFFF
        width = 1 if compressed else 2
        fc_end = fc_start + (cp_end - cp_start) * width
        if fc_end > len(streams["WordDocument"]):
            raise ProbeError("CLX physical range outside WordDocument")
        prm = _u16(plc, pcd + 6)
        complex_index = prm >> 1 if prm & 1 else None
        if complex_index is not None and complex_index >= len(prcs):
            raise ProbeError("complex PCD property index outside CLX")
        pieces.append({
            "cp_start": cp_start,
            "cp_end": cp_end,
            "fc_start": fc_start,
            "fc_end": fc_end,
            "compressed": compressed,
            "prm": f"{prm:04x}",
            "complex_index": complex_index,
            "stream": table_stream,
            "offset": clx_start + plc_start + pcd,
            "end": clx_start + plc_start + pcd + 8,
        })
    if expected_cp < fib["ccp_text"]:
        raise ProbeError("CLX does not cover the main story")
    return prcs, pieces


def _document_parts(streams, work=None, include_hex=False):
    streams = _streams(streams)
    if work is None:
        work = _WorkBudget(MAX_INSPECT_WORK, MAX_INSPECT_PAYLOAD_BYTES)
    fib = _fib(streams)
    prcs, pieces = _clx(streams, fib, work, include_hex)
    runs = _papx_runs(streams, fib, work, include_hex)
    return streams, fib, runs, prcs, pieces


def inspect_document(streams):
    """Inspect bounded, independent PAP FKP and CLX provenance records."""
    work = _WorkBudget(MAX_INSPECT_WORK, MAX_INSPECT_PAYLOAD_BYTES)
    streams, fib, runs, prcs, pieces = _document_parts(
        streams, work=work, include_hex=True
    )
    traces = []
    trace_ids = {}

    def add_trace(stream, start, end):
        key = (stream, start, end)
        if key in trace_ids:
            return trace_ids[key]
        entries = _trace_properties(streams, stream, start, end, _TraceBudget(work))
        trace_id = len(traces)
        trace_ids[key] = trace_id
        traces.append({
            "stream": stream,
            "offset": start,
            "end": end,
            "entries": entries,
        })
        return trace_id

    for run in runs:
        papx = run["papx"]
        if papx is None:
            run["direct_trace_id"] = None
        else:
            run["direct_trace_id"] = add_trace(
                "WordDocument", papx["offset"] + 2, papx["end"]
            )
    for prc in prcs:
        prc["trace_id"] = add_trace(
            prc["stream"], prc["offset"], prc["end"]
        )
    return {
        "table_stream": fib["table_stream"],
        "papx_runs": runs,
        "prcs": prcs,
        "pieces": pieces,
        "property_traces": traces,
    }


def _target_code(value):
    if isinstance(value, str) and value in TARGET_CODES:
        return TARGET_CODES[value]
    if isinstance(value, int) and value in TARGET_OPERAND_SIZES:
        return value
    if isinstance(value, str) and re.fullmatch(r"(?:0x)?[0-9a-fA-F]{4}", value):
        code = int(value, 16)
        if code in TARGET_OPERAND_SIZES:
            return code
    raise ProbeError("target code must be TIstd or TTlp")


def _expected_operand(value, code):
    if value is None:
        return None
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-fA-F]*", value) or len(value) % 2:
        raise ProbeError("target operand must be hex or null")
    operand = bytes.fromhex(value)
    if len(operand) != TARGET_OPERAND_SIZES[code]:
        raise ProbeError("target operand has the wrong scalar size")
    return operand.hex()


def _active_operand(trace, code):
    matches = [
        item for item in trace
        if item.get("kind") == "prl"
        and item.get("applied")
        and int(item["code"], 16) == code
    ]
    return (None, None) if not matches else (matches[-1]["operand"], matches[-1])


def _paragraph_prm0_entry(piece, budget):
    """Decode one closed-table Prm0 for paragraph-property acquisition."""
    prm = int(piece["prm"], 16)
    if prm & 1:
        raise ProbeError("internal Prm0 decoder received a complex PCD PRM")
    isprm = (prm >> 1) & 0x7F
    value = prm >> 8
    if isprm == 0 and value == 0:
        return None
    budget.array(2)
    budget.prl()
    if budget.work is not None:
        budget.work.payload(2)
    code = PRM0_PARAGRAPH_SPRMS.get(isprm)
    applied = code is not None
    reason = None
    if not applied:
        code = PRM0_CHARACTER_SPRMS.get(isprm)
        if code is None:
            raise ProbeError("reserved or unknown Prm0 isprm")
        reason = "non-paragraph-Prm0"
    entry = {
        "kind": "prl",
        "stream": piece["stream"],
        # The canonical three-byte Prl below is synthesized from the actual
        # two-byte Pcd.Prm. This range therefore identifies origin, not Prl bytes.
        "offset": piece["offset"] + 6,
        "end": piece["offset"] + 8,
        "code": f"{code:04x}",
        "operand": f"{value:02x}",
        "applied": applied,
        "origin": "Pcd.Prm0",
        "framing": "synthetic-Prl-from-2-byte-Prm0",
        "prm": f"{prm:04x}",
        "isprm": f"{isprm:02x}",
    }
    if reason is not None:
        entry["reason"] = reason
    return entry


def _effective_target(parts, fc, code, trace_budget, lookup_budget, trace_cache):
    streams, _fib_data, all_runs, prcs, all_pieces = parts
    lookup_budget.charge(len(all_runs))
    runs = [run for run in all_runs if run["fc_start"] <= fc < run["fc_end"]]
    if len(runs) != 1:
        raise ProbeError("target FC does not select exactly one PAPX run")
    papx = runs[0]["papx"]
    trace = []
    if papx is not None:
        key = ("WordDocument", papx["offset"] + 2, papx["end"])
        if key not in trace_cache:
            trace_cache[key] = _trace_properties(streams, *key, trace_budget)
        trace.extend(trace_cache[key])

    lookup_budget.charge(len(all_pieces))
    pieces = [piece for piece in all_pieces if piece["fc_start"] <= fc < piece["fc_end"]]
    if len(pieces) != 1:
        raise ProbeError("target FC does not select exactly one CLX piece")
    piece = pieces[0]
    width = 1 if piece["compressed"] else 2
    if (fc - piece["fc_start"]) % width:
        raise ProbeError("target FC is not at a character boundary")
    marker = _range(streams["WordDocument"], fc, width, "target character outside WordDocument")
    if marker != (b"\x07" if width == 1 else b"\x07\x00"):
        raise ProbeError("target FC is not a top-level TTP character")
    index = piece["complex_index"]
    if index is not None:
        prc = prcs[index]
        key = ("complexPCD", prc["stream"], prc["offset"], prc["end"])
        if key not in trace_cache:
            trace_cache[key] = _trace_piece_paragraph_properties(
                streams, *key[1:], trace_budget
            )
        trace.extend(trace_cache[key])

    depth, _ = _active_operand(trace, 0x6649)
    row_end, _ = _active_operand(trace, 0x2417)
    if depth != "01000000" or row_end != "01":
        raise ProbeError("target FC lacks effective top-level TTP properties")
    return _active_operand(trace, code)


def _strict_keys(value, required, optional, label):
    if not isinstance(value, Mapping):
        raise ProbeError(f"{label} must be an object")
    keys = set(value)
    if required - keys:
        raise ProbeError(f"{label} is missing required fields")
    if keys - required - optional:
        raise ProbeError(f"{label} contains unknown fields")


def _trace_code(value, label="property code"):
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-f]{4}", value):
        raise ProbeError(f"{label} must be four lowercase hexadecimal digits")
    return int(value, 16)


class AcquiredPropertyTraceSession:
    """Bounded per-document lookup and trace cache for acquired PAP properties."""

    def __init__(self, document):
        if not isinstance(document, LoadedDocument):
            raise ProbeError("LoadedDocument is required for acquired property traces")
        self.document = document
        self.parts = _document_parts(document)
        self.streams, _fib_data, self.runs, self.prcs, self.pieces = self.parts
        self.run_index = self._interval_index(self.runs)
        self.piece_index = self._interval_index(self.pieces)
        self.trace_budget = _TraceBudget(
            _WorkBudget(MAX_INSPECT_WORK, MAX_INSPECT_PAYLOAD_BYTES)
        )
        self.lookup_budget = _WorkBudget(MAX_TARGET_LOOKUP_WORK)
        self.acquire_budget = _WorkBudget(MAX_INSPECT_WORK)
        self.trace_cache = {}

    @staticmethod
    def _interval_index(items):
        ordered = sorted(items, key=lambda item: (item["fc_start"], item["fc_end"]))
        size = 1
        while size < len(ordered):
            size *= 2
        maxima = [-1] * (size * 2)
        for index, item in enumerate(ordered):
            maxima[size + index] = item["fc_end"]
        for index in range(size - 1, 0, -1):
            maxima[index] = max(maxima[index * 2], maxima[index * 2 + 1])
        return ordered, [item["fc_start"] for item in ordered], maxima, size

    def _containing(self, interval_index, fc, label):
        items, starts, maxima, size = interval_index
        upper = bisect_right(starts, fc)
        matches = []

        def visit(node, left, right):
            self.lookup_budget.charge(1)
            if left >= upper or maxima[node] <= fc or len(matches) > 1:
                return
            if right - left == 1:
                if left < len(items) and items[left]["fc_start"] <= fc < items[left]["fc_end"]:
                    matches.append(items[left])
                return
            middle = (left + right) // 2
            visit(node * 2, left, middle)
            visit(node * 2 + 1, middle, right)

        if items:
            visit(1, 0, size)
        if len(matches) != 1:
            raise ProbeError(f"target FC does not select exactly one {label}")
        return matches[0]

    def _trace(self, key):
        if key not in self.trace_cache:
            self.trace_cache[key] = _trace_properties(
                self.streams, *key, self.trace_budget
            )
        return self.trace_cache[key]

    def _piece_trace(self, key):
        cache_key = ("complexPCD", *key)
        if cache_key not in self.trace_cache:
            self.trace_cache[cache_key] = _trace_piece_paragraph_properties(
                self.streams, *key, self.trace_budget
            )
        return self.trace_cache[cache_key]

    def acquire(self, fc, owner):
        if type(fc) is not int or fc < 0:
            raise ProbeError("target fc must be a nonnegative integer physical FC")
        if owner not in ("ttp", "paragraph"):
            raise ProbeError("target owner must be ttp or paragraph")
        run = self._containing(self.run_index, fc, "PAPX run")
        piece = self._containing(self.piece_index, fc, "CLX piece")
        width = 1 if piece["compressed"] else 2
        if (fc - piece["fc_start"]) % width:
            raise ProbeError("target FC is not at a character boundary")
        marker = _range(
            self.streams["WordDocument"], fc, width,
            "target character outside WordDocument",
        )
        trace = []
        papx = run["papx"]
        if papx is not None:
            trace.extend(self._trace(("WordDocument", papx["offset"] + 2, papx["end"])))
        if piece["complex_index"] is not None:
            prc = self.prcs[piece["complex_index"]]
            trace.extend(self._piece_trace((prc["stream"], prc["offset"], prc["end"])))
        else:
            prm0_entry = _paragraph_prm0_entry(piece, self.trace_budget)
            if prm0_entry is not None:
                trace.append(prm0_entry)
        self.acquire_budget.charge(len(trace))

        depth, _ = _active_operand(trace, 0x6649)
        row_end, _ = _active_operand(trace, 0x2417)
        in_table, _ = _active_operand(trace, 0x2416)
        inner_ttp, _ = _active_operand(trace, 0x244C)
        is_u07 = marker == (b"\x07" if width == 1 else b"\x07\x00")
        is_ttp = (
            is_u07 and depth == "01000000"
            and row_end == "01" and in_table == "01"
        )
        if owner == "ttp" and not is_ttp:
            raise ProbeError("target FC lacks acquired top-level TTP ownership")
        has_ttp_marker = row_end == "01" or (
            marker == (b"\x0d" if width == 1 else b"\x0d\x00")
            and inner_ttp == "01"
        )
        if owner == "paragraph" and has_ttp_marker:
            raise ProbeError("paragraph owner rejects a TTP character")
        if owner == "paragraph" and depth not in (None, "00000000") and in_table != "01":
            raise ProbeError("paragraph target has inconsistent acquired in-table ownership")
        return {
            "fc": fc,
            "owner": owner,
            "character": marker.hex(),
            "entries": [dict(item) for item in trace],
        }


def acquire_property_trace(document, fc, owner):
    """Acquire FKP/Data then PCD serialized property provenance for one owner."""
    session = (
        document if isinstance(document, AcquiredPropertyTraceSession)
        else AcquiredPropertyTraceSession(document)
    )
    return session.acquire(fc, owner)


def validate_trace_assertions(document, assertion_manifest, session=None):
    """Validate hash-bound, read-only acquired serialized property assertions."""
    if not isinstance(document, LoadedDocument):
        raise ProbeError("LoadedDocument is required to validate source_sha256")
    _strict_keys(
        assertion_manifest, {"schema", "source_sha256", "targets"}, set(),
        "trace assertion manifest",
    )
    if assertion_manifest["schema"] != TRACE_ASSERTION_SCHEMA:
        raise ProbeError("unsupported trace assertion schema")
    expected_hash = assertion_manifest["source_sha256"]
    if (not isinstance(expected_hash, str)
            or not re.fullmatch(r"[0-9a-f]{64}", expected_hash)
            or expected_hash != document.source_sha256):
        raise ProbeError("source_sha256 does not match the serialized source")
    targets = assertion_manifest["targets"]
    if not isinstance(targets, list) or not targets:
        raise ProbeError("trace assertion manifest requires targets")
    if len(targets) > MAX_TARGETS:
        raise ProbeError("target count exceeds policy")
    if session is None:
        context = AcquiredPropertyTraceSession(document)
    elif (not isinstance(session, AcquiredPropertyTraceSession)
            or session.document is not document):
        raise ProbeError("trace session is not bound to this LoadedDocument")
    else:
        context = session
    seen = set()
    results = []
    retained = 0
    for target in targets:
        _strict_keys(target, {"fc", "owner", "properties"}, {"order"}, "trace target")
        fc = target["fc"]
        owner = target["owner"]
        if type(fc) is not int or fc < 0:
            raise ProbeError("target fc must be a nonnegative integer physical FC")
        if owner not in ("ttp", "paragraph"):
            raise ProbeError("target owner must be ttp or paragraph")
        identity = (fc, owner)
        if identity in seen:
            raise ProbeError("duplicate trace target")
        seen.add(identity)
        properties = target["properties"]
        if not isinstance(properties, Mapping) or not properties:
            raise ProbeError("trace target properties must be a nonempty object")
        expected = {}
        retained += 64
        for code_text, operands in properties.items():
            context.acquire_budget.charge(1)
            _trace_code(code_text)
            retained += len(code_text) + 8
            if retained > MAX_TRACE_BYTES:
                raise ProbeError("trace assertion retained output exceeds policy")
            if not isinstance(operands, list):
                raise ProbeError("property assertion must be an operand array")
            normalized = []
            for operand in operands:
                context.acquire_budget.charge(1)
                if (not isinstance(operand, str)
                        or not re.fullmatch(r"(?:[0-9a-f]{2})*", operand)):
                    raise ProbeError("asserted operand must be lowercase whole-byte hex")
                if retained + len(operand) + 8 > MAX_TRACE_BYTES:
                    raise ProbeError("trace assertion retained output exceeds policy")
                operand_bytes = bytes.fromhex(operand)
                probe_budget = _TraceBudget()
                try:
                    _item, end, parsed_code, parsed_operand = _read_prl(
                        int(code_text, 16).to_bytes(2, "little") + operand_bytes,
                        0, 2 + len(operand_bytes), "assertion", probe_budget,
                    )
                except ProbeError as error:
                    raise ProbeError("asserted operand does not match SPRM framing") from error
                if (end != 2 + len(operand_bytes)
                        or parsed_code != int(code_text, 16)
                        or parsed_operand != operand_bytes):
                    raise ProbeError("asserted operand does not match SPRM framing")
                retained += len(operand) + 8
                normalized.append(operand)
            expected[code_text] = normalized
        order = target.get("order")
        if order is not None:
            if not isinstance(order, list):
                raise ProbeError("trace target order must be an array")
            if len(order) > MAX_TRACE_PRLS:
                raise ProbeError("trace target order exceeds policy")
            for code_text in order:
                context.acquire_budget.charge(1)
                _trace_code(code_text, "order code")
                if code_text not in expected:
                    raise ProbeError("order may contain only asserted property codes")
                retained += len(code_text) + 4
                if retained > MAX_TRACE_BYTES:
                    raise ProbeError("trace assertion retained output exceeds policy")
        acquired = context.acquire(fc, owner)
        context.acquire_budget.charge(len(acquired["entries"]))
        applied = [
            item for item in acquired["entries"]
            if item.get("kind") == "prl" and item.get("applied") is True
        ]
        actual = {code_text: [] for code_text in expected}
        actual_order = []
        for item in applied:
            code_text = item["code"]
            if code_text in actual:
                actual[code_text].append(item["operand"])
                actual_order.append(code_text)
        if actual != expected:
            raise ProbeError("acquired property operand assertion does not match")
        if order is not None and actual_order != order:
            raise ProbeError("acquired property order assertion does not match")
        results.append({
            "fc": fc,
            "owner": owner,
            "properties": actual,
            **({"order": actual_order} if order is not None else {}),
        })
    return {
        "valid": True,
        "schema": TRACE_ASSERTION_SCHEMA,
        "scope": "acquired serialized property trace only; not Word property-family precedence",
        "source_sha256": document.source_sha256,
        "targets": results,
    }


def _hex_edit(value, field_name):
    if not isinstance(value, str) or not value or not re.fullmatch(r"[0-9a-fA-F]+", value):
        raise ProbeError(f"edit {field_name} must be nonempty hex")
    if len(value) % 2:
        raise ProbeError(f"edit {field_name} must contain whole bytes")
    return bytes.fromhex(value)


def validate_plan(before_streams, after_streams, plan):
    """Validate exact same-length edits and effective TIstd/TTlp assertions."""
    if not isinstance(before_streams, LoadedDocument):
        raise ProbeError("source bytes are required to validate source_sha256")
    before = _streams(before_streams)
    after = _streams(after_streams)
    if not isinstance(plan, Mapping):
        raise ProbeError("plan must be an object")
    expected_hash = plan.get("source_sha256")
    actual_hash = sha256(before_streams.source_bytes).hexdigest()
    if expected_hash != actual_hash:
        raise ProbeError("source_sha256 does not match the serialized source")
    if set(before) != set(after):
        raise ProbeError("candidate stream set differs from source")

    edits = plan.get("edits")
    targets = plan.get("targets")
    if not isinstance(edits, list) or not edits:
        raise ProbeError("plan requires edits")
    if not isinstance(targets, list) or not targets:
        raise ProbeError("plan requires target assertions")
    if len(edits) > MAX_EDITS:
        raise ProbeError("edit count exceeds policy")
    if len(targets) > MAX_TARGETS:
        raise ProbeError("target count exceeds policy")
    ranges = {}
    normalized_edits = []
    edit_bytes = 0
    for edit in edits:
        if not isinstance(edit, Mapping):
            raise ProbeError("edit must be an object")
        stream = edit.get("stream")
        offset = edit.get("offset")
        if stream not in before or type(offset) is not int or offset < 0:
            raise ProbeError("edit selects an invalid stream offset")
        old = _hex_edit(edit.get("before"), "before")
        new = _hex_edit(edit.get("after"), "after")
        if len(old) != len(new) or old == new:
            raise ProbeError("edits must change a same-length byte range")
        edit_bytes += len(old)
        if edit_bytes > MAX_PLAN_EDIT_BYTES:
            raise ProbeError("plan edit bytes exceed policy")
        end = offset + len(old)
        if end > len(before[stream]) or before[stream][offset:end] != old:
            raise ProbeError("edit before bytes do not match source")
        if len(after[stream]) != len(before[stream]) or after[stream][offset:end] != new:
            raise ProbeError("edit after bytes do not match candidate")
        stream_ranges = ranges.setdefault(stream, [])
        stream_ranges.append((offset, end))
        normalized_edits.append({"stream": stream, "offset": offset, "end": end})
    for stream_ranges in ranges.values():
        stream_ranges.sort()
        if any(left[1] > right[0] for left, right in zip(stream_ranges, stream_ranges[1:])):
            raise ProbeError("edit ranges overlap")
    for stream in before:
        if len(before[stream]) != len(after[stream]):
            raise ProbeError("candidate stream length differs from source")
        stream_ranges = ranges.get(stream, [])
        range_index = 0
        for offset, (old, new) in enumerate(zip(before[stream], after[stream])):
            if old == new:
                continue
            while range_index < len(stream_ranges) and stream_ranges[range_index][1] <= offset:
                range_index += 1
            if range_index >= len(stream_ranges) or not (
                stream_ranges[range_index][0] <= offset < stream_ranges[range_index][1]
            ):
                raise ProbeError("candidate contains an undeclared stream change")

    before_parts = _document_parts(before)
    after_parts = _document_parts(after)
    before_trace_budget = _TraceBudget(
        _WorkBudget(MAX_INSPECT_WORK, MAX_INSPECT_PAYLOAD_BYTES)
    )
    after_trace_budget = _TraceBudget(
        _WorkBudget(MAX_INSPECT_WORK, MAX_INSPECT_PAYLOAD_BYTES)
    )
    before_lookup_budget = _WorkBudget(MAX_TARGET_LOOKUP_WORK)
    after_lookup_budget = _WorkBudget(MAX_TARGET_LOOKUP_WORK)
    before_trace_cache = {}
    after_trace_cache = {}
    target_results = []
    seen_targets = set()
    for target in targets:
        if (not isinstance(target, Mapping)
                or type(target.get("fc")) is not int
                or "before" not in target or "after" not in target):
            raise ProbeError("target requires an integer physical FC and before/after fields")
        fc = target["fc"]
        code = _target_code(target.get("code"))
        identity = (fc, code)
        if identity in seen_targets:
            raise ProbeError("duplicate target assertion")
        seen_targets.add(identity)
        expected_before = _expected_operand(target.get("before"), code)
        expected_after = _expected_operand(target.get("after"), code)
        if expected_before == expected_after:
            raise ProbeError("target assertion must describe an effective change")
        actual_before, before_source = _effective_target(
            before_parts, fc, code, before_trace_budget,
            before_lookup_budget, before_trace_cache,
        )
        actual_after, after_source = _effective_target(
            after_parts, fc, code, after_trace_budget,
            after_lookup_budget, after_trace_cache,
        )
        if actual_before != expected_before or actual_after != expected_after:
            raise ProbeError("effective target assertion does not match PAPX/Data/PCD cascade")
        if (code == TARGET_CODES["TTlp"]
                and ("0000" if actual_before is None else actual_before[4:])
                == ("0000" if actual_after is None else actual_after[4:])):
            raise ProbeError("TTlp target changes only historical itl")
        target_results.append({
            "fc": fc,
            "code": f"{code:04x}",
            "before": actual_before,
            "after": actual_after,
            "before_source": before_source,
            "after_source": after_source,
        })
    return {
        "valid": True,
        "scope": "acquired top-level TTP scalar provenance only; not display or style validity",
        "source_sha256": actual_hash,
        "edits": normalized_edits,
        "targets": target_results,
    }


def _load_plan(path):
    if path.stat().st_size > MAX_PLAN_BYTES:
        raise ProbeError("plan exceeds the size policy")
    def no_duplicate_keys(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ProbeError(f"duplicate JSON field {key}")
            result[key] = value
        return result

    try:
        return json.loads(
            path.read_text(encoding="utf-8"), object_pairs_hook=no_duplicate_keys
        )
    except json.JSONDecodeError as error:
        raise ProbeError("invalid JSON plan") from error


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    inspect = commands.add_parser("inspect", help="inspect a binary DOC")
    inspect.add_argument("source", type=Path)
    validate = commands.add_parser("validate", help="validate a candidate against an exact plan")
    validate.add_argument("source", type=Path)
    validate.add_argument("candidate", type=Path)
    validate.add_argument("plan", type=Path)
    check_trace = commands.add_parser(
        "check-trace", help="check hash-bound acquired property trace assertions"
    )
    check_trace.add_argument("source", type=Path)
    check_trace.add_argument("assertions", type=Path)
    args = parser.parse_args(argv)
    if args.command == "inspect":
        loaded = load_document(args.source)
        result = inspect_document(loaded)
        result["source_sha256"] = loaded.source_sha256
    elif args.command == "validate":
        source = load_document(args.source)
        candidate = load_document(args.candidate)
        plan = _load_plan(args.plan)
        result = validate_plan(source, candidate, plan)
    else:
        source = load_document(args.source)
        assertions = _load_plan(args.assertions)
        result = validate_trace_assertions(source, assertions)
    print(json.dumps(result, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
