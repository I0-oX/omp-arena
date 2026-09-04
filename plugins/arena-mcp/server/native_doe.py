"""Linux-native .doe reader (no Arena, no COM, no Windows).

A real Arena `.doe` is an OLE2 compound file whose `Contents` stream is a
proprietary MFC-serialized binary. Fully decoding modules/operands needs
Arena itself (see the COM tools). What this module *can* do natively:

- document header (document/product version, generation date) via regex,
- OLE metadata (SummaryInformation) via olefile,
- stream inventory,
- VBA macro source via oletools (optional dependency),
- printable-string inventory with Arena module-definition hits
  (approximation for triage/search — NOT a module list).

Only dependency nuances: `olefile` required; `oletools` optional (VBA section
degrades to ``available: false`` without it).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from arena_extractor import (
    ARENA_PROG_ID,
    ASSISTED_MODULES,
    AUTOMATIC_MODULES,
    MODEL_EXTENSIONS,
    ArenaExtractorError,
    _json_value,
    _resolve_allowed_file,
)

try:
    import olefile
except ImportError:
    olefile = None  # type: ignore

NATIVE_SCHEMA_VERSION = "1.0.0"
DEFAULT_MAX_STRINGS = 500
DEFAULT_MAX_STRING_CHARS = 100_000
DEFAULT_MAX_VBA_CHARS = 200_000

_HEADER_PATTERNS = {
    "document_type": rb"Document Type:\s*([ -~]{1,16})",
    "document_version": rb"Document Version:\s*(\d{1,8})",
    "product_version": rb"Product Version:\s*([0-9.]{1,16})",
    "generation_date": rb"Generation Date:\s*([0-9/ :]{4,20})",
    "copyright_year": rb"Copyright \(C\) Rockwell Software,\s*(\d{4})",
}


def _extract_ascii_strings(data: bytes, min_len: int = 5) -> list[str]:
    """Printable-ASCII runs, in file order (pure function, no I/O)."""
    if min_len < 1:
        raise ValueError("min_len must be >= 1.")
    out: list[str] = []
    cur = bytearray()
    for byte in data:
        if 32 <= byte < 127:
            cur.append(byte)
        else:
            if len(cur) >= min_len:
                out.append(bytes(cur).decode("ascii"))
            cur = bytearray()
    if len(cur) >= min_len:
        out.append(bytes(cur).decode("ascii"))
    return out


def _parse_doe_header(data: bytes) -> dict[str, Any]:
    """Document header fields from raw bytes (pure function, no I/O)."""
    header: dict[str, Any] = {}
    for key, pattern in _HEADER_PATTERNS.items():
        match = re.search(pattern, data)
        header[key] = (
            match.group(1).decode("ascii", "replace").strip() if match else None
        )
    return header


def _module_hits(strings: list[str]) -> dict[str, Any]:
    """Count exact module-definition name occurrences (triage approximation)."""
    automatic: dict[str, int] = {}
    assisted: dict[str, int] = {}
    for text in strings:
        if text in AUTOMATIC_MODULES:
            automatic[text] = automatic.get(text, 0) + 1
        elif text in ASSISTED_MODULES:
            assisted[text] = assisted.get(text, 0) + 1
    return {
        "automatic": automatic,
        "assisted": assisted,
        "note": (
            "Occurrence counts in printable strings, not a module list. "
            "Authoritative module data needs Arena COM (extract_arena_model)."
        ),
    }


def _decode_meta(value: Any) -> Any:
    if isinstance(value, (bytes, bytearray)):
        for encoding in ("utf-8", "cp1252"):
            try:
                return bytes(value).decode(encoding).strip("\x00").strip()
            except (UnicodeDecodeError, ValueError):
                continue
    return _json_value(value)


def _ole_metadata(path: Path) -> dict[str, Any]:
    assert olefile is not None
    with olefile.OleFileIO(str(path)) as ole:
        meta = ole.get_metadata()
        return {
            name: _decode_meta(getattr(meta, name, None))
            for name in (
                "title",
                "author",
                "subject",
                "keywords",
                "comments",
                "creating_application",
                "codepage",
            )
        }


def _stream_inventory(path: Path) -> list[dict[str, Any]]:
    assert olefile is not None
    entries: list[dict[str, Any]] = []
    with olefile.OleFileIO(str(path)) as ole:
        for entry in ole.direntries:
            if entry is None or entry.entry_type != 2:  # streams only
                continue
            entries.append({"name": entry.name, "size": entry.size})
    entries.sort(key=lambda item: item["size"], reverse=True)
    return entries


def _extract_vba(path: Path, max_chars: int) -> dict[str, Any]:
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        return {
            "available": False,
            "reason": "oletools not installed (`pip install oletools`).",
            "macros": [],
        }
    parser = VBA_Parser(str(path))
    try:
        if not parser.detect_vba_macros():
            return {"available": True, "macro_count": 0, "macros": []}
        macros: list[dict[str, Any]] = []
        used = 0
        truncated = False
        for _filename, _stream_path, vba_filename, vba_code in parser.extract_macros():
            code = vba_code or ""
            if used + len(code) > max_chars:
                code = code[: max(0, max_chars - used)]
                truncated = True
            used += len(code)
            macros.append({"name": vba_filename, "chars": len(code), "code": code})
            if truncated:
                break
        return {
            "available": True,
            "macro_count": len(macros),
            "truncated": truncated,
            "macros": macros,
        }
    finally:
        parser.close()


def inspect_native_doe(
    model_path: str,
    include_strings: bool = True,
    include_vba: bool = True,
    max_strings: int = DEFAULT_MAX_STRINGS,
    max_string_chars: int = DEFAULT_MAX_STRING_CHARS,
    max_vba_chars: int = DEFAULT_MAX_VBA_CHARS,
) -> dict[str, Any]:
    """Read everything Linux-native out of a .doe (header, metadata, VBA, strings)."""
    if max_strings < 0 or max_strings > 100_000:
        raise ValueError("max_strings must be between 0 and 100000.")
    if max_string_chars < 0 or max_vba_chars < 0:
        raise ValueError("char limits must be >= 0.")
    if olefile is None:
        raise ArenaExtractorError(
            "Reading .doe files natively requires `olefile` (`pip install olefile`)."
        )
    path = _resolve_allowed_file(model_path, MODEL_EXTENSIONS)
    if not olefile.isOleFile(str(path)):
        raise ArenaExtractorError(f"Not an OLE compound file: {path}")
    with olefile.OleFileIO(str(path)) as ole:
        contents = (
            ole.openstream("Contents").read() if ole.exists("Contents") else b""
        )
    header = _parse_doe_header(contents or path.read_bytes())
    strings = _extract_ascii_strings(contents) if include_strings else []
    kept: list[str] = []
    used = 0
    for text in strings:
        if len(kept) >= max_strings or used + len(text) > max_string_chars:
            break
        kept.append(text[:500])
        used += len(text)
    return {
        "native_schema": NATIVE_SCHEMA_VERSION,
        "prog_id": ARENA_PROG_ID,
        "path": str(path),
        "size_bytes": path.stat().st_size,
        "header": header,
        "metadata": _ole_metadata(path),
        "streams": _stream_inventory(path),
        "module_hits": _module_hits(strings),
        "strings": {
            "total": len(strings),
            "returned": len(kept),
            "truncated": len(kept) < len(strings),
            "sample": kept,
        },
        "vba": _extract_vba(path, max_vba_chars)
        if include_vba
        else {"available": False, "reason": "not requested", "macros": []},
    }
