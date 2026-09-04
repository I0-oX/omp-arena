"""Tests for the Linux-native .doe reader.

No real .doe needed (proprietary files are not committed): pure helpers run
on synthetic bytes, and bridge wiring is checked through clean error paths.
"""

import json
import subprocess
import sys
import unittest
from pathlib import Path

SERVER = Path(__file__).resolve().parent.parent / "plugins" / "arena-mcp" / "server"
sys.path.insert(0, str(SERVER))

from native_doe import (  # noqa: E402
    _extract_ascii_strings,
    _module_hits,
    _parse_doe_header,
)


class NativeHelperTests(unittest.TestCase):
    def test_ascii_strings_in_order(self):
        data = b"\x00Hello\x00\x01World!\x00ab\x00Trailing run here"
        self.assertEqual(
            _extract_ascii_strings(data, min_len=5),
            ["Hello", "World!", "Trailing run here"],
        )

    def test_ascii_strings_rejects_bad_min_len(self):
        with self.assertRaises(ValueError):
            _extract_ascii_strings(b"abc", min_len=0)

    def test_header_parses_arena_fields(self):
        blob = (
            b"Copyright (C) Rockwell Software, 2014\x00"
            b"Document Type:    .DOE\x00"
            b"Document Version: 563\x00"
            b"Product Version:  14.50.00\x00"
            b"Generation Date:  12/09/2014 20:11\x00"
        )
        self.assertEqual(
            _parse_doe_header(blob),
            {
                "document_type": ".DOE",
                "document_version": "563",
                "product_version": "14.50.00",
                "generation_date": "12/09/2014 20:11",
                "copyright_year": "2014",
            },
        )

    def test_header_missing_fields_are_none(self):
        header = _parse_doe_header(b"nothing relevant here")
        self.assertTrue(all(value is None for value in header.values()))

    def test_module_hits_counts_exact_names(self):
        hits = _module_hits(["Create", "Create", "Route", "Belt1", "Process"])
        self.assertEqual(hits["automatic"], {"Create": 2, "Process": 1})
        self.assertEqual(hits["assisted"], {"Route": 1})
        self.assertIn("not a module list", hits["note"])


class NativeBridgeTests(unittest.TestCase):
    def _bridge(self, tool, args):
        proc = subprocess.run(
            [sys.executable, str(SERVER / "omp_bridge.py"), "--json-call",
             json.dumps({"tool": tool, "args": args})],
            capture_output=True, text=True, timeout=120,
        )
        return json.loads(proc.stdout)

    def test_native_tool_is_dispatched(self):
        payload = self._bridge("inspect_arena_native", {"model_path": "/tmp/no.doe"})
        self.assertFalse(payload["ok"])
        self.assertIn("File not found", payload["error"])

    def test_native_tool_rejects_bad_args(self):
        payload = self._bridge("inspect_arena_native", {"bogus": 1})
        self.assertFalse(payload["ok"])
        self.assertIn("Unexpected argument", payload["error"])

    def test_unknown_tool_lists_native(self):
        payload = self._bridge("nope", {})
        self.assertFalse(payload["ok"])
        self.assertIn("inspect_arena_native", payload["error"])


if __name__ == "__main__":
    unittest.main()
