#!/usr/bin/env python3
"""Exit-code checks for the COM-free deck validator."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VALIDATOR = ROOT / "validate-deck.py"
EXAMPLES = (
    ROOT / "examples" / "network-diagram.json",
    ROOT / "examples" / "image-asset-demo.json",
    ROOT / "examples" / "consulting-deck.json",
    ROOT / "examples" / "research-brief.json",
    ROOT / "examples" / "product-overview.json",
    ROOT / "tests" / "fixtures" / "valid" / "object-assets.json",
)


def run_validator(*specs: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(VALIDATOR), *(str(spec) for spec in specs)],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )


class ValidateDeckTests(unittest.TestCase):
    def test_validator_does_not_reference_powerpoint_com(self) -> None:
        source = VALIDATOR.read_text(encoding="utf-8")
        for forbidden in ("PowerPoint.Application", "win32com", "comtypes", "New-Object"):
            self.assertNotIn(forbidden, source)

    def test_positive_examples_pass(self) -> None:
        for spec in EXAMPLES:
            with self.subTest(spec=spec.name):
                result = run_validator(spec)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn("Spec validation OK.", result.stdout)
                self.assertEqual(result.stderr, "")

    def test_missing_required_field_reports_path(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "missing-required.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("/slides/0/elements/0/text: missing required property", result.stderr)

    def test_wrong_element_type_reports_path(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "wrong-element-type.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("/slides/0/elements/0/type:", result.stderr)
        self.assertIn("'diagram'", result.stderr)

    def test_non_positive_size_reports_width_and_height(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "non-positive-size.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("/slides/0/elements/0/w: 0 must be greater than 0", result.stderr)
        self.assertIn("/slides/0/elements/0/h: -8 must be greater than 0", result.stderr)
        self.assertNotIn("outside the canvas", result.stderr)

    def test_rectangle_outside_canvas_reports_field_path(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "outside-canvas.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("/slides/0/elements/0/x:", result.stderr)
        self.assertIn("/slides/0/elements/0/y:", result.stderr)
        self.assertIn("outside the canvas", result.stderr)
        self.assertNotIn("must be greater than 0", result.stderr)

    def test_unknown_asset_id_reports_field_path(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "unknown-asset.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn(
            "/slides/0/elements/0/asset: unknown asset id 'missing-hero'",
            result.stderr,
        )

    def test_bad_color_reports_path_and_reason(self) -> None:
        result = run_validator(ROOT / "tests" / "fixtures" / "invalid" / "bad-color.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("/slides/0/background:", result.stderr)
        self.assertIn("#RRGGBB", result.stderr)

    def test_partial_overflow_and_line_endpoints(self) -> None:
        overflow = {
            "page": {"width": 960, "height": 540},
            "slides": [
                {
                    "elements": [
                        {
                            "type": "box",
                            "shape": "rect",
                            "x": 900,
                            "y": 20,
                            "w": 200,
                            "h": 40,
                        },
                        {
                            "type": "line",
                            "from": [10, 10],
                            "to": [2000, 10],
                            "line": {"color": "#64748B", "width": 1.5},
                        },
                    ]
                },
                {
                    "elements": [
                        {
                            "type": "connector",
                            "connector": "elbow",
                            "x1": 10,
                            "y1": 10,
                            "x2": 20,
                            "y2": 800,
                        }
                    ]
                },
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            spec = Path(directory) / "overflow.json"
            spec.write_text(json.dumps(overflow), encoding="utf-8")
            result = run_validator(spec)
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("/slides/0/elements/0/w:", result.stderr)
        self.assertIn("outside the canvas", result.stderr)
        self.assertIn("/slides/0/elements/1/to:", result.stderr)
        self.assertIn("/slides/1/elements/0/y2:", result.stderr)

    def test_duplicate_asset_id(self) -> None:
        spec_body = {
            "assets": [
                {"id": "hero", "prompt": "first"},
                {"id": "hero", "prompt": "second"},
            ],
            "slides": [
                {
                    "elements": [
                        {
                            "type": "text",
                            "text": "Hello",
                            "x": 10,
                            "y": 10,
                            "w": 80,
                            "h": 24,
                        }
                    ]
                }
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            spec = Path(directory) / "duplicate.json"
            spec.write_text(json.dumps(spec_body), encoding="utf-8")
            result = run_validator(spec)
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("/assets/1/id: duplicate asset id 'hero'", result.stderr)


if __name__ == "__main__":
    unittest.main()
