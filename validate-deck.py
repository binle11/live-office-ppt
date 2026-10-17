#!/usr/bin/env python3
"""Validate a Live Office PPT deck spec without PowerPoint.

Agents and CI should run this program. It loads schema/deck.schema.json, then
applies canvas, asset, and file rules that depend on more than one field.
PowerPoint COM rendering stays in render-live-ppt.ps1 and is never started here.
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from pathlib import Path

try:
    from jsonschema import Draft202012Validator
    from jsonschema.exceptions import SchemaError, ValidationError
except ImportError:
    print(
        "jsonschema is required for deck validation.\n"
        "Install it with: python3 -m pip install -r requirements-validate.txt",
        file=sys.stderr,
    )
    sys.exit(2)


ROOT = Path(__file__).resolve().parent
DEFAULT_SCHEMA = ROOT / "schema" / "deck.schema.json"
DEFAULT_PAGE_WIDTH = 960
DEFAULT_PAGE_HEIGHT = 540
CANVAS_EPSILON = 1e-6
POSITIONED_TYPES = {"text", "title", "box", "shape", "image"}
LINE_TYPES = {"line", "connector"}
COLOR_PATTERN = "^#[0-9A-Fa-f]{6}$"


def pointer_token(part: object) -> str:
    return str(part).replace("~", "~0").replace("/", "~1")


def json_pointer(parts: object) -> str:
    tokens = [pointer_token(part) for part in parts]
    if not tokens:
        return "/"
    return "/" + "/".join(tokens)


def join_pointer(base: str, token: str) -> str:
    escaped = pointer_token(token)
    if base == "/":
        return "/" + escaped
    return f"{base}/{escaped}"


def display_number(value: object) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def required_property_name(error: ValidationError) -> str | None:
    marker = " is a required property"
    if not error.message.endswith(marker):
        return None
    try:
        value = ast.literal_eval(error.message[: -len(marker)])
    except (SyntaxError, ValueError):
        return None
    if isinstance(value, str):
        return value
    return None


def format_error(error: ValidationError) -> str:
    pointer = json_pointer(error.absolute_path)
    if error.validator == "required":
        name = required_property_name(error)
        if name is not None:
            return f"{join_pointer(pointer, name)}: missing required property"
    if error.validator == "exclusiveMinimum":
        return (
            f"{pointer}: {display_number(error.instance)} must be greater than "
            f"{display_number(error.validator_value)}"
        )
    if error.validator == "minimum":
        return (
            f"{pointer}: {display_number(error.instance)} must be greater than or equal to "
            f"{display_number(error.validator_value)}"
        )
    if error.validator == "maximum":
        return (
            f"{pointer}: {display_number(error.instance)} must be less than or equal to "
            f"{display_number(error.validator_value)}"
        )
    if error.validator == "exclusiveMaximum":
        return (
            f"{pointer}: {display_number(error.instance)} must be less than "
            f"{display_number(error.validator_value)}"
        )
    if error.validator == "enum":
        choices = ", ".join(repr(choice) for choice in error.validator_value)
        return f"{pointer}: {error.instance!r} is not one of: {choices}"
    if error.validator == "const":
        return f"{pointer}: expected {error.validator_value!r}"
    if error.validator == "pattern" and error.validator_value == COLOR_PATTERN:
        return f"{pointer}: {error.instance!r} must be a #RRGGBB color"
    if error.validator == "minLength":
        amount = error.validator_value
        unit = "character" if amount == 1 else "characters"
        return f"{pointer}: must be at least {amount} {unit}"
    if error.validator == "minItems":
        amount = error.validator_value
        unit = "item" if amount == 1 else "items"
        return f"{pointer}: must contain at least {amount} {unit}"
    if error.validator == "maxItems":
        amount = error.validator_value
        unit = "item" if amount == 1 else "items"
        return f"{pointer}: must contain at most {amount} {unit}"
    if error.validator == "type":
        expected = error.validator_value
        if isinstance(expected, list):
            expected = " or ".join(str(item) for item in expected)
        return f"{pointer}: expected {expected}"
    if error.validator in {"oneOf", "anyOf"}:
        description = ""
        if isinstance(error.schema, dict):
            description = str(error.schema.get("description") or "")
        if description:
            if isinstance(error.instance, (str, int, float, bool)) or error.instance is None:
                return f"{pointer}: {description} (got {error.instance!r})"
            return f"{pointer}: {description}"
        return f"{pointer}: does not match the expected schema"
    return f"{pointer}: {error.message}"


def leaf_errors(error: ValidationError) -> list[ValidationError]:
    """Prefer field-level failures inside oneOf/anyOf over the wrapper message."""
    if error.validator in {"oneOf", "anyOf"} and error.context:
        leaves: list[ValidationError] = []
        for child in error.context:
            leaves.extend(leaf_errors(child))
        parent_length = len(tuple(error.absolute_path))
        deeper = [leaf for leaf in leaves if len(tuple(leaf.absolute_path)) > parent_length]
        if deeper:
            return deeper
        if isinstance(error.schema, dict) and error.schema.get("description"):
            return [error]
        if leaves and all(leaf.validator == "required" for leaf in leaves):
            return [error]
        return leaves or [error]
    return [error]


def schema_errors(validator: Draft202012Validator, instance: object) -> list[str]:
    lines: list[str] = []
    for error in validator.iter_errors(instance):
        for leaf in leaf_errors(error):
            lines.append(format_error(leaf))
    unique: list[str] = []
    seen: set[str] = set()
    for line in lines:
        if line not in seen:
            seen.add(line)
            unique.append(line)
    unique.sort()
    return unique


def past_limit(value: float, limit: float) -> bool:
    return value > limit + CANVAS_EPSILON


def page_size(spec: dict) -> tuple[float, float]:
    page = spec.get("page") or {}
    width = page.get("width", DEFAULT_PAGE_WIDTH)
    height = page.get("height", DEFAULT_PAGE_HEIGHT)
    return float(width), float(height)


def iter_assets(spec: dict):
    assets = spec.get("assets", [])
    if isinstance(assets, list):
        for index, asset in enumerate(assets):
            asset_id = asset.get("id") if isinstance(asset, dict) else None
            yield asset_id, asset if isinstance(asset, dict) else {}, f"/assets/{index}"
        return
    if isinstance(assets, dict):
        for key, asset in assets.items():
            body = asset if isinstance(asset, dict) else {}
            explicit_id = body.get("id")
            asset_id = explicit_id if isinstance(explicit_id, str) and explicit_id else key
            yield asset_id, body, join_pointer("/assets", str(key))


def resolve_local_file(spec_dir: Path, raw: str) -> Path:
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = spec_dir / candidate
    return candidate


def asset_file_field(asset: dict) -> tuple[str, str] | None:
    path = asset.get("path")
    if isinstance(path, str) and path:
        return "path", path
    file_name = asset.get("file")
    if isinstance(file_name, str) and file_name:
        return "file", file_name
    return None


def check_rectangle(base: str, element: dict, width: float, height: float) -> list[str]:
    errors: list[str] = []
    x = float(element["x"])
    y = float(element["y"])
    box_w = float(element["w"])
    box_h = float(element["h"])
    if past_limit(x, width):
        errors.append(
            f"{base}/x: rectangle is outside the canvas "
            f"(x={display_number(x)} is past page width {display_number(width)})"
        )
    elif past_limit(x + box_w, width):
        errors.append(
            f"{base}/w: rectangle extends outside the canvas "
            f"(x + w = {display_number(x + box_w)} > page width {display_number(width)})"
        )
    if past_limit(y, height):
        errors.append(
            f"{base}/y: rectangle is outside the canvas "
            f"(y={display_number(y)} is past page height {display_number(height)})"
        )
    elif past_limit(y + box_h, height):
        errors.append(
            f"{base}/h: rectangle extends outside the canvas "
            f"(y + h = {display_number(y + box_h)} > page height {display_number(height)})"
        )
    return errors


def check_axis(path: str, value: object, limit: float, axis: str) -> str | None:
    number = float(value)
    if number >= -CANVAS_EPSILON and not past_limit(number, limit):
        return None
    edge = "page width" if axis == "x" else "page height"
    return (
        f"{path}: {axis}={display_number(number)} is outside the canvas "
        f"({edge} {display_number(limit)})"
    )


def check_point(base: str, point: list, width: float, height: float) -> str | None:
    px = float(point[0])
    py = float(point[1])
    problems: list[str] = []
    if px < -CANVAS_EPSILON or past_limit(px, width):
        problems.append(f"x={display_number(px)} is outside page width {display_number(width)}")
    if py < -CANVAS_EPSILON or past_limit(py, height):
        problems.append(f"y={display_number(py)} is outside page height {display_number(height)}")
    if not problems:
        return None
    return (
        f"{base}: point [{display_number(px)}, {display_number(py)}] is outside the canvas "
        f"({'; '.join(problems)})"
    )


def cross_field_errors(spec: dict, spec_dir: Path) -> list[str]:
    errors: list[str] = []
    width, height = page_size(spec)
    assets: dict[str, tuple[str, dict]] = {}

    for asset_id, asset, pointer in iter_assets(spec):
        if not isinstance(asset_id, str) or asset_id == "":
            errors.append(f"{pointer}/id: asset id is required")
            continue
        if asset_id in assets:
            previous = assets[asset_id][0]
            errors.append(
                f"{pointer}/id: duplicate asset id {asset_id!r} (also defined at {previous})"
            )
            continue
        assets[asset_id] = (pointer, asset)
        located = asset_file_field(asset)
        if located is None:
            continue
        field, raw = located
        if not resolve_local_file(spec_dir, raw).is_file():
            errors.append(f"{pointer}/{field}: file not found: {raw}")

    for slide_index, slide in enumerate(spec["slides"]):
        for element_index, element in enumerate(slide["elements"]):
            base = f"/slides/{slide_index}/elements/{element_index}"
            element_type = element["type"]
            if element_type in POSITIONED_TYPES:
                errors.extend(check_rectangle(base, element, width, height))
            elif element_type in LINE_TYPES:
                if "from" in element and "to" in element:
                    for name in ("from", "to"):
                        message = check_point(f"{base}/{name}", element[name], width, height)
                        if message:
                            errors.append(message)
                if all(field in element for field in ("x1", "y1", "x2", "y2")):
                    for field, axis, limit in (
                        ("x1", "x", width),
                        ("y1", "y", height),
                        ("x2", "x", width),
                        ("y2", "y", height),
                    ):
                        message = check_axis(f"{base}/{field}", element[field], limit, axis)
                        if message:
                            errors.append(message)
            if element_type != "image":
                continue
            if "asset" in element:
                asset_id = element["asset"]
                if asset_id not in assets:
                    errors.append(f"{base}/asset: unknown asset id {asset_id!r}")
                    continue
                asset = assets[asset_id][1]
                if asset_file_field(asset) is None:
                    errors.append(f"{base}/asset: asset {asset_id!r} has no path")
                continue
            raw_path = element.get("path")
            if isinstance(raw_path, str) and not resolve_local_file(spec_dir, raw_path).is_file():
                errors.append(f"{base}/path: file not found: {raw_path}")

    errors.sort()
    return errors


def load_json(path: Path) -> tuple[object | None, list[str]]:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        return None, [f"/: spec file not found: {path} ({exc.strerror})"]
    try:
        return json.loads(text), []
    except json.JSONDecodeError as exc:
        return None, [
            f"/: invalid JSON at line {exc.lineno}, column {exc.colno}: {exc.msg}"
        ]


def asset_count(spec: dict) -> int:
    assets = spec.get("assets", [])
    if isinstance(assets, (list, dict)):
        return len(assets)
    return 0


def validate_file(spec_path: Path, validator: Draft202012Validator) -> tuple[list[str], dict | None]:
    if not spec_path.is_file():
        return [f"/: spec file not found: {spec_path}"], None
    instance, errors = load_json(spec_path)
    if errors:
        return errors, None
    errors = schema_errors(validator, instance)
    if errors:
        return errors, None
    if not isinstance(instance, dict):
        return ["/: expected an object"], None
    return cross_field_errors(instance, spec_path.parent), instance


def load_validator(schema_path: Path) -> Draft202012Validator:
    if not schema_path.is_file():
        print(f"Schema file not found: {schema_path}", file=sys.stderr)
        sys.exit(2)
    try:
        schema = json.loads(schema_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(
            f"Schema is not valid JSON at line {exc.lineno}, column {exc.colno}: {exc.msg}",
            file=sys.stderr,
        )
        sys.exit(2)
    try:
        Draft202012Validator.check_schema(schema)
    except SchemaError as exc:
        print(f"Schema is invalid: {exc.message}", file=sys.stderr)
        sys.exit(2)
    return Draft202012Validator(schema)


def report_success(spec_path: Path, spec: dict) -> None:
    print("Spec validation OK.")
    print(f"Spec: {spec_path}")
    print(f"Slides: {len(spec['slides'])}")
    print(f"Assets: {asset_count(spec)}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Validate a Live Office PPT deck spec. Loads the JSON schema and "
            "cross-field canvas rules. Does not start PowerPoint."
        )
    )
    parser.add_argument(
        "specs",
        nargs="+",
        type=Path,
        metavar="SPEC",
        help="Path to a deck JSON spec",
    )
    parser.add_argument(
        "--schema",
        type=Path,
        default=DEFAULT_SCHEMA,
        help="JSON Schema file (default: schema/deck.schema.json)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    validator = load_validator(args.schema)
    failed = False
    for index, spec_path in enumerate(args.specs):
        if index:
            print()
        errors, spec = validate_file(spec_path, validator)
        if errors or spec is None:
            failed = True
            if len(args.specs) > 1:
                print(spec_path, file=sys.stderr)
            for error in errors or ["/: expected an object"]:
                print(error, file=sys.stderr)
            continue
        report_success(spec_path, spec)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
