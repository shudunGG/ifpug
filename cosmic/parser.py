"""Configuration parser for COSMIC measurement definitions."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

try:  # pragma: no cover - optional dependency
    import yaml  # type: ignore
except ModuleNotFoundError:  # pragma: no cover - fallback to lightweight parser
    yaml = None  # type: ignore

from .models import SystemMeasurement


class MeasurementParserError(RuntimeError):
    """Raised when a measurement configuration cannot be parsed."""


def _strip_comments(lines: Iterable[str]) -> List[str]:
    cleaned: List[str] = []
    for line in lines:
        stripped = line.rstrip()
        if not stripped:
            continue
        if stripped.lstrip().startswith("#"):
            continue
        cleaned.append(stripped)
    return cleaned


def _parse_scalar(value: str) -> Any:
    if value.startswith(("'", '"')) and value.endswith(("'", '"')) and len(value) >= 2:
        return value[1:-1]
    lowered = value.lower()
    if lowered in {"true", "false"}:
        return lowered == "true"
    if lowered == "null":
        return None
    try:
        if "." in value:
            return float(value)
        return int(value)
    except ValueError:
        return value


def _parse_yaml_lines(lines: List[Tuple[int, str]], start: int, indent: int) -> Tuple[Any, int]:
    result: Any = None
    index = start

    while index < len(lines):
        current_indent, content = lines[index]
        if current_indent < indent:
            break
        if content.startswith("- "):
            if result is None:
                result = []
            elif not isinstance(result, list):
                raise MeasurementParserError("Mixed list and mapping structures are not supported in simple YAML parser.")
            item_text = content[2:].strip()
            index += 1

            if not item_text:
                if index >= len(lines):
                    result.append(None)
                else:
                    next_indent, _ = lines[index]
                    value, index = _parse_yaml_lines(lines, index, next_indent)
                    result.append(value)
                continue

            if ":" in item_text:
                key, _, value_part = item_text.partition(":")
                key = key.strip()
                value_part = value_part.lstrip()
                if value_part:
                    item_value: Any = {key: _parse_scalar(value_part)}
                elif index < len(lines) and lines[index][0] > current_indent:
                    nested_value, index = _parse_yaml_lines(lines, index, lines[index][0])
                    item_value = {key: nested_value}
                else:
                    item_value = {key: None}

                while index < len(lines) and lines[index][0] > current_indent:
                    extra_indent = lines[index][0]
                    extra_value, index = _parse_yaml_lines(lines, index, extra_indent)
                    if not isinstance(extra_value, dict):
                        raise MeasurementParserError(
                            "List item mappings must contain dictionary structures at consistent indentation."
                        )
                    item_value.update(extra_value)
                result.append(item_value)
                continue

            item_value = _parse_scalar(item_text)
            if index < len(lines) and lines[index][0] > current_indent:
                nested_value, index = _parse_yaml_lines(lines, index, lines[index][0])
                if isinstance(nested_value, dict):
                    item_value = {str(item_value): nested_value}
                else:
                    if not isinstance(nested_value, list):
                        nested_value = [nested_value]
                    item_value = [item_value, *nested_value]
            result.append(item_value)
            continue
        else:
            key, _, value_part = content.partition(":")
            key = key.strip()
            if result is None:
                result = {}
            elif not isinstance(result, dict):
                raise MeasurementParserError("Mixed list and mapping structures are not supported in simple YAML parser.")
            value_part = value_part.lstrip()
            if value_part:
                result[key] = _parse_scalar(value_part)
                index += 1
            else:
                index += 1
                if index < len(lines) and lines[index][0] > current_indent:
                    value, index = _parse_yaml_lines(lines, index, lines[index][0])
                else:
                    value = None
                result[key] = value
    if result is None:
        result = {}
    return result, index


def _simple_yaml_load(raw: str) -> Dict[str, Any]:
    cleaned = _strip_comments(raw.splitlines())
    processed: List[Tuple[int, str]] = []
    for line in cleaned:
        indent = len(line) - len(line.lstrip(" "))
        processed.append((indent, line.lstrip(" ")))
    parsed, _ = _parse_yaml_lines(processed, 0, processed[0][0] if processed else 0)
    if not isinstance(parsed, dict):
        raise MeasurementParserError("Top-level YAML structure must be a mapping.")
    return parsed


def _read_file(path: Path) -> Dict[str, Any]:
    """Read a JSON or YAML configuration file."""
    if not path.exists():
        raise MeasurementParserError(f"Configuration file '{path}' does not exist.")

    try:
        raw = path.read_text(encoding="utf-8")
    except OSError as exc:  # pragma: no cover - defensive branch
        raise MeasurementParserError(f"Failed to read configuration file '{path}': {exc}") from exc

    try:
        if path.suffix.lower() in {".json"}:
            return json.loads(raw)
        if yaml is not None:  # pragma: no cover - executed when PyYAML is available
            return yaml.safe_load(raw)
        return _simple_yaml_load(raw)
    except json.JSONDecodeError as exc:
        raise MeasurementParserError(f"Failed to parse JSON configuration file '{path}': {exc}") from exc
    except Exception as exc:
        raise MeasurementParserError(f"Failed to parse YAML configuration file '{path}': {exc}") from exc


def load_measurement(path: Path | str) -> SystemMeasurement:
    """Load a :class:`SystemMeasurement` definition from a file."""
    file_path = Path(path)
    payload = _read_file(file_path)
    if not isinstance(payload, dict):
        raise MeasurementParserError(
            "Measurement configuration must define a mapping with 'system' and 'functional_processes' keys."
        )
    return SystemMeasurement.from_dict(payload)
