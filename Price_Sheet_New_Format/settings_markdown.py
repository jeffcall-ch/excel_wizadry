from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List


SECTION_LABELS: Dict[str, str] = {
    "General": "Section 1. General",
    "MaterialTypes": "Section 2. MaterialTypes",
    "RequiredFileCategories": "Section 3. RequiredFileCategories",
    "MaterialTypeLabels": "Section 4. MaterialTypeLabels",
    "ColumnMappings_Normal": "Section 5. ColumnMappings_Normal",
    "ColumnMappings_Erection": "Section 6. ColumnMappings_Erection",
    "ColumnMappings_Paint": "Section 7. ColumnMappings_Paint",
    "Aliases": "Section 8. Aliases",
    "FileKeywords": "Section 9. FileKeywords",
    "SqliteNumericColumns": "Section 10. SqliteNumericColumns",
    "AggregateSumColumns": "Section 11. AggregateSumColumns",
    "AggregationMaterialTypeRules": "Section 12. AggregationMaterialTypeRules",
    "SpareGeneral": "Section 13. SpareGeneral",
    "SpareMainRules": "Section 14. SpareMainRules",
    "SpareErectionRules": "Section 15. SpareErectionRules",
    "SpareKeywords": "Section 16. SpareKeywords",
    "SpareRuleGuide": "Section 17. SpareRuleGuide",
    "BlindDiskThickness": "Section 18. BlindDiskThickness",
    "PaintErectionDefaults": "Section 19. PaintErectionDefaults",
    "BlindDiskDefaults": "Section 20. BlindDiskDefaults",
    "FlangeGuardSystems": "Section 21. FlangeGuardSystems",
    "FlangeGuardDefaults": "Section 22. FlangeGuardDefaults",
    "ColumnMappings_Orifice": "Section 23. ColumnMappings_Orifice",
    "OrificeDefaults": "Section 24. OrificeDefaults",
    "MapressSealOrderSize": "Section 25. MapressSealOrderSize",
}


@dataclass
class ParsedSection:
    number: str
    title: str
    records: List[Dict[str, str]]

    @property
    def label(self) -> str:
        return f"Section {self.number}. {self.title}"


@dataclass
class ParsedSettingsDocument:
    sections: Dict[str, ParsedSection]

    def get_records(self, section_title: str) -> List[Dict[str, str]]:
        key = _normalize_section_title(section_title)
        if key not in self.sections:
            raise ValueError(
                f"Missing settings section: {section_label(section_title)}."
            )
        return self.sections[key].records


def section_label(section_title: str) -> str:
    return SECTION_LABELS.get(section_title, f"section '{section_title}'")


def sections_hint(*section_titles: str) -> str:
    joined = ", ".join(section_label(name) for name in section_titles)
    return f"Check settings file sections: {joined}."


def load_markdown_settings(settings_file: Path) -> ParsedSettingsDocument:
    if not settings_file.exists():
        raise ValueError(f"Settings file not found: {settings_file}")

    text = settings_file.read_text(encoding="utf-8")
    lines = text.splitlines()

    heading_pattern = re.compile(r"^#{2,3}\s+(\d+(?:\.\d+)*)\.\s+(.+?)\s*$")
    sections: Dict[str, ParsedSection] = {}

    current_number = ""
    current_title = ""
    current_lines: List[str] = []

    def flush_section() -> None:
        nonlocal current_number, current_title, current_lines
        if not current_title:
            return
        records = _parse_first_table(current_lines)
        key = _normalize_section_title(current_title)
        sections[key] = ParsedSection(number=current_number, title=current_title, records=records)

    for line in lines:
        match = heading_pattern.match(line.strip())
        if match:
            flush_section()
            current_number = match.group(1)
            current_title = match.group(2).strip()
            current_lines = []
        else:
            if current_title:
                current_lines.append(line)

    flush_section()

    return ParsedSettingsDocument(sections=sections)


def _normalize_section_title(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.strip().lower())


def _is_table_separator(line: str) -> bool:
    stripped = line.strip()
    if "|" not in stripped:
        return False
    candidate = stripped.replace("|", "").replace("-", "").replace(":", "").replace(" ", "")
    return candidate == ""


def _parse_table_row(line: str) -> List[str]:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split("|")]


def _parse_first_table(lines: List[str]) -> List[Dict[str, str]]:
    start_index = -1
    for idx in range(0, len(lines) - 1):
        first = lines[idx].strip()
        second = lines[idx + 1].strip()
        if "|" in first and _is_table_separator(second):
            start_index = idx
            break

    if start_index < 0:
        return []

    headers = _parse_table_row(lines[start_index])
    records: List[Dict[str, str]] = []

    for idx in range(start_index + 2, len(lines)):
        row_line = lines[idx].rstrip()
        if not row_line.strip():
            break
        if "|" not in row_line:
            break
        values = _parse_table_row(row_line)
        if len(values) < len(headers):
            values.extend([""] * (len(headers) - len(values)))
        if len(values) > len(headers):
            values = values[: len(headers)]
        record = {headers[col_idx]: values[col_idx] for col_idx in range(len(headers)) if headers[col_idx]}
        if any(str(value).strip() for value in record.values()):
            records.append(record)

    return records