#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional

from docx import Document


@dataclass
class AnchorMatch:
    name: str
    index: int


class TemplateLocatorError(RuntimeError):
    pass


def load_schema(schema_path: Path) -> Dict:
    with schema_path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", (text or "").strip())


def find_paragraph_index(doc: Document, pattern: str) -> Optional[int]:
    regex = re.compile(pattern)
    for idx, paragraph in enumerate(doc.paragraphs):
        text = (paragraph.text or "").strip()
        if text and regex.match(text):
            return idx
    return None


def locate_anchors(doc: Document, schema: Dict, required: bool = True) -> Dict[str, int]:
    anchor_map: Dict[str, int] = {}
    for name, pattern in schema["anchors"].items():
        match = find_paragraph_index(doc, pattern)
        if match is None:
            if required:
                raise TemplateLocatorError(f"未找到锚点: {name} / {pattern}")
            continue
        anchor_map[name] = match
    return anchor_map


def find_next_anchor(anchor_map: Dict[str, int], start_name: str, *candidate_names: str) -> Optional[str]:
    if start_name not in anchor_map:
        return None
    start_index = anchor_map[start_name]
    best_name = None
    best_index = None
    for name in candidate_names:
        index = anchor_map.get(name)
        if index is None or index <= start_index:
            continue
        if best_index is None or index < best_index:
            best_name = name
            best_index = index
    return best_name
