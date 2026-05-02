#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import re
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple


SPECIAL_H1 = {
    "摘要": "cn_abstract",
    "abstract": "en_abstract",
    "参考文献": "references",
    "references": "references",
    "致谢": "acknowledgements",
    "附录": "appendix",
    "外文原文及译文": "foreign_translation",
}


@dataclass
class Block:
    type: str
    level: int = 0
    text: str = ""
    caption: str = ""
    path: str = ""
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class PaperContent:
    title: str = ""
    english_title: str = ""
    cover: Dict[str, str] = field(default_factory=dict)
    cn_abstract: List[str] = field(default_factory=list)
    cn_keywords: List[str] = field(default_factory=list)
    en_abstract: List[str] = field(default_factory=list)
    en_keywords: List[str] = field(default_factory=list)
    body: List[Block] = field(default_factory=list)
    references: List[str] = field(default_factory=list)
    acknowledgements: List[str] = field(default_factory=list)
    appendix: List[str] = field(default_factory=list)
    foreign_translation: List[str] = field(default_factory=list)


def clean_heading(text: str) -> str:
    text = re.sub(r"^#{1,6}\s*", "", text).strip()
    text = re.sub(r"^[一二三四五六七八九十]+\s*[、.]\s*", "", text)
    text = re.sub(r"^第[一二三四五六七八九十百]+章\s*", "", text)
    text = re.sub(r"^（[一二三四五六七八九十]+）\s*", "", text)
    text = re.sub(r"^\(\d+\)\s*", "", text)
    text = re.sub(r"^[（(]\d+[）)]\s*", "", text)
    text = re.sub(r"^\d+(?:\.\d+)*\s*", "", text)
    return text.strip()


def _flush_paragraph(buffer: List[str], target: List[str]) -> None:
    text = "\n".join(part.rstrip() for part in buffer).strip()
    if text:
        target.append(text)
    buffer.clear()


def _flush_body_paragraph(buffer: List[str], blocks: List[Block]) -> None:
    text = "\n".join(part.rstrip() for part in buffer).strip()
    if text:
        blocks.append(Block(type="paragraph", text=text))
    buffer.clear()


def _strip_generated_number_prefix(text: str) -> str:
    cleaned = text.strip()
    cleaned = re.sub(r"^(?:表|图)\s*\d+(?:[-－]\d+)*\s*", "", cleaned)
    cleaned = re.sub(r"^(?:table|fig(?:ure)?)\s*\d+(?:[-－]\d+)*\s*", "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


def _parse_caption_directive(text: str) -> Optional[Tuple[str, str]]:
    match = re.match(r"^(表题|图题)\s*[：:]\s*(.+)$", text)
    if match:
        caption_type = "table" if match.group(1) == "表题" else "image"
        return caption_type, _strip_generated_number_prefix(match.group(2))
    return None


def _parse_image_markdown(text: str) -> Optional[Tuple[str, str]]:
    match = re.match(r"^!\[(.*?)\]\((.+?)\)\s*$", text)
    if not match:
        return None
    caption = _strip_generated_number_prefix(match.group(1).strip())
    path = match.group(2).strip()
    if path.startswith("<") and path.endswith(">"):
        path = path[1:-1].strip()
    title_match = re.match(r'(.+?)\s+"([^"]+)"$', path)
    if title_match:
        path = title_match.group(1).strip()
        if not caption:
            caption = _strip_generated_number_prefix(title_match.group(2).strip())
    return caption, path


def _looks_like_table_separator(text: str) -> bool:
    stripped = text.strip()
    if "|" not in stripped:
        return False
    parts = [part.strip() for part in stripped.strip("|").split("|")]
    if not parts:
        return False
    return all(re.match(r"^:?-{3,}:?$", part) for part in parts)


def _split_table_row(text: str) -> List[str]:
    stripped = text.strip().strip("|")
    return [part.strip() for part in stripped.split("|")]


def _parse_markdown_table(lines: List[str], start_index: int) -> Tuple[Optional[List[List[str]]], int]:
    if start_index + 1 >= len(lines):
        return None, start_index

    header_line = lines[start_index].strip()
    separator_line = lines[start_index + 1].strip()
    if "|" not in header_line or not _looks_like_table_separator(separator_line):
        return None, start_index

    rows = [_split_table_row(header_line)]
    index = start_index + 2
    while index < len(lines):
        stripped = lines[index].strip()
        if not stripped or "|" not in stripped:
            break
        rows.append(_split_table_row(stripped))
        index += 1

    width = len(rows[0])
    normalized_rows: List[List[str]] = []
    for row in rows:
        if len(row) < width:
            row = row + [""] * (width - len(row))
        elif len(row) > width:
            row = row[:width]
        normalized_rows.append(row)
    return normalized_rows, index


def _section_key_from_heading(title: str) -> Optional[str]:
    compact = re.sub(r"\s+", "", title).lower()
    return SPECIAL_H1.get(compact)


def _extract_keywords(line: str) -> List[str]:
    parts = re.split(r"[：:]", line, maxsplit=1)
    value = parts[1] if len(parts) > 1 else ""
    if not value:
        return []
    return [item.strip() for item in re.split(r"[，,;；]", value) if item.strip()]


def parse_markdown(markdown_path: Path) -> PaperContent:
    content = markdown_path.read_text(encoding="utf-8")
    lines = content.splitlines()

    paper = PaperContent()
    current_section = "body"
    paragraph_buffer: List[str] = []
    body_buffer: List[str] = []
    first_h1_as_title = True
    pending_table_caption: Optional[str] = None
    pending_image_caption: Optional[str] = None

    index = 0
    while index < len(lines):
        raw_line = lines[index]
        line = raw_line.rstrip()
        stripped = line.strip()

        if not stripped:
            if current_section == "body":
                _flush_body_paragraph(body_buffer, paper.body)
            elif current_section in {"cn_abstract", "en_abstract", "acknowledgements", "appendix", "foreign_translation"}:
                target = getattr(paper, current_section)
                _flush_paragraph(paragraph_buffer, target)
            index += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading_match:
            if current_section == "body":
                _flush_body_paragraph(body_buffer, paper.body)
            elif current_section in {"cn_abstract", "en_abstract", "acknowledgements", "appendix", "foreign_translation"}:
                target = getattr(paper, current_section)
                _flush_paragraph(paragraph_buffer, target)

            level = len(heading_match.group(1))
            title_text = clean_heading(heading_match.group(2))
            special_key = _section_key_from_heading(title_text)

            if level == 1 and first_h1_as_title and special_key is None:
                paper.title = title_text
                first_h1_as_title = False
                current_section = "body"
                index += 1
                continue

            if level == 1 and special_key is not None:
                current_section = special_key
                pending_table_caption = None
                pending_image_caption = None
                index += 1
                continue

            current_section = "body"
            if level <= 3:
                paper.body.append(Block(type="heading", level=level, text=title_text))
            else:
                body_buffer.append(title_text)
            pending_table_caption = None
            pending_image_caption = None
            index += 1
            continue

        if current_section == "cn_abstract":
            if re.match(r"^关键词\s*[：:]", stripped):
                _flush_paragraph(paragraph_buffer, paper.cn_abstract)
                paper.cn_keywords = _extract_keywords(stripped)
            else:
                paragraph_buffer.append(stripped)
            index += 1
            continue

        if current_section == "en_abstract":
            if re.match(r"^KEY\s*WORDS?\s*[：:]", stripped, flags=re.IGNORECASE):
                _flush_paragraph(paragraph_buffer, paper.en_abstract)
                paper.en_keywords = _extract_keywords(stripped)
            else:
                paragraph_buffer.append(stripped)
            index += 1
            continue

        if current_section == "references":
            if re.match(r"^\[\d+\]", stripped):
                paper.references.append(stripped)
            elif stripped:
                paper.references.append(stripped)
            index += 1
            continue

        if current_section in {"acknowledgements", "appendix", "foreign_translation"}:
            paragraph_buffer.append(stripped)
            index += 1
            continue

        caption_directive = _parse_caption_directive(stripped)
        if caption_directive:
            _flush_body_paragraph(body_buffer, paper.body)
            caption_type, caption_text = caption_directive
            if caption_type == "table":
                pending_table_caption = caption_text
            else:
                pending_image_caption = caption_text
            index += 1
            continue

        image_markdown = _parse_image_markdown(stripped)
        if image_markdown:
            _flush_body_paragraph(body_buffer, paper.body)
            caption_text, path = image_markdown
            paper.body.append(
                Block(
                    type="image",
                    caption=pending_image_caption or caption_text,
                    path=path,
                )
            )
            pending_image_caption = None
            index += 1
            continue

        table_rows, next_index = _parse_markdown_table(lines, index)
        if table_rows is not None:
            _flush_body_paragraph(body_buffer, paper.body)
            paper.body.append(
                Block(
                    type="table",
                    caption=pending_table_caption or "",
                    rows=table_rows,
                )
            )
            pending_table_caption = None
            index = next_index
            continue

        body_buffer.append(stripped)
        index += 1

    if current_section == "body":
        _flush_body_paragraph(body_buffer, paper.body)
    elif current_section in {"cn_abstract", "en_abstract", "acknowledgements", "appendix", "foreign_translation"}:
        target = getattr(paper, current_section)
        _flush_paragraph(paragraph_buffer, target)

    return paper


def load_metadata(meta_path: Optional[Path]) -> Dict[str, str]:
    if meta_path is None:
        return {}
    with meta_path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def merge_content_with_metadata(paper: PaperContent, metadata: Dict[str, str]) -> PaperContent:
    merged = deepcopy(paper)
    merged.cover.update(metadata.get("cover", {}))
    if metadata.get("title"):
        merged.title = metadata["title"]
    if metadata.get("english_title"):
        merged.english_title = metadata["english_title"]
    if metadata.get("cn_abstract"):
        merged.cn_abstract = metadata["cn_abstract"]
    if metadata.get("cn_keywords"):
        merged.cn_keywords = metadata["cn_keywords"]
    if metadata.get("en_abstract"):
        merged.en_abstract = metadata["en_abstract"]
    if metadata.get("en_keywords"):
        merged.en_keywords = metadata["en_keywords"]
    if metadata.get("references"):
        merged.references = metadata["references"]
    if metadata.get("acknowledgements"):
        merged.acknowledgements = metadata["acknowledgements"]
    if metadata.get("appendix"):
        merged.appendix = metadata["appendix"]
    if metadata.get("foreign_translation"):
        merged.foreign_translation = metadata["foreign_translation"]
    return merged
