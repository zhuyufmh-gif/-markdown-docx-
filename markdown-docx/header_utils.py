#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


HEADER_VISUAL_WIDTH = 72
HEADER_MIN_SPACES = 4


def _clear_tab_stops(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    tabs_nodes = p_pr.xpath("./w:tabs")
    for tabs in tabs_nodes:
        p_pr.remove(tabs)


def _visual_length(text: str) -> int:
    length = 0
    for ch in text:
        if ch.isspace():
            length += 1
        elif ord(ch) < 128:
            length += 1
        else:
            length += 2
    return length


def _apply_header_layout(paragraph, section, title: str) -> bool:
    drawing_seen = False
    title_run = None
    spacer_runs = []

    for run in paragraph.runs:
        has_drawing = run._element.find(qn("w:drawing")) is not None
        if has_drawing:
            drawing_seen = True
            continue
        if run.text and run.text.strip():
            title_run = run
        elif drawing_seen:
            spacer_runs.append(run)

    if not drawing_seen or title_run is None:
        return False

    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _clear_tab_stops(paragraph)
    target_spaces = max(HEADER_MIN_SPACES, HEADER_VISUAL_WIDTH - _visual_length(title))

    if spacer_runs:
        spacer_runs[0].text = " " * target_spaces
        for run in spacer_runs[1:]:
            run.text = ""
    else:
        title_run.text = (" " * target_spaces) + title
        return True

    title_run.text = title
    return True


def _update_header_title_for_container(header_container, section, title: str) -> None:
    for paragraph in header_container.paragraphs:
        if not (paragraph.text or "").strip():
            continue
        if _apply_header_layout(paragraph, section, title):
            continue
        title_run = None
        for run in paragraph.runs:
            if run.text and run.text.strip():
                title_run = run
        if title_run is None:
            paragraph.add_run(title)
        else:
            title_run.text = title


def update_headers(doc: Document, title: str) -> None:
    if not title:
        return

    seen = set()
    for section in doc.sections:
        containers = [section.header]
        if hasattr(section, "even_page_header"):
            containers.append(section.even_page_header)
        if hasattr(section, "first_page_header"):
            containers.append(section.first_page_header)

        for container in containers:
            key = id(container._element)
            if key in seen:
                continue
            seen.add(key)
            _update_header_title_for_container(container, section, title)
