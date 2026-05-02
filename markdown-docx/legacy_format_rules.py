#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from docx_utils import RunStyle


MAJOR_HEADING_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="黑体",
    size_pt=16.0,
    bold=False,
)

SUBHEADING_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="黑体",
    size_pt=14.0,
    bold=False,
)

BODY_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="宋体",
    size_pt=12.0,
    bold=False,
)

ENGLISH_BODY_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="Times New Roman",
    size_pt=12.0,
    bold=False,
)

KEYWORD_LABEL_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="黑体",
    size_pt=12.0,
    bold=True,
)

ENGLISH_KEYWORD_LABEL_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="Times New Roman",
    size_pt=12.0,
    bold=True,
)

TABLE_CAPTION_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="宋体",
    size_pt=12.0,
    bold=True,
)

CONTINUED_TABLE_LABEL_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="宋体",
    size_pt=12.0,
    bold=False,
)

FIGURE_CAPTION_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="宋体",
    size_pt=12.0,
    bold=True,
)

TABLE_CELL_STYLE = RunStyle(
    ascii_font="Times New Roman",
    east_asia_font="宋体",
    size_pt=12.0,
    bold=False,
)


def format_major_heading(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(12)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.5


def format_major_heading_on_new_page(paragraph) -> None:
    format_major_heading(paragraph)
    paragraph.paragraph_format.page_break_before = True


def format_subheading(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.5


def format_body(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(0.74)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.5


def format_keywords(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.5


def format_reference(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(-0.74)
    paragraph.paragraph_format.left_indent = Cm(0.74)
    paragraph.paragraph_format.line_spacing = 1.0


def format_table_caption(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.25


def format_continued_table_label(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.0


def format_figure_caption(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.0


def format_picture_block(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.line_spacing = 1.0
