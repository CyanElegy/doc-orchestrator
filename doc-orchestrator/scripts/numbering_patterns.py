#!/usr/bin/env python3
"""Shared Chinese numbering pattern detection for doc-orchestrator scripts.

Used by both extract_template.py and assemble_docx.py to ensure consistent
numbering prefix detection across the pipeline.
"""

import re

# Recognised numbering patterns in order of priority
NUMBERING_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"^[一二三四五六七八九十百千]+[、．]?\s*"), "一、"),  # 中文数字
    (re.compile(r"^（[一二三四五六七八九十百千]+）\s*"), "（一）"),  # （一）
    (re.compile(r"^\d+\.[\d.]*\s*"), "1.1"),  # 1.1 / 1.1.1
    (re.compile(r"^第[一二三四五六七八九十百千]+[章节条]\s*"), "第一章"),  # 第X章
    (re.compile(r"^\(\d+\)\s*"), "(1)"),  # (1)
    (re.compile(r"^\d+[、．.]\s*"), "1、"),  # 1、 / 1．
    (re.compile(r"^①\s*"), "①"),  # circled digit (standalone)
    (re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*"), "①"),  # more circled digits
]


def detect_numbering_prefix(text: str) -> tuple[str, str]:
    """Detect numbering prefix in heading text.

    Returns (prefix, pattern_name) where pattern_name is a human-readable
    label like '一、' or '1.' indicating the numbering style.
    """
    for pattern, label in NUMBERING_PATTERNS:
        m = pattern.match(text)
        if m:
            return m.group().strip(), label
    return "", ""


def strip_numbering(text: str) -> str:
    """Remove numbering prefix from heading text."""
    for pattern, _label in NUMBERING_PATTERNS:
        m = pattern.match(text)
        if m:
            return text[m.end():].strip()
    return text.strip()
