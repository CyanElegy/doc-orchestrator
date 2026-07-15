#!/usr/bin/env python3
"""
assemble_docx.py — Assemble a .docx from skeleton + content + template.

Usage:
    python3 assemble_docx.py skeleton.json content.json <template.docx> --output <output.docx>

Steps:
    1. Copy template → output (inherit all styles, headers, footers, page setup)
    2. Load skeleton.json and content.json
    3. Build a content lookup by unit id
    4. Clear body content from output
    5. Walk skeleton units in order, writing each unit:
       - heading: apply template heading style
       - paragraph (fixed): write template text as-is
       - paragraph (generated): write content from content.json
       - placeholder: replace with value from content.json
       - table: create table with data from content.json
       - image: embed referenced image file
    6. Insert TOC field code
    7. Save

If content.json only contains a subset of units, only those are replaced.
All other units remain as defined in the skeleton (fixed text) or are skipped (generated with no content).

Requires: python-docx, Pillow
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any, Optional

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Emu, RGBColor
except ImportError:
    print("python-docx is not installed. Run: pip install python-docx", file=sys.stderr)
    sys.exit(1)

try:
    from PIL import Image as PILImage
except ImportError:
    print("Pillow is not installed. Run: pip install Pillow", file=sys.stderr)
    sys.exit(1)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

HEADING_STYLE_MAP = {
    1: "Heading 1", 2: "Heading 2", 3: "Heading 3",
    4: "Heading 4", 5: "Heading 5", 6: "Heading 6",
    7: "Heading 7", 8: "Heading 8", 9: "Heading 9",
}


# ---------------------------------------------------------------------------
# Assembly
# ---------------------------------------------------------------------------

def assemble(
    skeleton_path: Path,
    content_path: Path,
    template_path: Path,
    output_path: Path,
) -> None:
    """Main assembly workflow."""

    # -- Validate inputs -------------------------------------------------------
    for p, label in [
        (skeleton_path, "Skeleton JSON"),
        (content_path, "Content JSON"),
        (template_path, "Template"),
    ]:
        if not p.exists():
            print(f"Error: {label} not found: {p}", file=sys.stderr)
            sys.exit(1)

    if template_path.suffix.lower() != ".docx":
        print(f"Error: Template must be .docx: {template_path}", file=sys.stderr)
        sys.exit(1)

    # Validate template is a valid docx
    try:
        with zipfile.ZipFile(template_path) as zf:
            if "[Content_Types].xml" not in zf.namelist():
                print(f"Error: Not a valid .docx: {template_path}", file=sys.stderr)
                sys.exit(1)
    except zipfile.BadZipFile:
        print(f"Error: Not a valid .docx (corrupt): {template_path}", file=sys.stderr)
        sys.exit(1)

    # -- Load data files -------------------------------------------------------
    try:
        skeleton = json.loads(skeleton_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"Error: Invalid skeleton JSON: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        content_data = json.loads(content_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"Error: Invalid content JSON: {exc}", file=sys.stderr)
        sys.exit(1)

    # Build content lookup
    content_lookup: dict[str, str] = {}
    for unit in content_data.get("units", []):
        unit_id = unit.get("id", "")
        unit_content = unit.get("content", "")
        if unit_id:
            content_lookup[unit_id] = str(unit_content)

    # -- Copy template to output -----------------------------------------------
    shutil.copy2(str(template_path), str(output_path))
    doc = Document(str(output_path))

    # -- Clear body ------------------------------------------------------------
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))
    children = list(body)
    for child in children:
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)
    if sect_pr is not None:
        body.append(sect_pr)

    # -- Walk skeleton and write content ---------------------------------------
    heading_count = 0
    table_count = 0
    units = skeleton.get("units", [])

    for unit in units:
        unit_id = unit.get("id", "")
        element = unit.get("element", "paragraph")
        unit_type = unit.get("type", "generated")

        try:
            if element == "heading":
                heading_count += 1
                level = unit.get("level", 1)
                text = unit.get("text", "")
                numbering = unit.get("numbering", "")

                style_name = _get_heading_style(doc, level)
                full_text = f"{numbering}{text}" if numbering else text

                para = doc.add_paragraph()
                para.style = doc.styles[style_name]
                para.add_run(full_text)

            elif element == "placeholder":
                key = unit.get("key", "")
                replacement = content_lookup.get(unit_id)
                if replacement is None:
                    # Try matching by placeholder key
                    replacement = content_lookup.get(f"ph-{key}", "")

                para = doc.add_paragraph()
                para.add_run(replacement if replacement else f"《{key}》")

            elif element == "paragraph":
                if unit_type == "fixed":
                    text = unit.get("text", "")
                    para = doc.add_paragraph()
                    para.style = doc.styles["Normal"]
                    para.add_run(text)
                else:
                    # generated paragraph
                    content = content_lookup.get(unit_id, "")
                    if content:
                        para = doc.add_paragraph()
                        para.style = doc.styles["Normal"]
                        _write_mixed_content(para, content, doc)
                    else:
                        # Empty generated unit — leave placeholder in output
                        desc = unit.get("description", "待生成")
                        para = doc.add_paragraph()
                        para.style = doc.styles["Normal"]
                        run = para.add_run(f"[{desc}]")
                        try:
                            run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                        except Exception:
                            pass

            elif element == "table":
                table_count += 1
                table_data_str = content_lookup.get(unit_id, "")
                headers = unit.get("headers", [])
                # Parse table data: JSON array of arrays
                rows_data: list[list[str]] = []
                if table_data_str:
                    try:
                        parsed = json.loads(table_data_str)
                        if isinstance(parsed, list):
                            rows_data = [[str(c) for c in row] for row in parsed]
                    except json.JSONDecodeError:
                        # Try as single string → one row
                        rows_data = [[table_data_str]]

                # Build full table: header row + data rows
                all_rows = [headers] + rows_data if headers else rows_data
                if not all_rows:
                    all_rows = [[""]]

                num_cols = max(len(r) for r in all_rows)
                table = doc.add_table(rows=len(all_rows), cols=num_cols)
                try:
                    table.style = doc.styles["Table Grid"]
                except KeyError:
                    pass

                for row_idx, row_cells in enumerate(all_rows):
                    for col_idx in range(num_cols):
                        cell_text = row_cells[col_idx] if col_idx < len(row_cells) else ""
                        table.cell(row_idx, col_idx).text = cell_text

            elif element == "image":
                alt = unit.get("alt_text", "image")
                para = doc.add_paragraph()
                _embed_image(para, alt, doc)

        except Exception as exc:
            print(
                f"Warning: Failed to write unit {unit_id} ({element}): {exc}",
                file=sys.stderr,
            )

    # -- Insert TOC ------------------------------------------------------------
    _insert_toc(doc)

    # -- Save ------------------------------------------------------------------
    doc.save(str(output_path))
    print(f"Done! {heading_count} headings, {table_count} tables → {output_path}")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_heading_style(doc: Document, level: int) -> str:
    """Get the best available heading style for the given level."""
    for lvl in range(level, 0, -1):
        sname = HEADING_STYLE_MAP.get(lvl, "Normal")
        try:
            doc.styles[sname]
            return sname
        except KeyError:
            continue
    return "Normal"


def _write_mixed_content(para, text: str, doc: Document) -> None:
    """Write paragraph text, handling {{asset:path}} and {{ref:target}} macros.

    Standard text is written directly. {{asset:path}} embeds an image.
    {{ref:target}} renders as blue underlined reference text.
    """
    pattern = re.compile(r"\{\{(asset|ref):([^}]+)\}\}")
    last_end = 0
    has_macro = False

    for m in pattern.finditer(text):
        has_macro = True
        before = text[last_end:m.start()]
        if before:
            para.add_run(before)

        macro_type = m.group(1)
        macro_arg = m.group(2).strip()

        if macro_type == "asset":
            _embed_image(para, macro_arg, doc)
        elif macro_type == "ref":
            run = para.add_run(f"[{macro_arg}]")
            try:
                run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            except Exception:
                pass
            run.font.underline = True

        last_end = m.end()

    if not has_macro:
        para.add_run(text)
    else:
        trailing = text[last_end:]
        if trailing:
            para.add_run(trailing)


def _embed_image(para, filename: str, doc: Document) -> None:
    """Embed an image into a paragraph. Falls back to placeholder text."""
    candidates = [Path(filename)]
    if not candidates[0].is_absolute():
        candidates.append(Path.cwd() / filename)

    img_path: Optional[Path] = None
    for cand in candidates:
        if cand.exists():
            img_path = cand
            break

    if img_path is None:
        run = para.add_run(f"[Image: {filename}]")
        run.font.italic = True
        try:
            run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        except Exception:
            pass
        return

    try:
        run = para.add_run()
        run.add_picture(str(img_path))
    except Exception as exc:
        print(f"Warning: Could not embed {filename}: {exc}", file=sys.stderr)
        run = para.add_run(f"[Image: {filename}]")
        run.font.italic = True


def _insert_toc(doc: Document) -> None:
    """Insert TOC heading and field code at document start."""
    heading_para = doc.add_paragraph()
    heading_para.style = doc.styles["Heading 1"]
    heading_para.add_run("目  录")

    field_para = doc.add_paragraph()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    placeholder = OxmlElement("w:t")
    placeholder.set(qn("xml:space"), "preserve")
    placeholder.text = '[请在 Word 中更新目录 — 右键点击此处选择"更新域"]'

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    r1 = field_para.add_run()
    r1._element.append(fld_begin)
    r2 = field_para.add_run()
    r2._element.append(instr)
    r3 = field_para.add_run()
    r3._element.append(fld_sep)
    r4 = field_para.add_run()
    r4._element.append(placeholder)
    r5 = field_para.add_run()
    r5._element.append(fld_end)

    pb_para = doc.add_paragraph()
    pb_run = pb_para.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    pb_run._element.append(br)

    body = doc.element.body
    body.remove(heading_para._element)
    body.remove(field_para._element)
    body.remove(pb_para._element)
    body.insert(0, pb_para._element)
    body.insert(0, field_para._element)
    body.insert(0, heading_para._element)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)

    if len(args) < 3:
        print(
            "Usage: python3 assemble_docx.py skeleton.json content.json "
            "<template.docx> --output <output.docx>",
            file=sys.stderr,
        )
        sys.exit(1)

    output_path: Optional[Path] = None
    i = 3
    while i < len(args):
        if args[i] == "--output" and i + 1 < len(args):
            output_path = Path(args[i + 1])
            i += 2
        else:
            i += 1

    if output_path is None:
        print("Error: --output is required", file=sys.stderr)
        sys.exit(1)

    assemble(
        Path(args[0]),
        Path(args[1]),
        Path(args[2]),
        output_path,
    )


if __name__ == "__main__":
    main()
