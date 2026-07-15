#!/usr/bin/env python3
"""
assemble_docx.py — Assemble a .docx from skeleton + content + template.

Usage:
    python3 assemble_docx.py skeleton.json content.json <template.docx> --output <output.docx>

How it works (v2 — "in-place modification"):
    1. Copy template → output (inherit all styles, headers, footers, page setup)
    2. Load skeleton.json and content.json
    3. Build a lookup: {para_index: (unit, content_text)}
    4. Iterate body children by index — for each matching unit:
       - fixed:       leave as-is (all formatting preserved)
       - generated:   replace run text in-place (formatting preserved)
       - placeholder: replace placeholder patterns in runs (formatting preserved)
    5. All other body children (no matching unit) are left untouched
    6. Save

Key difference from v1: body is NEVER cleared. Paragraph formatting (indentation,
spacing, alignment, fonts, numbering) is fully preserved because the paragraph
XML element stays intact — only <w:t> text nodes are modified.

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
# Assembly
# ---------------------------------------------------------------------------

def assemble(
    skeleton_path: Path,
    content_path: Path,
    template_path: Path,
    output_path: Path,
) -> None:
    """Main assembly workflow — in-place text replacement."""

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

    # Build content lookup by unit id
    content_lookup: dict[str, str] = {}
    for unit in content_data.get("units", []):
        unit_id = unit.get("id", "")
        unit_content = unit.get("content", "")
        if unit_id and unit_content:
            content_lookup[unit_id] = str(unit_content)

    # -- Build para_index → (unit, content) lookup -----------------------------
    units = skeleton.get("units", [])
    index_map: dict[int, tuple[dict, Optional[str]]] = {}

    for unit in units:
        para_idx = unit.get("para_index")
        if para_idx is None:
            continue
        unit_id = unit.get("id", "")
        content = content_lookup.get(unit_id)
        index_map[para_idx] = (unit, content)

    # -- Copy template to output -----------------------------------------------
    shutil.copy2(str(template_path), str(output_path))
    doc = Document(str(output_path))

    # -- In-place modification of body children --------------------------------
    body = doc.element.body
    children = list(body)

    heading_count = 0
    table_count = 0
    replaced_count = 0

    for para_index, child in enumerate(children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'sectPr':
            continue  # preserve section properties

        unit_info = index_map.get(para_index)
        if unit_info is None:
            continue  # no unit matches — leave as-is

        unit, content = unit_info
        unit_type = unit.get("type", "")
        element = unit.get("element", "")

        if unit_type == "fixed":
            if element == "heading":
                heading_count += 1
            # Fixed: leave completely untouched
            continue

        if tag == 'p':
            if element == "heading":
                heading_count += 1
                # Heading text replacement (though headings are usually fixed)
                if content:
                    _replace_paragraph_text(child, content)
                    replaced_count += 1

            elif element == "paragraph":
                if unit_type == "generated":
                    if content:
                        # Check if content is a JSON 2D array → create real table
                        stripped = content.strip()
                        if stripped.startswith('[[') and stripped.endswith(']]'):
                            try:
                                table_data = json.loads(stripped)
                                if isinstance(table_data, list) and len(table_data) > 0:
                                    _replace_paragraph_with_table(
                                        child, table_data, doc)
                                    table_count += 1
                                    continue
                            except json.JSONDecodeError:
                                pass  # Fall through to text replacement

                        # Handle \n\n paragraph splits
                        if '\n\n' in content:
                            new_paras = _split_paragraphs(child, content, doc)
                            # Insert additional paragraphs after current
                            for new_p in reversed(new_paras[1:]):
                                child.addnext(new_p)
                        else:
                            _replace_paragraph_text(child, content, doc)
                        replaced_count += 1
                    else:
                        # No content provided — leave original text as-is
                        pass

                elif unit_type == "placeholder":
                    if content:
                        key = unit.get("key", "")
                        pattern = unit.get("pattern", f"《{key}》")
                        _replace_placeholder_in_runs(child, pattern, content)
                        replaced_count += 1

            elif element == "placeholder":
                if content:
                    key = unit.get("key", "")
                    pattern = unit.get("pattern", f"《{key}》")
                    _replace_placeholder_in_runs(child, pattern, content)
                    replaced_count += 1

        elif tag == 'tbl':
            if element == "table" and unit_type == "generated":
                if content:
                    table_count += 1
                    _fill_table_cells(child, content)

    # -- Save ------------------------------------------------------------------
    doc.save(str(output_path))
    print(
        f"Done! {heading_count} headings, {table_count} tables, "
        f"{replaced_count} paragraphs replaced → {output_path}"
    )


# ---------------------------------------------------------------------------
# In-place text replacement helpers
# ---------------------------------------------------------------------------

def _replace_paragraph_text(para_element, new_text: str, doc=None) -> None:
    """Replace the text content of a paragraph element while preserving all
    run formatting (fonts, sizes, bold, italic, colors, etc.).

    Strategy:
    - Put all new text into the first run's first <w:t> element
    - Clear text from all other <w:t> elements
    - Handle \\n as line breaks (<w:br/>)
    - Preserve all <w:rPr> (run properties) on every run

    If doc is provided, handles {{asset:path}} and {{ref:target}} macros.
    """
    if doc is not None and ('{{asset:' in new_text or '{{ref:' in new_text):
        _replace_with_macros(para_element, new_text, doc)
        return

    runs = para_element.findall(qn('w:r'))
    if not runs:
        # No runs exist — create one
        run = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = new_text
        run.append(t)
        para_element.append(run)
        return

    # Handle \\n as line breaks within the paragraph
    parts = new_text.split('\n')

    if len(parts) == 1:
        # Simple case: no newlines
        _set_run_text(runs[0], new_text, clear_others=True)
    else:
        # Multi-line: use line breaks within the first run
        first_run = runs[0]
        _clear_run_text_nodes(first_run)
        # Add first part (skip leading empty parts from \\n\\n)
        start_idx = 0
        while start_idx < len(parts) and parts[start_idx] == '':
            start_idx += 1
        if start_idx == len(parts):
            # All parts empty
            _set_run_text(runs[0], '', clear_others=True)
        else:
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = parts[start_idx]
            first_run.append(t)
            # Add breaks and subsequent parts
            for part in parts[start_idx + 1:]:
                br = OxmlElement('w:br')
                first_run.append(br)
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = part
                first_run.append(t)

    # Clear text from all other runs (preserve formatting, remove text)
    for run in runs[1:]:
        _clear_run_text_nodes(run)


def _split_paragraphs(para_element, new_text: str, doc=None) -> list:
    """Split content by \\n\\n into multiple paragraph elements.
    Returns list of paragraph elements (first = modified input, rest = new clones)."""
    paragraphs = [p for p in new_text.split('\n\n') if p.strip()]
    if len(paragraphs) <= 1:
        # No split needed — treat as single paragraph with \\n line breaks
        _replace_paragraph_text(para_element, new_text, doc)
        return [para_element]

    # First paragraph replaces current
    _replace_paragraph_text(para_element, paragraphs[0], doc)
    result = [para_element]

    # Create additional paragraphs from clones
    from copy import deepcopy
    for part in paragraphs[1:]:
        new_para = deepcopy(para_element)
        _replace_paragraph_text(new_para, part, doc)
        result.append(new_para)

    return result


def _replace_paragraph_with_table(para_element, table_data: list, doc) -> None:
    """Replace a paragraph element with a real Word table.

    table_data: 2D list — first row is header, subsequent rows are data.
    The paragraph element is replaced in-place with a <w:tbl> element.
    """
    if not table_data:
        return

    headers = table_data[0] if table_data else []
    data_rows = table_data[1:] if len(table_data) > 1 else []
    all_rows = [headers] + data_rows if headers else data_rows
    if not all_rows:
        return

    num_cols = max(len(r) for r in all_rows)
    num_rows = len(all_rows)

    # Create table element
    tbl = OxmlElement('w:tbl')

    # Table properties
    tblPr = OxmlElement('w:tblPr')
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    tbl.append(tblPr)

    # Table grid (column widths)
    tblGrid = OxmlElement('w:tblGrid')
    col_width = 5000 // num_cols
    for _ in range(num_cols):
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(col_width))
        tblGrid.append(gridCol)
    tbl.append(tblGrid)

    # Table rows
    for row_idx, row_data in enumerate(all_rows):
        tr = OxmlElement('w:tr')
        for col_idx in range(num_cols):
            cell_text = str(row_data[col_idx]) if col_idx < len(row_data) else ''
            tc = OxmlElement('w:tc')

            # Cell properties
            tcPr = OxmlElement('w:tcPr')
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(col_width))
            tcW.set(qn('w:type'), 'pct')
            tcPr.append(tcW)
            # Header row shading
            if row_idx == 0:
                shading = OxmlElement('w:shd')
                shading.set(qn('w:val'), 'clear')
                shading.set(qn('w:color'), 'auto')
                shading.set(qn('w:fill'), 'D9E2F3')
                tcPr.append(shading)
            tc.append(tcPr)

            # Paragraph in cell
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            # Center align header, left align data
            if row_idx == 0:
                jc = OxmlElement('w:jc')
                jc.set(qn('w:val'), 'center')
                pPr.append(jc)
            p.append(pPr)

            r = OxmlElement('w:r')
            # Bold for header
            if row_idx == 0:
                rPr = OxmlElement('w:rPr')
                b = OxmlElement('w:b')
                rPr.append(b)
                r.append(rPr)

            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = cell_text
            r.append(t)
            p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)

    # Replace paragraph with table
    parent = para_element.getparent()
    if parent is not None:
        parent.replace(para_element, tbl)


def _set_run_text(run, text: str, clear_others: bool = False) -> None:
    """Set text on a single run's first <w:t>, clearing other <w:t> elements."""
    t_elements = run.findall(qn('w:t'))
    if t_elements:
        t_elements[0].text = text
        t_elements[0].set(qn('xml:space'), 'preserve')
        for t in t_elements[1:]:
            t.text = ''
    else:
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        run.append(t)


def _clear_run_text_nodes(run) -> None:
    """Clear text from all <w:t> elements in a run (preserving the run itself)."""
    for t in run.findall(qn('w:t')):
        t.text = ''


def _replace_placeholder_in_runs(para_element, pattern: str, replacement: str) -> None:
    """Replace a placeholder pattern in paragraph runs while preserving formatting.

    Searches for `pattern` in the concatenated text of all runs and replaces it
    with `replacement`. The replacement text inherits the formatting of the
    run where the pattern was found.
    """
    runs = para_element.findall(qn('w:r'))
    if not runs:
        return

    # Collect text from all runs
    run_data: list[tuple[Any, str, list[Any]]] = []
    # (run, full_text, [t_element, ...])
    full_text = ""
    for run in runs:
        t_elements = run.findall(qn('w:t'))
        run_text = ''.join(t.text or '' for t in t_elements)
        run_data.append((run, run_text, t_elements))
        full_text += run_text

    if pattern not in full_text:
        return

    # Find which run contains the pattern
    new_full = full_text.replace(pattern, replacement)

    # Strategy: put all new text in first run, clear others
    first_run = runs[0]
    _clear_run_text_nodes(first_run)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = new_full
    first_run.append(t)

    for run in runs[1:]:
        _clear_run_text_nodes(run)


def _replace_with_macros(para_element, text: str, doc: Document) -> None:
    """Replace paragraph text with content that includes {{asset:path}} or
    {{ref:target}} macros. Preserves formatting from the first run.

    This creates new runs for each text segment and macro, using the first
    run's properties as the base style. Original formatting from subsequent
    runs is lost — this is a trade-off for macro support.
    """
    macro_pattern = re.compile(r"\{\{(asset|ref):([^}]+)\}\}")

    runs = para_element.findall(qn('w:r'))
    first_run = runs[0] if runs else None

    # Save first run's properties for reuse
    first_rPr = None
    if first_run is not None:
        first_rPr = first_run.find(qn('w:rPr'))
        # Deep copy if exists (simplified: just copy the element)
        if first_rPr is not None:
            from copy import deepcopy
            first_rPr = deepcopy(first_rPr)

    # Clear all existing runs
    for run in runs:
        para_element.remove(run)

    # Create new runs for each segment
    last_end = 0
    for m in macro_pattern.finditer(text):
        before = text[last_end:m.start()]
        if before:
            run = OxmlElement('w:r')
            if first_rPr is not None:
                from copy import deepcopy
                run.append(deepcopy(first_rPr))
            _add_text_to_run(run, before)
            para_element.append(run)

        macro_type = m.group(1)
        macro_arg = m.group(2).strip()

        if macro_type == "asset":
            _add_image_run(para_element, macro_arg, doc, first_rPr)
        elif macro_type == "ref":
            run = OxmlElement('w:r')
            if first_rPr is not None:
                from copy import deepcopy
                run.append(deepcopy(first_rPr))
            _add_text_to_run(run, f"[{macro_arg}]")
            para_element.append(run)

        last_end = m.end()

    trailing = text[last_end:]
    if trailing:
        run = OxmlElement('w:r')
        if first_rPr is not None:
            from copy import deepcopy
            run.append(deepcopy(first_rPr))
        _add_text_to_run(run, trailing)
        para_element.append(run)


def _add_text_to_run(run, text: str) -> None:
    """Add <w:t> with text to a run element."""
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)


def _add_image_run(para_element, filename: str, doc: Document, base_rPr=None) -> None:
    """Add an image run to a paragraph element. Falls back to text on failure."""
    candidates = [Path(filename)]
    if not candidates[0].is_absolute():
        candidates.append(Path.cwd() / filename)

    img_path: Optional[Path] = None
    for cand in candidates:
        if cand.exists():
            img_path = cand
            break

    if img_path is None:
        # Fallback: text placeholder
        run = OxmlElement('w:r')
        if base_rPr is not None:
            from copy import deepcopy
            run.append(deepcopy(base_rPr))
        _add_text_to_run(run, f"[Image: {filename}]")
        para_element.append(run)
        return

    try:
        # Read image and determine dimensions
        from PIL import Image as PILImage
        from docx.shared import Emu

        img = PILImage.open(str(img_path))
        width_px, height_px = img.size
        img_dpi = img.info.get('dpi', (96, 96))
        dpi_x = img_dpi[0] if img_dpi[0] else 96

        # Convert to EMU (English Metric Units)
        max_width_emu = Emu(15240000)  # ~16cm max width
        width_emu = int(width_px / dpi_x * 914400)
        if width_emu > max_width_emu:
            ratio = max_width_emu / width_emu
            width_emu = max_width_emu
            height_px = int(height_px * ratio)
        height_emu = int(height_px / dpi_x * 914400)

        # Build drawing XML
        nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        }

        run = OxmlElement('w:r')
        if base_rPr is not None:
            from copy import deepcopy
            run.append(deepcopy(base_rPr))

        drawing = OxmlElement('w:drawing')
        inline = OxmlElement('wp:inline')
        inline.set('distT', '0')
        inline.set('distB', '0')
        inline.set('distL', '0')
        inline.set('distR', '0')

        extent = OxmlElement('wp:extent')
        extent.set('cx', str(width_emu))
        extent.set('cy', str(height_emu))
        inline.append(extent)

        effectExtent = OxmlElement('wp:effectExtent')
        effectExtent.set('l', '0')
        effectExtent.set('t', '0')
        effectExtent.set('r', '0')
        effectExtent.set('b', '0')
        inline.append(effectExtent)

        docPr = OxmlElement('wp:docPr')
        docPr.set('id', '1')
        docPr.set('name', Path(filename).name)
        inline.append(docPr)

        cNvGraphicFramePr = OxmlElement('wp:cNvGraphicFramePr')
        graphicFrameLocks = OxmlElement('a:graphicFrameLocks')
        graphicFrameLocks.set('noChangeAspect', '1')
        cNvGraphicFramePr.append(graphicFrameLocks)
        inline.append(cNvGraphicFramePr)

        graphic = OxmlElement('a:graphic')
        graphicData = OxmlElement('a:graphicData')
        graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

        pic = OxmlElement('pic:pic')
        nvPicPr = OxmlElement('pic:nvPicPr')
        cNvPr = OxmlElement('pic:cNvPr')
        cNvPr.set('id', '0')
        cNvPr.set('name', Path(filename).name)
        nvPicPr.append(cNvPr)
        cNvPicPr = OxmlElement('pic:cNvPicPr')
        nvPicPr.append(cNvPicPr)
        pic.append(nvPicPr)

        blipFill = OxmlElement('pic:blipFill')
        blip = OxmlElement('a:blip')
        blip.set('r:embed', 'rId1')  # Simplified — in production this needs relationship management
        blipFill.append(blip)
        stretch = OxmlElement('a:stretch')
        fillRect = OxmlElement('a:fillRect')
        stretch.append(fillRect)
        blipFill.append(stretch)
        pic.append(blipFill)

        spPr = OxmlElement('pic:spPr')
        xfrm = OxmlElement('a:xfrm')
        off = OxmlElement('a:off')
        off.set('x', '0')
        off.set('y', '0')
        xfrm.append(off)
        ext = OxmlElement('a:ext')
        ext.set('cx', str(width_emu))
        ext.set('cy', str(height_emu))
        xfrm.append(ext)
        spPr.append(xfrm)
        prstGeom = OxmlElement('a:prstGeom')
        prstGeom.set('prst', 'rect')
        avLst = OxmlElement('a:avLst')
        prstGeom.append(avLst)
        spPr.append(prstGeom)
        pic.append(spPr)

        graphicData.append(pic)
        graphic.append(graphicData)
        inline.append(graphic)
        drawing.append(inline)
        run.append(drawing)
        para_element.append(run)

    except Exception as exc:
        print(f"Warning: Could not embed {filename}: {exc}", file=sys.stderr)
        run = OxmlElement('w:r')
        if base_rPr is not None:
            from copy import deepcopy
            run.append(deepcopy(base_rPr))
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = f"[Image: {filename}]"
        run.append(t)
        para_element.append(run)


def _fill_table_cells(table_element, content) -> None:
    """Fill table cells with content data, preserving table structure and formatting.

    `content` can be:
    - A JSON string (2D array): [["cell", ...], ...]
    - A plain string: placed in the first data cell
    """
    rows = table_element.findall(qn('w:tr'))
    if not rows:
        return

    # Parse content
    data: list[list[str]] = []
    if isinstance(content, str):
        try:
            parsed = json.loads(content)
            if isinstance(parsed, list):
                data = [[str(c) for c in row] for row in parsed]
        except json.JSONDecodeError:
            data = [[content]]
    elif isinstance(content, list):
        data = [[str(c) for c in row] for row in content]

    if not data:
        return

    # Skip header row (first row), fill data starting from row 1
    for data_row_idx, row_cells in enumerate(data):
        table_row_idx = data_row_idx + 1  # +1 to skip header
        if table_row_idx >= len(rows):
            break

        row_element = rows[table_row_idx]
        cells = row_element.findall(qn('w:tc'))
        for col_idx, cell_text in enumerate(row_cells):
            if col_idx >= len(cells):
                break
            cell = cells[col_idx]
            # Find first paragraph in cell, replace its text
            first_para = cell.find(qn('w:p'))
            if first_para is not None:
                _replace_paragraph_text(first_para, cell_text)


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
