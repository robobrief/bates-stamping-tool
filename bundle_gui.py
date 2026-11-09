import csv
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Dict, List, Optional, Sequence, Tuple

try:
    from pypdf import PdfReader as PypdfReader
    from pypdf import PdfWriter
    from pypdf.constants import PageLabelStyle
    from pypdf.generic import NameObject
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing required Python dependencies for the bundle GUI.\n"
        "Install them with:\n"
        "  python -m pip install -r requirements.txt\n"
        "or install the packages manually:\n"
        "  python -m pip install PyPDF2 pypdf reportlab python-docx pymupdf pikepdf"
    ) from exc


REQUIRED_COLUMNS = ["Citation", "Case title", "C ¶", "D2 ¶", "Page"]
INDEX_TABLE_PLACEHOLDER_CAPTION = "[[INDEX_TABLE_PLACEHOLDER]]"
APP_NAME = "LawPage PDF Bundling Tool"
APP_VERSION = "v1.0.0"
STAMP_COLOR_PRESET_OPTIONS = [
    ("black", "Black, no background (text #000000, bg none)"),
    ("white", "White, no background (text #FFFFFF, bg none)"),
    ("red", "Red, no background (text #FF0000, bg none)"),
    ("black-box", "Black on white (text #000000, bg #FFFFFF)"),
    ("white-box", "White on black (text #FFFFFF, bg #000000)"),
    ("red-box", "Red on white (text #FF0000, bg #FFFFFF)"),
    ("reverse-red", "White on red (text #FFFFFF, bg #FF0000)"),
    ("blue-box", "Blue on white (text #0000FF, bg #FFFFFF)"),
    ("reverse-blue", "White on blue (text #FFFFFF, bg #0000FF)"),
    ("highlight", "Black on yellow (text #000000, bg #FFFF00)"),
]
STAMP_COLOR_PRESET_KEY_TO_DISPLAY = {k: d for k, d in STAMP_COLOR_PRESET_OPTIONS}
STAMP_COLOR_PRESET_DISPLAY_TO_KEY = {d: k for k, d in STAMP_COLOR_PRESET_OPTIONS}


@dataclass
class IndexRow:
    citation: str
    case_title: str
    c_para: str
    d2_para: str
    page: str


@dataclass
class PaginationRow:
    entry: int
    citation: str
    case_title: str
    filename: str
    source_pages: int
    start_page: int
    end_page: int


@dataclass
class IndexEntry:
    doc_index: str
    citation: str
    description: str
    c_para: str
    d2_para: str
    page: str
    pdf_page_index: str


def normalize_text(value: str) -> str:
    text = (value or "").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_header(value: str) -> str:
    return "".join(ch.lower() for ch in normalize_text(value) if ch.isalnum())


def bookmark_label(citation: str, case_title: str) -> str:
    c = normalize_text(citation)
    t = normalize_text(case_title)
    if c and t:
        if c.lower() == t.lower():
            return t
        return f"{c}:: {t}"
    return t or c


def color_preset_display_from_key(value: str) -> str:
    v = normalize_text(value)
    if v in STAMP_COLOR_PRESET_KEY_TO_DISPLAY:
        return STAMP_COLOR_PRESET_KEY_TO_DISPLAY[v]
    if v in STAMP_COLOR_PRESET_DISPLAY_TO_KEY:
        return v
    return STAMP_COLOR_PRESET_KEY_TO_DISPLAY["reverse-blue"]


def color_preset_key_from_display(value: str) -> str:
    v = normalize_text(value)
    if v in STAMP_COLOR_PRESET_KEY_TO_DISPLAY:
        return v
    if v in STAMP_COLOR_PRESET_DISPLAY_TO_KEY:
        return STAMP_COLOR_PRESET_DISPLAY_TO_KEY[v]
    for key, _display in STAMP_COLOR_PRESET_OPTIONS:
        if v.lower().startswith(key.lower()):
            return key
    return "reverse-blue"


def description_from_pdf_path(path: Path) -> str:
    return normalize_text(path.stem)


def read_index_rows(
    docx_path: Path,
    *,
    page_label_column_title: str = "Page Label",
    pdf_page_column_title: str = "PDF Page",
) -> List[IndexRow]:
    doc = Document(str(docx_path))
    rows: List[IndexRow] = []
    page_label_header = normalize_header(page_label_column_title)
    pdf_page_header = normalize_header(pdf_page_column_title)

    for table in doc.tables:
        if not table.rows:
            continue
        raw_headers = [normalize_text(c.text) for c in table.rows[0].cells]
        headers = [normalize_header(c.text) for c in table.rows[0].cells]
        col_map: Dict[str, int] = {}
        for idx, h in enumerate(headers):
            if h == "citation" and "Citation" not in col_map:
                col_map["Citation"] = idx
            elif h in {"casetitle", "description"} and "Case title" not in col_map:
                col_map["Case title"] = idx
            elif h.startswith("d2") and "D2 ¶" not in col_map:
                col_map["D2 ¶"] = idx
            elif h.startswith("c") and h != "citation" and "C ¶" not in col_map:
                col_map["C ¶"] = idx
            elif (
                h == "page"
                or h == "pagelabel"
                or h == page_label_header
                or (h.endswith("page") and h != "pdfpage" and h != pdf_page_header)
            ) and "Page" not in col_map:
                col_map["Page"] = idx
        has_text_col = ("Citation" in col_map) or ("Case title" in col_map)
        if not has_text_col:
            continue

        for row in table.rows[1:]:
            cells = row.cells
            citation = normalize_text(cells[col_map["Citation"]].text) if "Citation" in col_map else ""
            case_title = normalize_text(cells[col_map["Case title"]].text) if "Case title" in col_map else ""
            c_para = normalize_text(cells[col_map["C ¶"]].text) if "C ¶" in col_map else ""
            d2_para = normalize_text(cells[col_map["D2 ¶"]].text) if "D2 ¶" in col_map else ""
            page = normalize_text(cells[col_map["Page"]].text) if "Page" in col_map else ""
            if not citation and case_title:
                citation = case_title
            if not case_title and citation:
                case_title = citation
            if not citation and not case_title:
                continue
            rows.append(IndexRow(citation, case_title, c_para, d2_para, page))

    # Fallback parse for generated/customized index tables where core columns were renamed.
    if not rows:
        for table in doc.tables:
            if not table.rows or len(table.rows) < 2:
                continue
            headers = [normalize_header(c.text) for c in table.rows[0].cells]
            if not headers:
                continue

            desc_idx: Optional[int] = None
            for token in ("description", "casetitle", "citation"):
                if token in headers:
                    desc_idx = headers.index(token)
                    break
            if desc_idx is None:
                # Common generated layout: "#" then description.
                if headers and headers[0] == "" and len(headers) > 1:
                    desc_idx = 1
                elif headers and headers[0] == "description":
                    desc_idx = 0
                elif raw_headers and raw_headers[0] == "#":
                    desc_idx = 1 if len(headers) > 1 else 0
                else:
                    desc_idx = 0

            page_idx: Optional[int] = None
            if "page" in headers:
                page_idx = headers.index("page")
            elif "pagelabel" in headers:
                page_idx = headers.index("pagelabel")
            elif page_label_header and page_label_header in headers:
                page_idx = headers.index(page_label_header)
            else:
                for i, h in enumerate(headers):
                    if h.endswith("page") and h != "pdfpage" and h != pdf_page_header:
                        page_idx = i
                        break
            if page_idx is None and desc_idx + 1 < len(headers):
                page_idx = desc_idx + 1

            for row in table.rows[1:]:
                cells = row.cells
                if desc_idx >= len(cells):
                    continue
                desc = normalize_text(cells[desc_idx].text)
                if not desc:
                    continue
                page = normalize_text(cells[page_idx].text) if page_idx is not None and page_idx < len(cells) else ""
                rows.append(IndexRow(desc, desc, "", "", page))

    # keep first occurrence only (DOCX can contain repeated index tables)
    unique: List[IndexRow] = []
    seen = set()
    for row in rows:
        key = (row.citation, row.case_title, row.c_para, row.d2_para, row.page)
        if key in seen:
            continue
        seen.add(key)
        unique.append(row)
    if not unique:
        raise RuntimeError("No usable index table found in DOCX.")
    return unique


def make_rows_from_pdf_names(pdf_paths: Sequence[Path]) -> List[IndexRow]:
    rows: List[IndexRow] = []
    for p in pdf_paths:
        stem = description_from_pdf_path(p)
        rows.append(IndexRow(citation=stem, case_title=stem, c_para="", d2_para="", page=""))
    return rows


def merge_bundle_from_ordered_list(
    pdf_paths: Sequence[Path],
    index_rows: Sequence[IndexRow],
    output_pdf: Path,
    log: Callable[[str], None],
    *,
    preserve_source_bookmarks: bool = True,
) -> List[PaginationRow]:
    if len(pdf_paths) != len(index_rows):
        raise RuntimeError(
            f"PDF count ({len(pdf_paths)}) must equal index row count ({len(index_rows)})."
        )

    writer = PdfWriter()
    pagination: List[PaginationRow] = []
    current_page = 1
    navigate_parent = writer.add_outline_item("NAVIGATE BY DOCUMENT DESCRIPTION", 0)

    for i, (pdf_path, row) in enumerate(zip(pdf_paths, index_rows), start=1):
        if not pdf_path.is_file():
            raise RuntimeError(f"PDF not found: {pdf_path}")
        reader = PypdfReader(str(pdf_path))
        page_count = len(reader.pages)
        start = current_page
        end = start + page_count - 1

        for p in reader.pages:
            writer.add_page(p)

        document_label = bookmark_label(row.citation, row.case_title)
        doc_parent = writer.add_outline_item(
            document_label,
            start - 1,
            parent=navigate_parent,
        )
        if preserve_source_bookmarks:
            copy_outline_shifted(reader, writer, start - 1, parent=doc_parent)

        pagination.append(
            PaginationRow(
                entry=i,
                citation=row.citation,
                case_title=row.case_title,
                filename=pdf_path.name,
                source_pages=page_count,
                start_page=start,
                end_page=end,
            )
        )
        log(f"Merged {i}/{len(pdf_paths)}: {pdf_path.name} ({page_count} pages)")
        current_page = end + 1

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open("wb") as f:
        writer.write(f)
    return pagination


def _is_index_table(table) -> bool:
    if not table.rows:
        return False
    headers = [normalize_header(c.text) for c in table.rows[0].cells]
    has_desc = ("description" in headers) or ("casetitle" in headers)
    if not has_desc:
        return False
    if len(headers) < 2:
        return False
    # Treat any 2+ column table with a Description-like header as an index table.
    # This allows user-customized page header labels to still be recognized/removed.
    return True


def _remove_table(table) -> None:
    tbl = table._element
    parent = tbl.getparent()
    if parent is not None:
        parent.remove(tbl)


def _table_caption_text(table) -> str:
    try:
        tbl_pr = table._tbl.tblPr
        if tbl_pr is None:
            return ""
        tbl_caption = tbl_pr.find(qn("w:tblCaption"))
        if tbl_caption is None:
            return ""
        return normalize_text(tbl_caption.get(qn("w:val"), ""))
    except Exception:
        return ""


def _is_placeholder_seed_table(table) -> bool:
    caption = _table_caption_text(table)
    if caption and normalize_text(caption) == INDEX_TABLE_PLACEHOLDER_CAPTION:
        return True
    # Fallback: allow marker in first cell for templates that don't persist w:tblCaption cleanly.
    try:
        if table.rows and table.rows[0].cells:
            marker = normalize_text(table.rows[0].cells[0].text)
            if marker == INDEX_TABLE_PLACEHOLDER_CAPTION:
                return True
        tbl_elm = table._element
        prev_elm = tbl_elm.getprevious()
        if prev_elm is not None and prev_elm.tag.endswith("}p"):
            prev_text = _extract_paragraph_text_from_element(prev_elm)
            if _is_placeholder_caption_text(prev_text):
                return True
    except Exception:
        return False
    return False


def _extract_paragraph_text_from_element(p_elm) -> str:
    try:
        parts: List[str] = []
        t_tag = qn("w:t")
        for el in p_elm.iter():
            if getattr(el, "tag", None) == t_tag:
                parts.append(getattr(el, "text", "") or "")
        return normalize_text("".join(parts))
    except Exception:
        return ""


def _is_placeholder_caption_text(text: str) -> bool:
    t = normalize_text(text)
    marker = INDEX_TABLE_PLACEHOLDER_CAPTION
    return bool(t) and (t == marker or marker in t)


def _remove_element(elm) -> None:
    parent = elm.getparent()
    if parent is not None:
        parent.remove(elm)


def _remove_adjacent_placeholder_caption_paragraphs(table) -> int:
    removed = 0
    tbl_elm = table._element
    prev_elm = tbl_elm.getprevious()
    if prev_elm is not None and prev_elm.tag.endswith("}p"):
        prev_text = _extract_paragraph_text_from_element(prev_elm)
        if _is_placeholder_caption_text(prev_text):
            _remove_element(prev_elm)
            removed += 1
    return removed


def _set_table_borders_thin(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_borders = tbl_pr.find(qn("w:tblBorders"))
    if tbl_borders is None:
        tbl_borders = OxmlElement("w:tblBorders")
        tbl_pr.append(tbl_borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        element = tbl_borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tbl_borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "B0B0B0")


def discover_table_styles(doc_path: Optional[Path]) -> List[str]:
    candidates = []
    doc = None
    try:
        if doc_path is not None and doc_path.is_file():
            doc = Document(str(doc_path))
        else:
            doc = Document()
        for style in doc.styles:
            try:
                if style.type == WD_STYLE_TYPE.TABLE:
                    name = normalize_text(style.name)
                    if name:
                        candidates.append(name)
            except Exception:
                continue
    except Exception:
        return []
    finally:
        doc = None
    return sorted(set(candidates), key=lambda s: s.lower())


def filter_table_styles_by_family(style_names: Sequence[str], family: str) -> List[str]:
    family_norm = normalize_text(family).lower()
    if family_norm == "plain":
        filtered = [
            s for s in style_names
            if "plain table" in s.lower() and ("accent" in s.lower() or re.search(r"plain table\s+\d", s.lower()))
        ]
    else:
        filtered = [
            s for s in style_names
            if "grid table" in s.lower() and ("accent" in s.lower() or re.search(r"grid table\s+\d", s.lower()))
        ]
    if not filtered:
        token = "plain table" if family_norm == "plain" else "grid table"
        filtered = [s for s in style_names if token in s.lower()]
    return sorted(set(filtered), key=lambda s: s.lower())


def _set_table_fixed_layout(table, total_width_cm: float) -> None:
    tbl_pr = table._tbl.tblPr
    tbl_layout = tbl_pr.find(qn("w:tblLayout"))
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed")

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    width_twips = int(max(1.0, total_width_cm) * 566.929)  # cm -> twips
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(width_twips))


def _set_repeat_header_row(table) -> None:
    if not table.rows:
        return
    tr = table.rows[0]._tr
    tr_pr = tr.get_or_add_trPr()
    tbl_header = tr_pr.find(qn("w:tblHeader"))
    if tbl_header is None:
        tbl_header = OxmlElement("w:tblHeader")
        tr_pr.append(tbl_header)
    tbl_header.set(qn("w:val"), "true")


def _set_table_look_options(
    table,
    *,
    header_row: bool,
    total_row: bool,
    banded_rows: bool,
    first_column: bool,
    last_column: bool,
    banded_columns: bool,
) -> None:
    tbl_pr = table._tbl.tblPr
    tbl_look = tbl_pr.find(qn("w:tblLook"))
    if tbl_look is None:
        tbl_look = OxmlElement("w:tblLook")
        tbl_pr.append(tbl_look)
    tbl_look.set(qn("w:firstRow"), "1" if header_row else "0")
    tbl_look.set(qn("w:lastRow"), "1" if total_row else "0")
    tbl_look.set(qn("w:firstColumn"), "1" if first_column else "0")
    tbl_look.set(qn("w:lastColumn"), "1" if last_column else "0")
    tbl_look.set(qn("w:noHBand"), "0" if banded_rows else "1")
    tbl_look.set(qn("w:noVBand"), "0" if banded_columns else "1")
    # Keep legacy bitmask in sync for clients that key off w:val.
    val = 0
    if header_row:
        val |= 0x0020
    if total_row:
        val |= 0x0040
    if first_column:
        val |= 0x0080
    if last_column:
        val |= 0x0100
    if not banded_rows:
        val |= 0x0200
    if not banded_columns:
        val |= 0x0400
    tbl_look.set(qn("w:val"), f"{val:04X}")


def _set_cell_vertical_align_center(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    valign = tc_pr.find(qn("w:vAlign"))
    if valign is None:
        valign = OxmlElement("w:vAlign")
        tc_pr.append(valign)
    valign.set(qn("w:val"), "center")


def apply_table_style(table, document: Document, style_name: str) -> Tuple[str, bool]:
    requested = normalize_text(style_name)
    fallback = "Table Grid"
    available = {normalize_text(s.name): s.name for s in document.styles if getattr(s, "type", None) == WD_STYLE_TYPE.TABLE}

    def resolve(name: str) -> Optional[str]:
        key = normalize_text(name)
        if key in available:
            return available[key]
        return None

    resolved = resolve(requested) if requested else None
    used_fallback = False
    if resolved is None:
        resolved = resolve(fallback)
        used_fallback = True
    if resolved is None:
        # Last resort: do not fail; keep existing style.
        return ("", True)
    table.style = resolved
    return (resolved, used_fallback)


def apply_index_table_format(table, document: Document, table_format: Dict[str, object]) -> Optional[str]:
    style_name = str(table_format.get("style", "") or "")
    resolved_style, used_fallback = apply_table_style(table, document, style_name)
    _set_table_look_options(
        table,
        header_row=bool(table_format.get("header_row", True)),
        total_row=bool(table_format.get("total_row", False)),
        banded_rows=bool(table_format.get("banded_rows", False)),
        first_column=bool(table_format.get("first_column", False)),
        last_column=bool(table_format.get("last_column", False)),
        banded_columns=bool(table_format.get("banded_columns", False)),
    )
    if bool(table_format.get("header_row", True)):
        _set_repeat_header_row(table)
    if used_fallback and normalize_text(style_name):
        return f"Requested table style '{style_name}' unavailable; using '{resolved_style or 'existing'}'."
    return None


def _apply_caption_style(paragraph) -> None:
    for candidate in ("Caption", "Table Caption"):
        try:
            paragraph.style = candidate
            return
        except Exception:
            continue
    for run in paragraph.runs:
        run.bold = True


def _add_index_table(
    doc: Document,
    title: str,
    entries: Sequence[IndexEntry],
    style_name: Optional[str],
    table_format: Dict[str, object],
    page_column_width_cm: float,
    add_page_break_before: bool,
    page_label_column_title: str,
    pdf_page_column_title: str,
    include_doc_index_column: bool,
    include_pdf_zero_index: bool,
) -> Optional[str]:
    if add_page_break_before:
        doc.add_page_break()
    title_paragraph = doc.add_paragraph(title)
    _apply_caption_style(title_paragraph)
    column_count = 2 + (1 if include_doc_index_column else 0) + (1 if include_pdf_zero_index else 0)
    table = doc.add_table(rows=1, cols=column_count)
    requested_style = str(table_format.get("style", "") or "").strip()
    style_to_use = requested_style or normalize_text(style_name or "")
    if style_to_use:
        table_format = dict(table_format)
        table_format["style"] = style_to_use
    warning = apply_index_table_format(table, doc, table_format)
    table.autofit = False
    headers: List[str] = []
    if include_doc_index_column:
        headers.append("#")
    page_label_text = normalize_text(page_label_column_title) or "Page Label"
    pdf_page_text = normalize_text(pdf_page_column_title) or "PDF Page"
    headers.extend(["Description", page_label_text])
    if include_pdf_zero_index:
        headers.append(pdf_page_text)
    for i, h in enumerate(headers):
        table.cell(0, i).text = h
    try:
        page_width = max(1.0, float(page_column_width_cm))
    except Exception:
        page_width = 2.0
    try:
        section = doc.sections[-1]
        usable_width_cm = (section.page_width - section.left_margin - section.right_margin) / Cm(1)
    except Exception:
        usable_width_cm = 16.0
    desc_min_cm = 3.0
    _set_table_fixed_layout(table, usable_width_cm)
    right_col_count = 1 + (1 if include_doc_index_column else 0) + (1 if include_pdf_zero_index else 0)
    max_right_width_each = max(1.0, (usable_width_cm - desc_min_cm) / right_col_count)
    page_width = min(page_width, max_right_width_each)
    doc_index_width = page_width if include_doc_index_column else 0.0
    pdf_page_width = page_width if include_pdf_zero_index else 0.0
    desc_width = max(desc_min_cm, usable_width_cm - page_width - doc_index_width - pdf_page_width)

    col = 0
    if include_doc_index_column:
        table.columns[col].width = Cm(doc_index_width)
        col += 1
    table.columns[col].width = Cm(desc_width)
    col += 1
    table.columns[col].width = Cm(page_width)
    col += 1
    if include_pdf_zero_index:
        table.columns[col].width = Cm(pdf_page_width)

    for entry in entries:
        row = table.add_row().cells
        col = 0
        if include_doc_index_column:
            row[col].text = entry.doc_index
            row[col].width = Cm(doc_index_width)
            col += 1
        row[col].text = entry.description
        row[col].width = Cm(desc_width)
        col += 1
        row[col].text = entry.page
        row[col].width = Cm(page_width)
        col += 1
        if include_pdf_zero_index:
            row[col].text = entry.pdf_page_index
            row[col].width = Cm(pdf_page_width)
    header_paragraphs = []
    for cell in table.rows[0].cells:
        header_paragraphs.extend(cell.paragraphs)
    for p in header_paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
    for cell in table.rows[0].cells:
        _set_cell_vertical_align_center(cell)
    for ridx, row in enumerate(table.rows[1:], start=1):
        for cidx, cell in enumerate(row.cells):
            # Description column stays left-aligned; all other columns centered.
            desc_col_idx = 1 if include_doc_index_column else 0
            if cidx == desc_col_idx:
                continue
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_table_borders_thin(table)
    doc.add_paragraph("")
    return warning


def _format_index_page_value(
    start_page: int,
    end_page: int,
    include_page_range: bool,
    *,
    stamp_enabled: bool,
    stamp_number_source: str,
    stamp_prefix: str,
    index_page_offset: int,
) -> str:
    range_sep = "–"
    if not stamp_enabled:
        return f"{start_page}{range_sep}{end_page}" if include_page_range else str(start_page)

    prefix = stamp_prefix or ""
    source = normalize_text(stamp_number_source).lower()

    def as_label(page_num: int) -> str:
        if source == "sequential":
            return f"{prefix}{page_num + int(index_page_offset):06d}"
        return f"{prefix}{page_num}"

    if include_page_range:
        return f"{as_label(start_page)}{range_sep}{as_label(end_page)}"
    return as_label(start_page)


def generate_index_docx_from_template(
    template_docx: Path,
    output_docx: Path,
    pdf_paths: Sequence[Path],
    index_rows: Sequence[IndexRow],
    pagination_rows: Sequence[PaginationRow],
    include_description_sorted: bool,
    use_existing_index_descriptions: bool,
    include_page_range: bool,
    page_column_width_cm: float,
    table_format: Dict[str, object],
    start_each_table_on_new_page: bool,
    page_label_column_title: str,
    pdf_page_column_title: str,
    include_doc_index_column: bool,
    include_pdf_zero_index: bool,
    pdf_zero_index_offset: int = 0,
    stamp_enabled: bool = False,
    stamp_number_source: str = "page-label",
    stamp_prefix: str = "",
    stamp_index_page_offset: int = 0,
) -> Tuple[int, int, List[str]]:
    if len(pdf_paths) != len(pagination_rows):
        raise RuntimeError(
        "Cannot generate index from template: PDF list and pagination rows differ in count."
        )
    if use_existing_index_descriptions and len(index_rows) != len(pagination_rows):
        raise RuntimeError(
            "Cannot use existing index descriptions: index metadata rows and pagination rows differ in count."
        )

    entries: List[IndexEntry] = []
    for idx, (pdf_path, page) in enumerate(zip(pdf_paths, pagination_rows)):
        description = description_from_pdf_path(pdf_path)
        if use_existing_index_descriptions:
            row = index_rows[idx]
            description = normalize_text(row.case_title) or normalize_text(row.citation) or description
        page_value = _format_index_page_value(
            page.start_page,
            page.end_page,
            include_page_range,
            stamp_enabled=stamp_enabled,
            stamp_number_source=stamp_number_source,
            stamp_prefix=stamp_prefix,
            index_page_offset=stamp_index_page_offset,
        )
        entries.append(
            IndexEntry(
                doc_index=str(idx + 1),
                citation=description,
                description=description,
                c_para="",
                d2_para="",
                page=page_value,
                pdf_page_index=str(page.start_page + int(pdf_zero_index_offset)),
            )
        )

    doc = Document(str(template_docx))
    index_tables = [t for t in doc.tables if _is_index_table(t)]
    placeholder_tables = [t for t in doc.tables if _is_placeholder_seed_table(t)]
    table_style = None
    if placeholder_tables:
        try:
            table_style = placeholder_tables[0].style.name
        except Exception:
            table_style = None
    elif index_tables:
        try:
            table_style = index_tables[0].style.name
        except Exception:
            table_style = None
    for t in placeholder_tables:
        _remove_adjacent_placeholder_caption_paragraphs(t)
        _remove_table(t)
    for t in index_tables:
        _remove_table(t)

    warnings: List[str] = []
    warning = _add_index_table(
        doc,
        "Index (by page label)",
        entries,
        table_style,
        table_format,
        page_column_width_cm,
        add_page_break_before=False,
        page_label_column_title=page_label_column_title,
        pdf_page_column_title=pdf_page_column_title,
        include_doc_index_column=include_doc_index_column,
        include_pdf_zero_index=include_pdf_zero_index,
    )
    if warning:
        warnings.append(warning)
    sections_added = 1
    if include_description_sorted:
        by_desc = sorted(entries, key=lambda e: e.description.lower())
        warning = _add_index_table(
            doc,
            "Index (by document description)",
            by_desc,
            table_style,
            table_format,
            page_column_width_cm,
            add_page_break_before=start_each_table_on_new_page and sections_added > 0,
            page_label_column_title=page_label_column_title,
            pdf_page_column_title=pdf_page_column_title,
            include_doc_index_column=include_doc_index_column,
            include_pdf_zero_index=include_pdf_zero_index,
        )
        if warning:
            warnings.append(warning)
        sections_added += 1

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return len(entries), sections_added, warnings


def generate_index_style_preview_docx(
    output_docx: Path,
    *,
    include_description_sorted: bool,
    include_page_range: bool,
    page_column_width_cm: float,
    table_format: Dict[str, object],
    start_each_table_on_new_page: bool,
    page_label_column_title: str,
    pdf_page_column_title: str,
    include_doc_index_column: bool,
    include_pdf_zero_index: bool,
) -> Tuple[int, List[str]]:
    entries = [
        IndexEntry(doc_index="1", citation="", description="Zulu Holdings v Delta", c_para="", d2_para="", page="45–52" if include_page_range else "45", pdf_page_index="45"),
        IndexEntry(doc_index="2", citation="", description="Alpha Bank v Meridian", c_para="", d2_para="", page="12–19" if include_page_range else "12", pdf_page_index="12"),
        IndexEntry(doc_index="3", citation="", description="Barton v Oceanic Ltd", c_para="", d2_para="", page="73–81" if include_page_range else "73", pdf_page_index="73"),
        IndexEntry(doc_index="4", citation="", description="Crown Plc v Harbour", c_para="", d2_para="", page="101–110" if include_page_range else "101", pdf_page_index="101"),
    ]
    doc = Document()
    doc.add_paragraph("Authorities Index Style Preview")
    doc.add_paragraph("This preview uses your current table and heading settings.")
    warnings: List[str] = []
    warning = _add_index_table(
        doc,
        "Index (by page label)",
        entries,
        style_name=None,
        table_format=table_format,
        page_column_width_cm=page_column_width_cm,
        add_page_break_before=False,
        page_label_column_title=page_label_column_title,
        pdf_page_column_title=pdf_page_column_title,
        include_doc_index_column=include_doc_index_column,
        include_pdf_zero_index=include_pdf_zero_index,
    )
    if warning:
        warnings.append(warning)
    sections_added = 1
    if include_description_sorted:
        by_desc = sorted(entries, key=lambda e: e.description.lower())
        warning = _add_index_table(
            doc,
            "Index (by document description)",
            by_desc,
            style_name=None,
            table_format=table_format,
            page_column_width_cm=page_column_width_cm,
            add_page_break_before=start_each_table_on_new_page and sections_added > 0,
            page_label_column_title=page_label_column_title,
            pdf_page_column_title=pdf_page_column_title,
            include_doc_index_column=include_doc_index_column,
            include_pdf_zero_index=include_pdf_zero_index,
        )
        if warning:
            warnings.append(warning)
        sections_added += 1

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return sections_added, warnings


def write_pagination_report(path: Path, rows: Sequence[PaginationRow]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "Entry",
                "Citation",
                "Case title",
                "Filename",
                "Source page count",
                "Start PDF page position",
                "End PDF page position",
            ],
        )
        writer.writeheader()
        for r in rows:
            writer.writerow(
                {
                    "Entry": r.entry,
                    "Citation": r.citation,
                    "Case title": r.case_title,
                    "Filename": r.filename,
                    "Source page count": r.source_pages,
                    "Start PDF page position": r.start_page,
                    "End PDF page position": r.end_page,
                }
            )


def set_cell_text_preserving_style(cell, text: str) -> None:
    value = normalize_text(text)
    if not cell.paragraphs:
        cell.text = value
        return
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = value
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.add_run(value)
    for pp in cell.paragraphs[1:]:
        for run in pp.runs:
            run.text = ""


def update_index_docx_pages(
    source_docx: Path,
    output_docx: Path,
    pagination: Sequence[PaginationRow],
    *,
    page_label_column_title: str = "Page Label",
    include_page_range: bool = False,
    stamp_enabled: bool = False,
    stamp_number_source: str = "page-label",
    stamp_prefix: str = "",
    stamp_index_page_offset: int = 0,
) -> int:
    page_lookup = {}
    desc_lookup = {}
    for r in pagination:
        page_value = _format_index_page_value(
            r.start_page,
            r.end_page,
            include_page_range,
            stamp_enabled=stamp_enabled,
            stamp_number_source=stamp_number_source,
            stamp_prefix=stamp_prefix,
            index_page_offset=stamp_index_page_offset,
        )
        page_lookup[(normalize_text(r.citation), normalize_text(r.case_title))] = page_value
        desc_lookup[description_from_pdf_path(Path(r.filename))] = page_value

    page_label_header = normalize_header(page_label_column_title)
    doc = Document(str(source_docx))
    updates = 0

    for table in doc.tables:
        if not table.rows:
            continue
        headers = [normalize_header(c.text) for c in table.rows[0].cells]
        col_map: Dict[str, int] = {}
        for idx, h in enumerate(headers):
            if h == "citation" and "Citation" not in col_map:
                col_map["Citation"] = idx
            elif h in {"casetitle", "description"} and "Case title" not in col_map:
                col_map["Case title"] = idx
            elif (
                h == "page"
                or h == "pagelabel"
                or h == page_label_header
                or h.endswith("page")
            ) and "Page" not in col_map:
                col_map["Page"] = idx
        if "Page" not in col_map:
            continue

        for row in table.rows[1:]:
            cells = row.cells
            citation = normalize_text(cells[col_map["Citation"]].text) if "Citation" in col_map else ""
            case_title = normalize_text(cells[col_map["Case title"]].text) if "Case title" in col_map else ""
            if not citation and not case_title:
                continue
            key = (citation, case_title)
            page_value = page_lookup.get(key)
            if page_value is None and case_title:
                page_value = desc_lookup.get(case_title)
            if page_value is None and citation:
                page_value = desc_lookup.get(citation)
            if page_value is None:
                continue
            set_cell_text_preserving_style(cells[col_map["Page"]], page_value)
            updates += 1

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return updates


def export_docx_to_pdf_exact(docx_path: Path, pdf_path: Path) -> None:
    # Uses Word COM via PowerShell to preserve exact formatting.
    command = rf"""
$docx = "{docx_path}"
$pdf = "{pdf_path}"
$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($docx, $false, $true)
  $wdExportFormatPDF = 17
  $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
}} finally {{
  if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
  if ($word -ne $null) {{ $word.Quit() | Out-Null }}
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", command],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "Word PDF export failed.\n"
            f"stdout:\n{result.stdout}\n\nstderr:\n{result.stderr}"
        )


def copy_outline_shifted(reader: PypdfReader, writer: PdfWriter, shift: int, parent=None) -> None:
    try:
        source_outline = reader.outline
    except Exception:
        return
    if not source_outline:
        return

    def walk(items, current_parent=None):
        if not items:
            return
        last = None
        for item in items:
            if isinstance(item, list):
                walk(item, last if last is not None else current_parent)
                continue
            title = getattr(item, "title", None)
            if not title:
                last = None
                continue
            try:
                pnum = reader.get_destination_page_number(item)
            except Exception:
                last = None
                continue
            if pnum is None:
                last = None
                continue
            last = writer.add_outline_item(title, pnum + shift, parent=current_parent)

    walk(source_outline, parent)


def _copy_page_labels(reader: PypdfReader, writer: PdfWriter) -> None:
    try:
        root = reader.trailer.get("/Root")
        if root is None:
            return
        page_labels = root.get("/PageLabels")
        if page_labels is None:
            return
        writer._root_object[NameObject("/PageLabels")] = page_labels.clone(writer)
    except Exception:
        return


def _build_page_label_nav_text(index_row: Optional[IndexRow], fallback_filename: str) -> str:
    if index_row is None:
        return normalize_text(Path(fallback_filename).stem)
    parts = [
        normalize_text(index_row.citation),
        normalize_text(index_row.case_title),
        normalize_text(index_row.c_para),
        normalize_text(index_row.d2_para),
        normalize_text(index_row.page),
    ]
    parts = [p for p in parts if p]
    deduped: List[str] = []
    for part in parts:
        if deduped and deduped[-1].lower() == part.lower():
            continue
        deduped.append(part)
    parts = deduped
    if not parts:
        return normalize_text(Path(fallback_filename).stem)
    return " :: ".join(parts)


def prefix_existing_page_labels_in_place(pdf_path: Path, prefix: str) -> None:
    import pikepdf

    if not prefix:
        return
    pdf = pikepdf.Pdf.open(str(pdf_path))
    try:
        root = pdf.Root
        page_labels = root.get("/PageLabels", None)
        if page_labels is None:
            return
        nums = page_labels.get("/Nums", None)
        if nums is None:
            return
        for i in range(1, len(nums), 2):
            label_dict = nums[i]
            existing_prefix = str(label_dict.get("/P", ""))
            label_dict["/P"] = pikepdf.String(prefix + existing_prefix)
        tmp = pdf_path.with_suffix(".tmp-prefix-pagelabels.pdf")
        if tmp.exists():
            tmp.unlink()
        pdf.save(str(tmp))
    finally:
        pdf.close()
    tmp.replace(pdf_path)


def prefix_page_label_bookmark_titles_in_place(pdf_path: Path, parent_title: str, prefix: str) -> int:
    import pymupdf

    if not prefix:
        return 0
    doc = pymupdf.open(str(pdf_path))
    try:
        toc = doc.get_toc(simple=True) or []
        if not toc:
            return 0
        count = 0
        i = 0
        while i < len(toc):
            level, title, _page = toc[i]
            if level == 1 and normalize_text(str(title)) == normalize_text(parent_title):
                i += 1
                while i < len(toc) and toc[i][0] > 1:
                    item = toc[i]
                    if item[0] == 2 and not str(item[1]).startswith(prefix):
                        item[1] = f"{prefix}{item[1]}"
                        count += 1
                    i += 1
                break
            i += 1
        if count == 0:
            return 0
        doc.set_toc(toc)
        tmp = pdf_path.with_suffix(".tmp-prefix-page-label-bookmarks.pdf")
        if tmp.exists():
            tmp.unlink()
        doc.save(str(tmp), garbage=3, deflate=True)
    finally:
        doc.close()
    tmp.replace(pdf_path)
    return count


def add_document_page_label_bookmarks_in_place(
    pdf_path: Path,
    pagination_rows: Sequence[PaginationRow],
    *,
    index_rows: Optional[Sequence[IndexRow]],
    index_page_count: int,
    parent_title: str,
) -> int:
    import pymupdf

    if not pagination_rows:
        return 0

    doc = pymupdf.open(str(pdf_path))
    try:
        toc = doc.get_toc(simple=True) or []
        page_count = doc.page_count
        parent_page = max(1, min(page_count, index_page_count + 1))
        toc.append([1, parent_title, parent_page])

        count = 0
        for idx, row in enumerate(pagination_rows):
            label = _build_page_label_nav_text(
                index_rows[idx] if index_rows is not None and idx < len(index_rows) else None,
                row.filename,
            )
            dest_page = max(1, min(page_count, index_page_count + row.start_page))
            toc.append([2, label, dest_page])
            count += 1

        doc.set_toc(toc)
        tmp = pdf_path.with_suffix(".tmp-doc-page-labels.pdf")
        if tmp.exists():
            tmp.unlink()
        doc.save(str(tmp), garbage=3, deflate=True)
    finally:
        doc.close()
    tmp.replace(pdf_path)
    return count


def group_top_level_bookmarks_under_parent_in_place(
    pdf_path: Path,
    *,
    parent_title: str,
    preserve_top_titles: Sequence[str],
) -> int:
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    try:
        toc = doc.get_toc(simple=True) or []
        if not toc:
            return 0
        preserve = {normalize_text(t) for t in preserve_top_titles if normalize_text(t)}
        preserve.add(normalize_text(parent_title))
        result: List[List[object]] = []
        grouped = 0
        inserted_parent = False

        i = 0
        while i < len(toc):
            level, title, page = toc[i]
            chunk = [toc[i]]
            i += 1
            while i < len(toc) and toc[i][0] > level:
                chunk.append(toc[i])
                i += 1

            if level == 1 and normalize_text(title) not in preserve:
                if not inserted_parent:
                    parent_page = int(page) if isinstance(page, int) and page > 0 else 1
                    result.append([1, parent_title, parent_page])
                    inserted_parent = True
                for entry in chunk:
                    result.append([entry[0] + 1, entry[1], entry[2]])
                grouped += 1
            else:
                result.extend(chunk)

        if grouped > 0:
            doc.set_toc(result)
            tmp = pdf_path.with_suffix(".tmp-group-bookmarks.pdf")
            if tmp.exists():
                tmp.unlink()
            doc.save(str(tmp), garbage=3, deflate=True)
        else:
            return 0
    finally:
        doc.close()
    tmp.replace(pdf_path)
    return grouped


def prepend_index_pdf(
    index_pdf: Path,
    body_pdf: Path,
    output_pdf: Path,
) -> int:
    index_reader = PypdfReader(str(index_pdf))
    body_reader = PypdfReader(str(body_pdf))
    writer = PdfWriter()

    for p in index_reader.pages:
        writer.add_page(p)
    index_pages = len(index_reader.pages)

    for p in body_reader.pages:
        writer.add_page(p)

    if index_pages > 0:
        writer.set_page_label(
            0,
            index_pages - 1,
            style=PageLabelStyle.LOWERCASE_ROMAN,
            start=1,
        )
    writer.set_page_label(
        index_pages,
        index_pages + len(body_reader.pages) - 1,
        style=PageLabelStyle.DECIMAL,
        start=1,
    )

    writer.add_outline_item("Index", 0)
    copy_outline_shifted(body_reader, writer, index_pages)

    with output_pdf.open("wb") as f:
        writer.write(f)
    return index_pages


def add_full_row_links(
    pdf_path: Path,
    pagination_rows: Sequence[PaginationRow],
    index_page_count: int,
    expected_start_pages: Optional[Sequence[int]] = None,
) -> int:
    import pymupdf

    def parse_token_for_start(token: str) -> Optional[int]:
        token = normalize_text(token).rstrip(".,;:)]")
        m = re.match(r"^(\d+)(?:\s*[-–]\s*\d+)?$", token)
        if not m:
            return None
        value = int(m.group(1))
        if value in valid_starts:
            return value
        return None

    def line_looks_like_page_range(text: str) -> bool:
        t = normalize_text(text)
        if not t:
            return False
        # Accept ranges like "1-8" and letter-prefixed forms like "A1-A8",
        # but reject UUID-like fragments or other long hyphen chains.
        if len(t) > 40:
            return False
        parts = re.split(r"\s*[-–]\s*", t)
        if len(parts) != 2:
            return False
        left = parts[0].strip()
        right = parts[1].strip()
        if not left or not right:
            return False

        def looks_like_label(segment: str) -> bool:
            if len(segment) > 12:
                return False
            if not re.fullmatch(r"[A-Za-z0-9]+", segment):
                return False
            return bool(re.search(r"\d", segment))

        if not looks_like_label(left) or not looks_like_label(right):
            return False
        return True

    def extract_start_page(row_vals: Sequence[object], row_obj, page_obj) -> Optional[int]:
        trailing_vals = list(row_vals[-3:]) if row_vals else []
        for cell in reversed(trailing_vals):
            token = normalize_text(str(cell or ""))
            if not token:
                continue
            parsed = parse_token_for_start(token)
            if parsed is not None:
                return parsed
        cell_boxes = getattr(row_obj, "cells", None)
        if cell_boxes:
            trailing_boxes = list(cell_boxes[-3:])
            for cell_bbox in reversed(trailing_boxes):
                if not cell_bbox:
                    continue
                try:
                    clip = pymupdf.Rect(cell_bbox)
                    text = normalize_text(page_obj.get_text("text", clip=clip))
                    if not text:
                        continue
                    parsed = parse_token_for_start(text)
                    if parsed is not None:
                        return parsed
                    # if text has multiple tokens, check rightmost token first
                    for token in reversed(re.split(r"\s+", text)):
                        parsed = parse_token_for_start(token)
                        if parsed is not None:
                            return parsed
                except Exception:
                    continue
        return None

    valid_starts = {r.start_page for r in pagination_rows}
    doc = pymupdf.open(str(pdf_path))
    links_added = 0

    def rect_key(rect, target_page: int) -> Tuple[int, int, int, int, int]:
        return (
            int(round(rect.x0 * 2)),
            int(round(rect.y0 * 2)),
            int(round(rect.x1 * 2)),
            int(round(rect.y1 * 2)),
            int(target_page),
        )

    def add_link_if_new(page_obj, rect, target_page: int, seen: set) -> bool:
        key = rect_key(rect, target_page)
        if key in seen:
            return False
        page_obj.insert_link(
            {
                "kind": pymupdf.LINK_GOTO,
                "from": rect,
                "page": target_page,
                "to": pymupdf.Point(0, 0),
                "zoom": 0,
            }
        )
        seen.add(key)
        return True

    def estimate_visual_rows_on_page(page_obj) -> int:
        words = page_obj.get_text("words")
        line_groups: Dict[Tuple[int, int], List[tuple]] = {}
        for w in words:
            x0, y0, x1, y1, text, block_no, line_no, *_ = w
            key = (int(block_no), int(line_no))
            line_groups.setdefault(key, []).append((x0, y0, x1, y1, str(text or "")))
        count = 0
        for key in sorted(line_groups.keys()):
            line_words = sorted(line_groups[key], key=lambda t: t[0])
            if not line_words:
                continue
            full_line = " ".join(normalize_text(t[4]) for t in line_words).strip().lower()
            if not full_line or ("description" in full_line and "page" in full_line):
                continue
            if line_looks_like_page_range(full_line):
                count += 1
        return count

    expected_pending = list(expected_start_pages or [])
    expected_pos = 0

    for pno in range(index_page_count):
        page = doc[pno]
        seen_link_keys = set()
        if expected_pending:
            for lnk in list(page.get_links() or []):
                if lnk.get("kind") != pymupdf.LINK_GOTO:
                    continue
                try:
                    page.delete_link(lnk)
                except Exception:
                    continue
            words = page.get_text("words")
            tables = getattr(page.find_tables(), "tables", []) or []
            table_bounds = None
            if tables:
                table_bounds = (
                    min(t.bbox[0] for t in tables),
                    min(t.bbox[1] for t in tables),
                    max(t.bbox[2] for t in tables),
                    max(t.bbox[3] for t in tables),
                )
            line_groups: Dict[Tuple[int, int], List[tuple]] = {}
            for w in words:
                x0, y0, x1, y1, text, block_no, line_no, *_ = w
                key = (int(block_no), int(line_no))
                line_groups.setdefault(key, []).append((x0, y0, x1, y1, str(text or "")))

            lines: List[Dict[str, object]] = []
            for key in sorted(line_groups.keys()):
                line_words = sorted(line_groups[key], key=lambda t: t[0])
                if not line_words:
                    continue
                full_line = " ".join(normalize_text(t[4]) for t in line_words).strip()
                full_line_lower = full_line.lower()
                if not full_line:
                    continue
                line_x0 = min(w[0] for w in line_words)
                line_y0 = min(w[1] for w in line_words)
                line_x1 = max(w[2] for w in line_words)
                line_y1 = max(w[3] for w in line_words)
                if table_bounds is not None:
                    _tx0, ty0, _tx1, ty1 = table_bounds
                    line_mid_y = (line_y0 + line_y1) / 2.0
                    if line_mid_y < (ty0 - 3.0) or line_mid_y > (ty1 + 3.0):
                        continue
                is_header = ("description" in full_line_lower and "page" in full_line_lower)
                has_range = line_looks_like_page_range(full_line)
                lines.append(
                    {
                        "x0": line_x0,
                        "y0": line_y0,
                        "x1": line_x1,
                        "y1": line_y1,
                        "is_header": is_header,
                        "has_range": has_range,
                    }
                )

            row_rects: List[pymupdf.Rect] = []
            for idx, line in enumerate(lines):
                if not line.get("has_range"):
                    continue
                y0 = float(line["y0"])
                y1 = float(line["y1"])
                x0 = float(line["x0"])
                x1 = float(line["x1"])
                if idx > 0:
                    prev = lines[idx - 1]
                    if (
                        not prev.get("is_header")
                        and not prev.get("has_range")
                        and (y0 - float(prev["y1"])) < 24.0
                    ):
                        y0 = min(y0, float(prev["y0"]))
                        x0 = min(x0, float(prev["x0"]))
                        x1 = max(x1, float(prev["x1"]))

                if table_bounds is not None:
                    tx0, _ty0, tx1, _ty1 = table_bounds
                    x0 = float(tx0)
                    x1 = float(tx1)

                row_rects.append(
                    pymupdf.Rect(
                        max(0.0, x0 - 2.0),
                        max(0.0, y0 - 1.0),
                        min(page.rect.width, x1 + 2.0),
                        min(page.rect.height, y1 + 1.0),
                    )
                )

            for rect in sorted(row_rects, key=lambda r: (round(r.y0, 2), round(r.x0, 2))):
                if expected_pos >= len(expected_pending):
                    break
                start_page = expected_pending[expected_pos]
                target = index_page_count + int(start_page) - 1
                if add_link_if_new(page, rect, target, seen_link_keys):
                    links_added += 1
                expected_pos += 1
            continue
        else:
            for lnk in page.get_links() or []:
                if lnk.get("kind") != pymupdf.LINK_GOTO:
                    continue
                from_rect = lnk.get("from")
                target_page = int(lnk.get("page", -1))
                if from_rect is None or target_page < 0:
                    continue
                seen_link_keys.add(rect_key(pymupdf.Rect(from_rect), target_page))

        tf = page.find_tables()
        tables = getattr(tf, "tables", []) if tf is not None else []
        table_row_rects: List[pymupdf.Rect] = []

        for table in tables:
            extracted = table.extract()
            if not extracted:
                continue
            for ridx, row in enumerate(table.rows):
                if ridx == 0:
                    continue
                row_vals = extracted[ridx] if ridx < len(extracted) else []
                row_join = " ".join((v or "") for v in row_vals).lower()
                if "description" in row_join and "page" in row_join:
                    continue
                rb = row.bbox
                table_row_rects.append(
                    pymupdf.Rect(rb[0] + 1.0, rb[1] + 1.0, rb[2] - 1.0, rb[3] - 1.0)
                )

                if not expected_pending:
                    page_num = extract_start_page(row_vals, row, page)
                    if page_num is None:
                        continue

                    target = index_page_count + page_num - 1
                    rect = table_row_rects[-1]
                    if add_link_if_new(page, rect, target, seen_link_keys):
                        links_added += 1

        if expected_pending and table_row_rects:
            ordered_table_rects = sorted(table_row_rects, key=lambda r: (round(r.y0, 2), round(r.x0, 2)))
            remaining = len(expected_pending) - expected_pos
            desired_rows = remaining
            estimated_rows = estimate_visual_rows_on_page(page)
            if estimated_rows > 0:
                desired_rows = min(remaining, estimated_rows)
            if len(ordered_table_rects) > desired_rows > 0:
                gaps: List[Tuple[float, int]] = []
                for i in range(len(ordered_table_rects) - 1):
                    gap = float(ordered_table_rects[i + 1].y0 - ordered_table_rects[i].y1)
                    gaps.append((gap, i))
                split_count = min(desired_rows - 1, len(ordered_table_rects) - 1)
                split_after = sorted(
                    i for _gap, i in sorted(gaps, key=lambda it: it[0], reverse=True)[:split_count]
                )

                grouped: List[List[pymupdf.Rect]] = []
                current_group: List[pymupdf.Rect] = []
                split_set = set(split_after)
                for i, rect in enumerate(ordered_table_rects):
                    current_group.append(rect)
                    if i in split_set:
                        grouped.append(current_group)
                        current_group = []
                if current_group:
                    grouped.append(current_group)

                table_x0 = max(0.0, min(r.x0 for r in ordered_table_rects))
                table_x1 = min(page.rect.width, max(r.x1 for r in ordered_table_rects))
                ordered_table_rects = [
                    pymupdf.Rect(
                        table_x0,
                        max(0.0, min(r.y0 for r in grp)),
                        table_x1,
                        min(page.rect.height, max(r.y1 for r in grp)),
                    )
                    for grp in grouped
                ]

            for rect in ordered_table_rects:
                if expected_pos >= len(expected_pending):
                    break
                start_page = expected_pending[expected_pos]
                target = index_page_count + int(start_page) - 1
                if add_link_if_new(page, rect, target, seen_link_keys):
                    links_added += 1
                expected_pos += 1

        if not expected_pending:
            # Robust fallback for Word-exported PDFs: line-based detection by text blocks.
            words = page.get_text("words")
            line_groups: Dict[Tuple[int, int], List[tuple]] = {}
            for w in words:
                x0, y0, x1, y1, text, block_no, line_no, *_ = w
                key = (int(block_no), int(line_no))
                line_groups.setdefault(key, []).append((x0, y0, x1, y1, str(text or "")))

            seen_rows = set()
            for key in sorted(line_groups.keys()):
                line_words = sorted(line_groups[key], key=lambda t: t[0])
                if not line_words:
                    continue
                full_line = " ".join(normalize_text(t[4]) for t in line_words).strip().lower()
                if not full_line or ("description" in full_line and "page" in full_line):
                    continue

                candidate = None
                candidate_x = None
                for x0, y0, x1, y1, text in reversed(line_words):
                    token = normalize_text(text)
                    if not token:
                        continue
                    cx = (x0 + x1) / 2
                    if cx < (page.rect.width * 0.58):
                        continue
                    parsed = parse_token_for_start(token)
                    if parsed is None or parsed not in valid_starts:
                        continue
                    candidate = parsed
                    candidate_x = cx
                    break
                if candidate is None:
                    continue

                # Avoid accidental links where the "page" token is not clearly in the right-most area.
                if candidate_x is None or candidate_x < (page.rect.width * 0.58):
                    continue

                y_min = min(w[1] for w in line_words)
                y_max = max(w[3] for w in line_words)
                x_min = min(w[0] for w in line_words)
                x_max = max(w[2] for w in line_words)
                row_key = (round(y_min, 1), round(y_max, 1), candidate)
                if row_key in seen_rows:
                    continue
                seen_rows.add(row_key)

                target = index_page_count + candidate - 1
                rect = pymupdf.Rect(
                    max(0, x_min - 2),
                    max(0, y_min - 1),
                    min(page.rect.width, x_max + 2),
                    min(page.rect.height, y_max + 1),
                )
                if add_link_if_new(page, rect, target, seen_link_keys):
                    links_added += 1

    # Last resort: sequential row linking by expected order over line rows
    if expected_pending and expected_pos < len(expected_pending):
        for pno in range(index_page_count):
            if expected_pos >= len(expected_pending):
                break
            page = doc[pno]
            seen_link_keys = set()
            for lnk in page.get_links() or []:
                if lnk.get("kind") != pymupdf.LINK_GOTO:
                    continue
                from_rect = lnk.get("from")
                target_page = int(lnk.get("page", -1))
                if from_rect is None or target_page < 0:
                    continue
                seen_link_keys.add(rect_key(pymupdf.Rect(from_rect), target_page))

            words = page.get_text("words")
            line_groups: Dict[Tuple[int, int], List[tuple]] = {}
            for w in words:
                x0, y0, x1, y1, text, block_no, line_no, *_ = w
                key = (int(block_no), int(line_no))
                line_groups.setdefault(key, []).append((x0, y0, x1, y1, str(text or "")))

            row_rects: List[pymupdf.Rect] = []
            for key in sorted(line_groups.keys()):
                line_words = sorted(line_groups[key], key=lambda t: t[0])
                if not line_words:
                    continue
                full_line = " ".join(normalize_text(t[4]) for t in line_words).strip().lower()
                if not full_line or ("description" in full_line and "page" in full_line):
                    continue
                y_min = min(w[1] for w in line_words)
                y_max = max(w[3] for w in line_words)
                x_min = min(w[0] for w in line_words)
                x_max = max(w[2] for w in line_words)
                row_rects.append(
                    pymupdf.Rect(
                        max(0, x_min - 2),
                        max(0, y_min - 1),
                        min(page.rect.width, x_max + 2),
                        min(page.rect.height, y_max + 1),
                    )
                )

            if row_rects:
                remaining = len(expected_pending) - expected_pos
                ordered = sorted(row_rects, key=lambda r: (round(r.y0, 2), round(r.x0, 2)))
                if len(ordered) > remaining > 0:
                    gaps: List[Tuple[float, int]] = []
                    for i in range(len(ordered) - 1):
                        gap = float(ordered[i + 1].y0 - ordered[i].y1)
                        gaps.append((gap, i))
                    split_count = min(remaining - 1, len(ordered) - 1)
                    split_after = sorted(i for _gap, i in sorted(gaps, key=lambda it: it[0], reverse=True)[:split_count])

                    grouped: List[List[pymupdf.Rect]] = []
                    current_group: List[pymupdf.Rect] = []
                    split_set = set(split_after)
                    for i, rect in enumerate(ordered):
                        current_group.append(rect)
                        if i in split_set:
                            grouped.append(current_group)
                            current_group = []
                    if current_group:
                        grouped.append(current_group)

                    table_x0 = max(0.0, min(r.x0 for r in ordered))
                    table_x1 = min(page.rect.width, max(r.x1 for r in ordered))
                    merged_rects: List[pymupdf.Rect] = []
                    for grp in grouped:
                        merged_rects.append(
                            pymupdf.Rect(
                                table_x0,
                                max(0.0, min(r.y0 for r in grp)),
                                table_x1,
                                min(page.rect.height, max(r.y1 for r in grp)),
                            )
                        )
                    row_rects = merged_rects
                else:
                    row_rects = ordered

            for rect in row_rects:
                if expected_pos >= len(expected_pending):
                    break
                start_page = expected_pending[expected_pos]
                target = index_page_count + int(start_page) - 1
                if add_link_if_new(page, rect, target, seen_link_keys):
                    links_added += 1
                expected_pos += 1

    tmp = pdf_path.with_suffix(".tmp-links.pdf")
    if tmp.exists():
        tmp.unlink()
    doc.save(str(tmp), garbage=3, deflate=True)
    doc.close()
    tmp.replace(pdf_path)
    return links_added


def compress_pdf_in_place(pdf_path: Path) -> None:
    import pikepdf

    tmp = pdf_path.with_suffix(".tmp-compress.pdf")
    if tmp.exists():
        tmp.unlink()
    pdf = pikepdf.Pdf.open(str(pdf_path))
    pdf.save(
        str(tmp),
        compress_streams=True,
        object_stream_mode=pikepdf.ObjectStreamMode.generate,
        recompress_flate=True,
        linearize=True,
    )
    pdf.close()
    tmp.replace(pdf_path)


def set_open_outlines_in_place(pdf_path: Path) -> None:
    import pikepdf

    tmp = pdf_path.with_suffix(".tmp-openmode.pdf")
    if tmp.exists():
        tmp.unlink()
    pdf = pikepdf.Pdf.open(str(pdf_path))
    pdf.Root.PageMode = pikepdf.Name("/UseOutlines")
    pdf.save(str(tmp))
    pdf.close()
    tmp.replace(pdf_path)
    verify_pdf = pikepdf.Pdf.open(str(pdf_path))
    try:
        page_mode = str(getattr(verify_pdf.Root, "PageMode", ""))
        if page_mode != "/UseOutlines":
            raise RuntimeError("Failed to set PDF open mode to /UseOutlines.")
    finally:
        verify_pdf.close()


def stamp_pdf_in_place(
    pdf_path: Path,
    log: Callable[[str], None],
    *,
    prefix: str,
    number_source: str,
    color_preset: str,
    position: str,
    stamp_mode: str,
    bookmark_page_labels: bool,
    page_label_bookmark_parent: str,
    set_page_labels_from_stamp: bool,
) -> None:
    stamper = Path(__file__).with_name("bates_stamper.py")
    if not stamper.exists():
        raise RuntimeError(f"Could not find bates stamper script: {stamper}")

    with tempfile.TemporaryDirectory(prefix="bundle-gui-stamp-") as tmpdir:
        tmp_root = Path(tmpdir)
        in_dir = tmp_root / "in"
        out_dir = tmp_root / "out"
        in_dir.mkdir(parents=True, exist_ok=True)
        out_dir.mkdir(parents=True, exist_ok=True)

        staged_input = in_dir / pdf_path.name
        shutil.copy2(pdf_path, staged_input)

        cmd = [
            sys.executable,
            str(stamper),
            str(in_dir),
            str(out_dir),
            "--prefix",
            prefix,
            "--number-source",
            number_source,
            "--color-preset",
            color_preset,
            "--position",
            position,
            "--stamp-mode",
            stamp_mode,
        ]
        if bookmark_page_labels:
            cmd.append("--bookmark-page-labels")
            cmd.extend(["--page-label-bookmark-parent", page_label_bookmark_parent])
        if set_page_labels_from_stamp and number_source == "sequential":
            cmd.append("--set-page-labels-from-stamp")

        log("Running bates_stamper.py...")
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if result.stdout.strip():
            for line in result.stdout.strip().splitlines():
                log(f"[stamper] {line}")
        if result.returncode != 0:
            raise RuntimeError(
                "Stamping failed.\n"
                f"stdout:\n{result.stdout}\n\nstderr:\n{result.stderr}"
            )

        stamped = out_dir / pdf_path.name
        if not stamped.exists():
            raise RuntimeError("Stamping completed but output file was not produced.")
        shutil.copy2(stamped, pdf_path)


class BundleGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1200x900")

        self.source_folder = tk.StringVar()
        self.index_docx = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_name = tk.StringVar(value="Authorities Bundle Final.pdf")
        self.include_index = tk.BooleanVar(value=True)
        self.link_rows = tk.BooleanVar(value=True)
        self.preserve_source_bookmarks = tk.BooleanVar(value=True)
        self.compress_pdf = tk.BooleanVar(value=True)
        self.open_outlines = tk.BooleanVar(value=True)
        self.generate_index_from_template = tk.BooleanVar(value=True)
        self.use_existing_index_descriptions = tk.BooleanVar(value=False)
        self.include_page_range = tk.BooleanVar(value=False)
        self.page_label_column_title = tk.StringVar(value="Page Label")
        self.pdf_page_column_title = tk.StringVar(value="PDF Page")
        self.include_doc_index_column = tk.BooleanVar(value=False)
        self.include_pdf_zero_index = tk.BooleanVar(value=False)
        self.page_column_width_cm = tk.StringVar(value="2.0")
        self.table_style_family = tk.StringVar(value="Grid")
        self.table_style_name = tk.StringVar(value="")
        self.table_style_header_row = tk.BooleanVar(value=True)
        self.table_style_banded_rows = tk.BooleanVar(value=False)
        self.table_style_first_column = tk.BooleanVar(value=False)
        self.table_style_last_column = tk.BooleanVar(value=False)
        self.table_style_banded_columns = tk.BooleanVar(value=False)
        self.start_each_table_on_new_page = tk.BooleanVar(value=False)
        self.include_sorted_by_description = tk.BooleanVar(value=False)
        self.enable_stamp = tk.BooleanVar(value=True)
        self.stamp_prefix = tk.StringVar(value="")
        self.stamp_number_source = tk.StringVar(value="page-label")
        self.stamp_color_preset = tk.StringVar(value=color_preset_display_from_key("reverse-blue"))
        self.stamp_position = tk.StringVar(value="bottom-right")
        self.stamp_mode = tk.StringVar(value="annotation")
        self.stamp_bookmark_page_labels = tk.BooleanVar(value=True)
        self.stamp_page_label_parent = tk.StringVar(value="NAVIGATE BY PAGE LABEL")
        self.group_document_bookmarks = tk.BooleanVar(value=True)
        self.document_bookmark_parent = tk.StringVar(value="NAVIGATE BY DOCUMENT DESCRIPTION")
        self.stamp_set_page_labels_from_stamp = tk.BooleanVar(value=False)
        self.preview_summary = tk.StringVar(value="Preview not generated yet.")
        self.preview_bundle_rows = tk.StringVar(value="Index (by page label): -")
        self.preview_desc_rows = tk.StringVar(value="Index (by document description): -")
        self.index_option_widgets: List[tk.Widget] = []
        self.index_preview_button: Optional[ttk.Button] = None

        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None

        self._build_ui()
        self.after(150, self._drain_log_queue)

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        tab_help = ttk.Frame(notebook)
        tab_sources_order = ttk.Frame(notebook)
        tab_index_settings = ttk.Frame(notebook)
        tab_stamp = ttk.Frame(notebook)
        tab_output = ttk.Frame(notebook)
        tab_run = ttk.Frame(notebook)
        notebook.add(tab_help, text="0. Help / Instructions")
        notebook.add(tab_sources_order, text="1. Inputs & PDF Order")
        notebook.add(tab_index_settings, text="2. Index Settings")
        notebook.add(tab_stamp, text="3. Page Label Settings")
        notebook.add(tab_output, text="4. Output")
        notebook.add(tab_run, text="5. Run")

        self._build_help_tab(tab_help)
        self._build_sources_order_tab(tab_sources_order)
        self._build_settings_tab(tab_index_settings)
        self._build_stamp_tab(tab_stamp)
        self._build_output_tab(tab_output)
        self._build_run_tab(tab_run)

    def _build_help_tab(self, parent: ttk.Frame) -> None:
        frm = self._make_scrollable_tab_body(parent)
        ttk.Label(frm, text="Recommended workflow", font=("", 11, "bold")).pack(anchor="w", pady=(4, 6))
        steps = [
            "1) Inputs & PDF Order: choose source/index/output and set the PDF order.",
            "2) Index Settings: choose whether to generate index tables or update existing tables only.",
            "3) Page Label Settings: choose stamp prefix/number source and bookmark options.",
            "4) Output: confirm final PDF name and output options.",
            "5) Run: execute pipeline and review log output.",
        ]
        for s in steps:
            ttk.Label(frm, text=s).pack(anchor="w", pady=1)

        ttk.Label(frm, text="Common modes", font=("", 11, "bold")).pack(anchor="w", pady=(12, 6))
        ttk.Label(
            frm,
            text="• Generate new tables: Generate index tables from template = ON, Use descriptions already in index DOCX = OFF",
        ).pack(anchor="w", pady=1)
        ttk.Label(
            frm,
            text="• Update existing tables only: Generate index tables from template = ON, Use descriptions already in index DOCX = ON",
        ).pack(anchor="w", pady=1)

        ttk.Label(frm, text="Placeholder seed table", font=("", 11, "bold")).pack(anchor="w", pady=(12, 6))
        ttk.Label(
            frm,
            text="To keep Word table styles in a template, include a seed table marked [[INDEX_TABLE_PLACEHOLDER]]; it will be removed during generation.",
            wraplength=1020,
            justify="left",
        ).pack(anchor="w", pady=1)

    def _make_scrollable_tab_body(self, parent: ttk.Frame) -> ttk.Frame:
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas, padding=12)
        body.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(canvas_window, width=e.width),
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return body

    def _set_widget_tree_state(self, widget: tk.Widget, state: str) -> None:
        try:
            widget.configure(state=state)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._set_widget_tree_state(child, state)

    def _toggle_index_controls(self, *_args) -> None:
        enabled = bool(self.include_index.get())
        state = "normal" if enabled else "disabled"
        for widget in self.index_option_widgets:
            self._set_widget_tree_state(widget, state)
        if self.index_preview_button is not None:
            self.index_preview_button.configure(state=state)

    def _build_index_table_format_config(self) -> Dict[str, object]:
        return {
            "family": self.table_style_family.get(),
            "style": self.table_style_name.get(),
            "header_row": bool(self.table_style_header_row.get()),
            "total_row": False,
            "banded_rows": bool(self.table_style_banded_rows.get()),
            "first_column": bool(self.table_style_first_column.get()),
            "last_column": bool(self.table_style_last_column.get()),
            "banded_columns": bool(self.table_style_banded_columns.get()),
        }

    def _build_expected_index_start_pages(
        self,
        pdf_list: Sequence[Path],
        index_rows: Sequence[IndexRow],
        pagination_rows: Sequence[PaginationRow],
        use_existing_index_descriptions: bool,
    ) -> List[int]:
        entries: List[Tuple[str, int]] = []
        for idx, (pdf_path, page) in enumerate(zip(pdf_list, pagination_rows)):
            description = description_from_pdf_path(pdf_path)
            if use_existing_index_descriptions and idx < len(index_rows):
                row = index_rows[idx]
                description = normalize_text(row.case_title) or normalize_text(row.citation) or description
            entries.append((description, int(page.start_page)))

        starts: List[int] = [s for _, s in entries]
        if self.include_sorted_by_description.get():
            by_desc = sorted(entries, key=lambda it: (it[0] or "").lower())
            starts.extend(s for _, s in by_desc)
        return starts

    def _refresh_table_style_options(self, *_args) -> None:
        style_doc = Path(self.index_docx.get()) if self.index_docx.get() and Path(self.index_docx.get()).is_file() else None
        styles = discover_table_styles(style_doc)
        filtered = filter_table_styles_by_family(styles, self.table_style_family.get())
        if hasattr(self, "table_style_combo"):
            self.table_style_combo.configure(values=filtered)
        current = normalize_text(self.table_style_name.get())
        if not current or current not in {normalize_text(x) for x in filtered}:
            self.table_style_name.set(filtered[0] if filtered else "")

    def _build_sources_order_tab(self, parent: ttk.Frame) -> None:
        root = self._make_scrollable_tab_body(parent)

        inputs = ttk.LabelFrame(root, text="Inputs", padding=10)
        inputs.pack(fill="x", pady=6)

        def row(label: str, var: tk.StringVar, browse_kind: str) -> None:
            r = ttk.Frame(inputs)
            r.pack(fill="x", pady=6)
            ttk.Label(r, text=label, width=22).pack(side="left")
            ttk.Entry(r, textvariable=var).pack(side="left", fill="x", expand=True, padx=8)
            if browse_kind == "folder":
                ttk.Button(r, text="Browse", command=lambda: self._pick_folder(var)).pack(side="left")
            else:
                ttk.Button(r, text="Browse", command=lambda: self._pick_file(var)).pack(side="left")

        row("Source PDFs folder", self.source_folder, "folder")

        order_box = ttk.LabelFrame(root, text="PDF Order", padding=10)
        order_box.pack(fill="both", expand=True, pady=6)

        top = ttk.Frame(order_box)
        top.pack(fill="x")
        ttk.Button(top, text="Load PDFs from source folder", command=self.load_pdfs).pack(side="left")
        ttk.Button(top, text="Add PDF", command=self.add_pdf).pack(side="left", padx=6)
        ttk.Button(top, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(top, text="Save order JSON", command=self.save_order).pack(side="left", padx=6)
        ttk.Button(top, text="Load order JSON", command=self.load_order).pack(side="left")

        mid = ttk.Frame(order_box)
        mid.pack(fill="both", expand=True, pady=8)

        self.pdf_listbox = tk.Listbox(mid, selectmode=tk.EXTENDED)
        self.pdf_listbox.configure(height=18)
        self.pdf_listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.pdf_listbox.yview)
        sb.pack(side="left", fill="y")
        self.pdf_listbox.config(yscrollcommand=sb.set)

        controls = ttk.Frame(mid)
        controls.pack(side="left", fill="y", padx=8)
        ttk.Button(controls, text="Up", width=16, command=lambda: self.move_selected(-1)).pack(pady=3)
        ttk.Button(controls, text="Down", width=16, command=lambda: self.move_selected(1)).pack(pady=3)
        ttk.Button(controls, text="To Top", width=16, command=self.move_top).pack(pady=3)
        ttk.Button(controls, text="To Bottom", width=16, command=self.move_bottom).pack(pady=3)

        global_defaults_row = ttk.Frame(root)
        global_defaults_row.pack(fill="x", pady=8)
        ttk.Button(
            global_defaults_row,
            text="Load All Preferred Defaults",
            command=self.load_preferred_defaults,
        ).pack(side="left")

    def _build_inputs_tab(self, parent: ttk.Frame) -> None:
        frm = ttk.Frame(parent, padding=12)
        frm.pack(fill="x")

        def row(label: str, var: tk.StringVar, browse_kind: str) -> None:
            r = ttk.Frame(frm)
            r.pack(fill="x", pady=6)
            ttk.Label(r, text=label, width=22).pack(side="left")
            ttk.Entry(r, textvariable=var).pack(side="left", fill="x", expand=True, padx=8)
            if browse_kind == "folder":
                ttk.Button(r, text="Browse", command=lambda: self._pick_folder(var)).pack(side="left")
            else:
                ttk.Button(r, text="Browse", command=lambda: self._pick_file(var)).pack(side="left")

        row("Source PDFs folder", self.source_folder, "folder")

    def _build_order_tab(self, parent: ttk.Frame) -> None:
        root = ttk.Frame(parent, padding=12)
        root.pack(fill="both", expand=True)

        top = ttk.Frame(root)
        top.pack(fill="x")
        ttk.Button(top, text="Load PDFs from source folder", command=self.load_pdfs).pack(side="left")
        ttk.Button(top, text="Add PDF", command=self.add_pdf).pack(side="left", padx=6)
        ttk.Button(top, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(top, text="Save order JSON", command=self.save_order).pack(side="left", padx=6)
        ttk.Button(top, text="Load order JSON", command=self.load_order).pack(side="left")

        mid = ttk.Frame(root)
        mid.pack(fill="both", expand=True, pady=8)

        self.pdf_listbox = tk.Listbox(mid, selectmode=tk.EXTENDED)
        self.pdf_listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.pdf_listbox.yview)
        sb.pack(side="left", fill="y")
        self.pdf_listbox.config(yscrollcommand=sb.set)

        controls = ttk.Frame(mid)
        controls.pack(side="left", fill="y", padx=8)
        ttk.Button(controls, text="Up", width=16, command=lambda: self.move_selected(-1)).pack(pady=3)
        ttk.Button(controls, text="Down", width=16, command=lambda: self.move_selected(1)).pack(pady=3)
        ttk.Button(controls, text="To Top", width=16, command=self.move_top).pack(pady=3)
        ttk.Button(controls, text="To Bottom", width=16, command=self.move_bottom).pack(pady=3)

    def _build_settings_tab(self, parent: ttk.Frame) -> None:
        frm = self._make_scrollable_tab_body(parent)

        docx_row = ttk.Frame(frm)
        docx_row.pack(fill="x", pady=6)
        ttk.Label(docx_row, text="Index/template DOCX", width=24).pack(side="left")
        ttk.Entry(docx_row, textvariable=self.index_docx).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(docx_row, text="Browse", command=self._pick_index_docx).pack(side="left")
        self.index_option_widgets.append(docx_row)

        ttk.Checkbutton(frm, text="Include index and index processing", variable=self.include_index).pack(anchor="w", pady=4)
        self.include_index.trace_add("write", self._toggle_index_controls)
        ttk.Label(
            frm,
            text="When OFF, all index options are disabled and no index DOCX/PDF is created or inserted.",
            foreground="#555",
        ).pack(anchor="w")
        index_controls = ttk.LabelFrame(frm, text="Index controls", padding=8)
        index_controls.pack(fill="x", pady=8)
        self.index_option_widgets.append(index_controls)
        ttk.Checkbutton(index_controls, text="Add full-row links in index tables", variable=self.link_rows).pack(anchor="w", pady=4)
        ttk.Checkbutton(
            index_controls,
            text="Generate index tables from template",
            variable=self.generate_index_from_template,
        ).pack(anchor="w", pady=4)
        ttk.Label(
            index_controls,
            text="Builds new index tables in the DOCX template from the current PDF order and pagination.",
            foreground="#555",
        ).pack(anchor="w")
        ttk.Checkbutton(
            index_controls,
            text="Use descriptions already in index DOCX (otherwise use PDF filenames)",
            variable=self.use_existing_index_descriptions,
        ).pack(anchor="w", pady=4)
        ttk.Label(
            index_controls,
            text="On: keep existing description text from DOCX and fill pages. Off: use PDF filenames as descriptions.",
            foreground="#555",
        ).pack(anchor="w")
        ttk.Checkbutton(
            index_controls,
            text="Show page range (start-end) instead of start page only",
            variable=self.include_page_range,
        ).pack(anchor="w", pady=4)
        ttk.Label(
            index_controls,
            text="Example: 245-263 instead of only 245 in the Page column.",
            foreground="#555",
        ).pack(anchor="w")
        ttk.Checkbutton(
            index_controls,
            text="Include # column (document index)",
            variable=self.include_doc_index_column,
        ).pack(anchor="w", pady=4)
        col_labels = ttk.Frame(index_controls)
        col_labels.pack(fill="x", pady=4)
        ttk.Label(col_labels, text="Page-label column title", width=36).grid(row=0, column=0, sticky="w")
        ttk.Entry(col_labels, textvariable=self.page_label_column_title, width=24).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(col_labels, text="PDF-page column title", width=36).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(col_labels, textvariable=self.pdf_page_column_title, width=24).grid(row=1, column=1, sticky="w", padx=8, pady=(4, 0))
        ttk.Checkbutton(
            index_controls,
            text="Include PDF Page column",
            variable=self.include_pdf_zero_index,
        ).pack(anchor="w", pady=4)
        ttk.Checkbutton(
            index_controls,
            text="Start each additional index table on a new page",
            variable=self.start_each_table_on_new_page,
        ).pack(anchor="w", pady=4)
        ttk.Label(
            index_controls,
            text="Index table titles use Word's built-in Caption style.",
            foreground="#555",
        ).pack(anchor="w")
        table_style_row = ttk.Frame(index_controls)
        table_style_row.pack(fill="x", pady=4)
        ttk.Label(table_style_row, text="Table style family", width=36).pack(side="left")
        family_combo = ttk.Combobox(
            table_style_row,
            textvariable=self.table_style_family,
            values=["Plain", "Grid"],
            state="readonly",
            width=18,
        )
        family_combo.pack(side="left", padx=8)
        family_combo.bind("<<ComboboxSelected>>", self._refresh_table_style_options)

        table_style_row2 = ttk.Frame(index_controls)
        table_style_row2.pack(fill="x", pady=4)
        ttk.Label(table_style_row2, text="Table style", width=36).pack(side="left")
        self.table_style_combo = ttk.Combobox(
            table_style_row2,
            textvariable=self.table_style_name,
            values=[],
            state="readonly",
            width=44,
        )
        self.table_style_combo.pack(side="left", padx=8)
        self._refresh_table_style_options()
        ttk.Label(
            index_controls,
            text="Styles are loaded from available Word table styles in the template/document and filtered by family.",
            foreground="#555",
        ).pack(anchor="w")

        style_flags = ttk.Frame(index_controls)
        style_flags.pack(fill="x", pady=4)
        ttk.Checkbutton(style_flags, text="Header Row", variable=self.table_style_header_row).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Checkbutton(style_flags, text="Banded Rows", variable=self.table_style_banded_rows).grid(row=0, column=1, sticky="w", padx=(0, 12))
        ttk.Checkbutton(style_flags, text="First Column", variable=self.table_style_first_column).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=(4, 0))
        ttk.Checkbutton(style_flags, text="Last Column", variable=self.table_style_last_column).grid(row=1, column=1, sticky="w", padx=(0, 12), pady=(4, 0))
        ttk.Checkbutton(style_flags, text="Banded Columns", variable=self.table_style_banded_columns).grid(row=1, column=2, sticky="w", pady=(4, 0))
        ttk.Label(
            index_controls,
            text="Table style options are applied natively via Word table style settings.",
            foreground="#555",
        ).pack(anchor="w")
        ttk.Label(
            index_controls,
            text="If selected style is unavailable, fallback style is Table Grid.",
            foreground="#555",
        ).pack(anchor="w")
        width_row = ttk.Frame(index_controls)
        width_row.pack(fill="x", pady=4)
        ttk.Label(width_row, text="Page / PDF index column width (cm)", width=36).pack(side="left")
        ttk.Entry(width_row, textvariable=self.page_column_width_cm, width=10).pack(side="left")

        extra_box = ttk.LabelFrame(index_controls, text="Index tables")
        extra_box.pack(fill="x", pady=6)
        ttk.Label(
            extra_box,
            text="Index (by page label) is always included.",
            foreground="#555",
        ).pack(anchor="w", padx=8, pady=2)
        ttk.Checkbutton(
            extra_box,
            text="Include Index (by document description)",
            variable=self.include_sorted_by_description,
        ).pack(anchor="w", padx=8, pady=2)

        preview_box = ttk.LabelFrame(index_controls, text="Index output preview", padding=8)
        preview_box.pack(fill="x", pady=8)
        ttk.Button(preview_box, text="Refresh preview", command=self.refresh_index_preview).pack(anchor="w")
        self.index_preview_button = ttk.Button(
            preview_box,
            text="Generate style preview DOCX",
            command=self.generate_style_preview_docx,
        )
        self.index_preview_button.pack(anchor="w", pady=(6, 0))
        ttk.Label(preview_box, textvariable=self.preview_summary).pack(anchor="w", pady=(6, 2))
        ttk.Label(preview_box, textvariable=self.preview_bundle_rows).pack(anchor="w", pady=1)
        ttk.Label(preview_box, textvariable=self.preview_desc_rows).pack(anchor="w", pady=1)

        profile_row = ttk.Frame(frm)
        profile_row.pack(fill="x", pady=8)
        ttk.Button(profile_row, text="Load Preferred Build Defaults", command=self.load_preferred_build_defaults).pack(side="left")
        ttk.Button(profile_row, text="Save profile JSON", command=self.save_profile).pack(side="left")
        ttk.Button(profile_row, text="Load profile JSON", command=self.load_profile).pack(side="left", padx=8)

        ttk.Label(
            frm,
            text=(
                "Tip: Save a profile to reuse folder paths, PDF ordering and all build/stamp settings."
            ),
            foreground="#444",
        ).pack(anchor="w", pady=10)
        self._toggle_index_controls()

    def _build_stamp_tab(self, parent: ttk.Frame) -> None:
        frm = self._make_scrollable_tab_body(parent)

        stamp_box = ttk.LabelFrame(frm, text="One-click Build + Stamp", padding=10)
        stamp_box.pack(fill="x", pady=10)
        ttk.Checkbutton(stamp_box, text="Enable stamping in this pipeline run", variable=self.enable_stamp).pack(anchor="w", pady=2)
        ttk.Label(
            stamp_box,
            text="Mode guide: annotation = overlay stamp layer (safer with complex PDFs); merge = merge stamp page content.",
            foreground="#555",
        ).pack(anchor="w", pady=(0, 6))
        ttk.Label(
            stamp_box,
            text="Page-label folder uses actual PDF page labels. Document-description grouping is controlled separately below.",
            foreground="#555",
        ).pack(anchor="w", pady=(0, 6))

        row1 = ttk.Frame(stamp_box)
        row1.pack(fill="x", pady=4)
        ttk.Label(row1, text="Prefix", width=18).pack(side="left")
        ttk.Entry(row1, textvariable=self.stamp_prefix, width=16).pack(side="left", padx=(0, 12))
        ttk.Label(row1, text="Number source", width=14).pack(side="left")
        ttk.Combobox(
            row1,
            textvariable=self.stamp_number_source,
            values=["page-label", "sequential"],
            width=14,
            state="readonly",
        ).pack(side="left")

        row2 = ttk.Frame(stamp_box)
        row2.pack(fill="x", pady=4)
        ttk.Label(row2, text="Color preset", width=18).pack(side="left")
        ttk.Combobox(
            row2,
            textvariable=self.stamp_color_preset,
            values=[display for _key, display in STAMP_COLOR_PRESET_OPTIONS],
            width=44,
            state="readonly",
        ).pack(side="left", padx=(0, 12))
        ttk.Label(row2, text="Position", width=14).pack(side="left")
        ttk.Combobox(
            row2,
            textvariable=self.stamp_position,
            values=["bottom-right", "bottom-left", "bottom-middle", "top-right", "top-left", "top-middle", "center"],
            width=14,
            state="readonly",
        ).pack(side="left")

        row3 = ttk.Frame(stamp_box)
        row3.pack(fill="x", pady=4)
        ttk.Label(row3, text="Stamp mode", width=18).pack(side="left")
        ttk.Combobox(
            row3,
            textvariable=self.stamp_mode,
            values=["annotation", "merge"],
            width=16,
            state="readonly",
        ).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(
            row3,
            text="Add document navigation bookmark folder",
            variable=self.stamp_bookmark_page_labels,
        ).pack(side="left")

        row4 = ttk.Frame(stamp_box)
        row4.pack(fill="x", pady=4)
        ttk.Label(row4, text="Page-label bookmark parent", width=24).pack(side="left")
        ttk.Entry(row4, textvariable=self.stamp_page_label_parent).pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Checkbutton(
            row4,
            text="Set page labels from sequential stamp",
            variable=self.stamp_set_page_labels_from_stamp,
        ).pack(side="left")

        row5 = ttk.Frame(stamp_box)
        row5.pack(fill="x", pady=4)
        ttk.Checkbutton(
            row5,
            text="Group existing document bookmarks under parent folder",
            variable=self.group_document_bookmarks,
        ).pack(side="left")

        row6 = ttk.Frame(stamp_box)
        row6.pack(fill="x", pady=4)
        ttk.Label(row6, text="Document bookmark parent", width=24).pack(side="left")
        ttk.Entry(row6, textvariable=self.document_bookmark_parent).pack(side="left", fill="x", expand=True, padx=(0, 10))

        defaults_row = ttk.Frame(frm)
        defaults_row.pack(fill="x", pady=8)
        ttk.Button(defaults_row, text="Load Page Label Defaults", command=self.load_preferred_stamp_defaults).pack(side="left")

    def _build_output_tab(self, parent: ttk.Frame) -> None:
        frm = self._make_scrollable_tab_body(parent)

        output_box = ttk.LabelFrame(frm, text="Output", padding=10)
        output_box.pack(fill="x", pady=6)

        out_row = ttk.Frame(output_box)
        out_row.pack(fill="x", pady=6)
        ttk.Label(out_row, text="Output folder", width=22).pack(side="left")
        ttk.Entry(out_row, textvariable=self.output_folder).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out_row, text="Browse", command=lambda: self._pick_folder(self.output_folder)).pack(side="left")

        name_row = ttk.Frame(output_box)
        name_row.pack(fill="x", pady=6)
        ttk.Label(name_row, text="Final output filename", width=22).pack(side="left")
        ttk.Entry(name_row, textvariable=self.output_name).pack(side="left", fill="x", expand=True, padx=8)

        ttk.Checkbutton(output_box, text="Preserve source document bookmarks during merge", variable=self.preserve_source_bookmarks).pack(anchor="w", pady=4)
        ttk.Checkbutton(output_box, text="Compress final PDF", variable=self.compress_pdf).pack(anchor="w", pady=4)
        ttk.Checkbutton(output_box, text="Open with bookmarks panel (/UseOutlines)", variable=self.open_outlines).pack(anchor="w", pady=4)
        ttk.Label(
            output_box,
            text="Sets PDF open mode to /UseOutlines so the bookmarks panel opens by default.",
            foreground="#555",
        ).pack(anchor="w")

    def _build_run_tab(self, parent: ttk.Frame) -> None:
        frm = ttk.Frame(parent, padding=12)
        frm.pack(fill="both", expand=True)

        top = ttk.Frame(frm)
        top.pack(fill="x")
        self.run_btn = ttk.Button(top, text="Run Pipeline", command=self.run_pipeline)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(top, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        self.log_text = tk.Text(frm, wrap="word", height=30)
        self.log_text.pack(fill="both", expand=True, pady=10)

    def _pick_folder(self, var: tk.StringVar) -> None:
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def _pick_file(self, var: tk.StringVar) -> None:
        p = filedialog.askopenfilename(filetypes=[("DOCX/PDF", "*.docx *.pdf"), ("All files", "*.*")])
        if p:
            var.set(p)

    def _pick_index_docx(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Word document", "*.docx"), ("All files", "*.*")])
        if p:
            self.index_docx.set(p)
            self._refresh_table_style_options()

    def _listbox_items(self) -> List[str]:
        return [self.pdf_listbox.get(i) for i in range(self.pdf_listbox.size())]

    def load_pdfs(self) -> None:
        source = Path(self.source_folder.get())
        if not source.is_dir():
            messagebox.showerror("Invalid folder", "Select a valid source PDFs folder first.")
            return
        pdfs = sorted([p for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"])
        self.pdf_listbox.delete(0, tk.END)
        for p in pdfs:
            self.pdf_listbox.insert(tk.END, str(p))

    def add_pdf(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for f in files:
            self.pdf_listbox.insert(tk.END, f)

    def remove_selected(self) -> None:
        selected = list(self.pdf_listbox.curselection())
        for i in reversed(selected):
            self.pdf_listbox.delete(i)

    def move_selected(self, delta: int) -> None:
        selected = list(self.pdf_listbox.curselection())
        if not selected:
            return
        items = self._listbox_items()
        for idx in (selected if delta < 0 else reversed(selected)):
            new_idx = idx + delta
            if new_idx < 0 or new_idx >= len(items):
                continue
            items[idx], items[new_idx] = items[new_idx], items[idx]
        self.pdf_listbox.delete(0, tk.END)
        for item in items:
            self.pdf_listbox.insert(tk.END, item)
        for idx in selected:
            ni = idx + delta
            if 0 <= ni < len(items):
                self.pdf_listbox.select_set(ni)

    def move_top(self) -> None:
        selected = list(self.pdf_listbox.curselection())
        if not selected:
            return
        items = self._listbox_items()
        picked = [items[i] for i in selected]
        rest = [items[i] for i in range(len(items)) if i not in selected]
        merged = picked + rest
        self.pdf_listbox.delete(0, tk.END)
        for item in merged:
            self.pdf_listbox.insert(tk.END, item)
        for i in range(len(picked)):
            self.pdf_listbox.select_set(i)

    def move_bottom(self) -> None:
        selected = list(self.pdf_listbox.curselection())
        if not selected:
            return
        items = self._listbox_items()
        picked = [items[i] for i in selected]
        rest = [items[i] for i in range(len(items)) if i not in selected]
        merged = rest + picked
        self.pdf_listbox.delete(0, tk.END)
        for item in merged:
            self.pdf_listbox.insert(tk.END, item)
        start = len(rest)
        for i in range(start, len(merged)):
            self.pdf_listbox.select_set(i)

    def save_order(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        data = {"pdfs": self._listbox_items()}
        Path(path).write_text(json.dumps(data, indent=2), encoding="utf-8")

    def load_order(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        pdfs = data.get("pdfs", [])
        self.pdf_listbox.delete(0, tk.END)
        for p in pdfs:
            self.pdf_listbox.insert(tk.END, p)

    def _collect_profile_data(self) -> Dict[str, object]:
        return {
            "source_folder": self.source_folder.get(),
            "index_docx": self.index_docx.get(),
            "output_folder": self.output_folder.get(),
            "output_name": self.output_name.get(),
            "include_index": bool(self.include_index.get()),
            "link_rows": bool(self.link_rows.get()),
            "preserve_source_bookmarks": bool(self.preserve_source_bookmarks.get()),
            "compress_pdf": bool(self.compress_pdf.get()),
            "open_outlines": bool(self.open_outlines.get()),
            "generate_index_from_template": bool(self.generate_index_from_template.get()),
            "use_existing_index_descriptions": bool(self.use_existing_index_descriptions.get()),
            "include_page_range": bool(self.include_page_range.get()),
            "page_label_column_title": self.page_label_column_title.get(),
            "pdf_page_column_title": self.pdf_page_column_title.get(),
            "include_doc_index_column": bool(self.include_doc_index_column.get()),
            "include_pdf_zero_index": bool(self.include_pdf_zero_index.get()),
            "page_column_width_cm": self.page_column_width_cm.get(),
            "table_style_family": self.table_style_family.get(),
            "table_style_name": self.table_style_name.get(),
            "table_style_header_row": bool(self.table_style_header_row.get()),
            "table_style_banded_rows": bool(self.table_style_banded_rows.get()),
            "table_style_first_column": bool(self.table_style_first_column.get()),
            "table_style_last_column": bool(self.table_style_last_column.get()),
            "table_style_banded_columns": bool(self.table_style_banded_columns.get()),
            "start_each_table_on_new_page": bool(self.start_each_table_on_new_page.get()),
            "include_sorted_by_description": bool(self.include_sorted_by_description.get()),
            "enable_stamp": bool(self.enable_stamp.get()),
            "stamp_prefix": self.stamp_prefix.get(),
            "stamp_number_source": self.stamp_number_source.get(),
            "stamp_color_preset": color_preset_key_from_display(self.stamp_color_preset.get()),
            "stamp_position": self.stamp_position.get(),
            "stamp_mode": self.stamp_mode.get(),
            "stamp_bookmark_page_labels": bool(self.stamp_bookmark_page_labels.get()),
            "stamp_page_label_parent": self.stamp_page_label_parent.get(),
            "group_document_bookmarks": bool(self.group_document_bookmarks.get()),
            "document_bookmark_parent": self.document_bookmark_parent.get(),
            "stamp_set_page_labels_from_stamp": bool(self.stamp_set_page_labels_from_stamp.get()),
            "pdfs": self._listbox_items(),
        }

    def refresh_index_preview(self) -> None:
        pdf_count = len(self._listbox_items())
        rows = 0
        source_label = "PDF filenames"
        try:
            index_path = Path(self.index_docx.get())
            if self.use_existing_index_descriptions.get() and index_path.is_file():
                rows = len(
                    read_index_rows(
                        index_path,
                        page_label_column_title=self.page_label_column_title.get(),
                        pdf_page_column_title=self.pdf_page_column_title.get(),
                    )
                )
                source_label = "index DOCX descriptions"
            elif self.generate_index_from_template.get():
                rows = pdf_count
                source_label = "PDF filenames"
            else:
                self.preview_summary.set("Preview unavailable: provide a valid index DOCX.")
                self.preview_bundle_rows.set("Index (by page label): -")
                self.preview_desc_rows.set("Index (by document description): -")
                return
        except Exception:
            if self.generate_index_from_template.get() and not self.use_existing_index_descriptions.get():
                rows = pdf_count
                source_label = "PDF filenames"
            else:
                self.preview_summary.set("Preview unavailable: could not parse index DOCX.")
                self.preview_bundle_rows.set("Index (by page label): -")
                self.preview_desc_rows.set("Index (by document description): -")
                return

        self.preview_summary.set(f"Detected {rows} entries from {source_label}.")
        self.preview_bundle_rows.set(f"Index (by page label): {rows} rows")
        self.preview_desc_rows.set(
            f"Index (by document description): {rows if self.include_sorted_by_description.get() else 0} rows"
        )

    def generate_style_preview_docx(self) -> None:
        try:
            try:
                width_cm = float(self.page_column_width_cm.get() or "2.0")
            except ValueError as exc:
                raise RuntimeError("Page column width must be a number (cm).") from exc

            output_base = Path(self.output_folder.get()) if self.output_folder.get() else Path.cwd()
            if not output_base.exists():
                output_base.mkdir(parents=True, exist_ok=True)
            preview_docx = output_base / "Authorities Index Style Preview.docx"
            sections, warnings = generate_index_style_preview_docx(
                preview_docx,
                include_description_sorted=bool(self.include_sorted_by_description.get()),
                include_page_range=bool(self.include_page_range.get()),
                include_doc_index_column=bool(self.include_doc_index_column.get()),
                include_pdf_zero_index=bool(self.include_pdf_zero_index.get()),
                page_column_width_cm=width_cm,
                table_format=self._build_index_table_format_config(),
                start_each_table_on_new_page=bool(self.start_each_table_on_new_page.get()),
                page_label_column_title=self.page_label_column_title.get(),
                pdf_page_column_title=self.pdf_page_column_title.get(),
            )
            self._log(f"Style preview DOCX generated: {preview_docx} (sections: {sections})")
            for warning in warnings:
                self._log(f"Warning: {warning}")
            try:
                os.startfile(str(preview_docx))
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror("Style preview failed", str(exc))

    def _preferred_build_defaults_data(self) -> Dict[str, object]:
        return {
            "include_index": True,
            "link_rows": True,
            "preserve_source_bookmarks": True,
            "compress_pdf": True,
            "open_outlines": True,
            "generate_index_from_template": True,
            "use_existing_index_descriptions": False,
            "include_page_range": False,
            "page_label_column_title": "Page Label",
            "pdf_page_column_title": "PDF Page",
            "include_doc_index_column": False,
            "include_pdf_zero_index": False,
            "page_column_width_cm": "2.0",
            "table_style_family": "Grid",
            "table_style_name": "Grid Table 4 - Accent 1",
            "table_style_header_row": True,
            "table_style_banded_rows": False,
            "table_style_first_column": False,
            "table_style_last_column": False,
            "table_style_banded_columns": False,
            "start_each_table_on_new_page": False,
            "include_sorted_by_description": False,
        }

    def _preferred_stamp_defaults_data(self) -> Dict[str, object]:
        return {
            "enable_stamp": True,
            "stamp_prefix": "",
            "stamp_number_source": "page-label",
            "stamp_color_preset": "reverse-blue",
            "stamp_position": "bottom-right",
            "stamp_mode": "annotation",
            "stamp_bookmark_page_labels": True,
            "stamp_page_label_parent": "NAVIGATE BY PAGE LABEL",
            "group_document_bookmarks": True,
            "document_bookmark_parent": "NAVIGATE BY DOCUMENT DESCRIPTION",
            "stamp_set_page_labels_from_stamp": False,
        }

    def _preferred_defaults_data(self) -> Dict[str, object]:
        data = self._preferred_build_defaults_data()
        data.update(self._preferred_stamp_defaults_data())
        return data

    def _apply_profile_data(self, data: Dict[str, object]) -> None:
        self.source_folder.set(str(data.get("source_folder", self.source_folder.get())))
        self.index_docx.set(str(data.get("index_docx", self.index_docx.get())))
        self.output_folder.set(str(data.get("output_folder", self.output_folder.get())))
        self.output_name.set(str(data.get("output_name", self.output_name.get())))
        self.include_index.set(bool(data.get("include_index", self.include_index.get())))
        self.link_rows.set(bool(data.get("link_rows", self.link_rows.get())))
        self.preserve_source_bookmarks.set(bool(data.get("preserve_source_bookmarks", self.preserve_source_bookmarks.get())))
        self.compress_pdf.set(bool(data.get("compress_pdf", self.compress_pdf.get())))
        self.open_outlines.set(bool(data.get("open_outlines", self.open_outlines.get())))
        self.generate_index_from_template.set(bool(data.get("generate_index_from_template", self.generate_index_from_template.get())))
        self.use_existing_index_descriptions.set(
            bool(data.get("use_existing_index_descriptions", self.use_existing_index_descriptions.get()))
        )
        self.include_page_range.set(bool(data.get("include_page_range", self.include_page_range.get())))
        self.page_label_column_title.set(str(data.get("page_label_column_title", self.page_label_column_title.get())))
        self.pdf_page_column_title.set(str(data.get("pdf_page_column_title", self.pdf_page_column_title.get())))
        self.include_doc_index_column.set(bool(data.get("include_doc_index_column", self.include_doc_index_column.get())))
        self.include_pdf_zero_index.set(bool(data.get("include_pdf_zero_index", self.include_pdf_zero_index.get())))
        self.page_column_width_cm.set(str(data.get("page_column_width_cm", self.page_column_width_cm.get())))
        self.table_style_family.set(str(data.get("table_style_family", self.table_style_family.get())))
        self.table_style_name.set(str(data.get("table_style_name", self.table_style_name.get())))
        self.table_style_header_row.set(bool(data.get("table_style_header_row", self.table_style_header_row.get())))
        self.table_style_banded_rows.set(bool(data.get("table_style_banded_rows", self.table_style_banded_rows.get())))
        self.table_style_first_column.set(bool(data.get("table_style_first_column", self.table_style_first_column.get())))
        self.table_style_last_column.set(bool(data.get("table_style_last_column", self.table_style_last_column.get())))
        self.table_style_banded_columns.set(bool(data.get("table_style_banded_columns", self.table_style_banded_columns.get())))
        self.start_each_table_on_new_page.set(
            bool(data.get("start_each_table_on_new_page", self.start_each_table_on_new_page.get()))
        )
        self._refresh_table_style_options()
        self.include_sorted_by_description.set(
            bool(data.get("include_sorted_by_description", self.include_sorted_by_description.get()))
        )
        self.enable_stamp.set(bool(data.get("enable_stamp", self.enable_stamp.get())))
        self.stamp_prefix.set(str(data.get("stamp_prefix", self.stamp_prefix.get())))
        self.stamp_number_source.set(str(data.get("stamp_number_source", self.stamp_number_source.get())))
        self.stamp_color_preset.set(color_preset_display_from_key(str(data.get("stamp_color_preset", self.stamp_color_preset.get()))))
        self.stamp_position.set(str(data.get("stamp_position", self.stamp_position.get())))
        self.stamp_mode.set(str(data.get("stamp_mode", self.stamp_mode.get())))
        self.stamp_bookmark_page_labels.set(bool(data.get("stamp_bookmark_page_labels", self.stamp_bookmark_page_labels.get())))
        self.stamp_page_label_parent.set(str(data.get("stamp_page_label_parent", self.stamp_page_label_parent.get())))
        self.group_document_bookmarks.set(bool(data.get("group_document_bookmarks", self.group_document_bookmarks.get())))
        self.document_bookmark_parent.set(str(data.get("document_bookmark_parent", self.document_bookmark_parent.get())))
        self.stamp_set_page_labels_from_stamp.set(
            bool(data.get("stamp_set_page_labels_from_stamp", self.stamp_set_page_labels_from_stamp.get()))
        )
        pdfs = data.get("pdfs")
        if isinstance(pdfs, list):
            self.pdf_listbox.delete(0, tk.END)
            for p in pdfs:
                self.pdf_listbox.insert(tk.END, str(p))

    def load_preferred_defaults(self) -> None:
        self._apply_profile_data(self._preferred_defaults_data())
        self._log("Applied all preferred defaults (build + stamp).")
        messagebox.showinfo("Defaults applied", "Applied all preferred defaults.")

    def load_preferred_build_defaults(self) -> None:
        self._apply_profile_data(self._preferred_build_defaults_data())
        self._log("Applied preferred build defaults.")
        messagebox.showinfo("Defaults applied", "Applied preferred build defaults.")

    def load_preferred_stamp_defaults(self) -> None:
        self._apply_profile_data(self._preferred_stamp_defaults_data())
        self._log("Applied preferred stamp defaults.")
        messagebox.showinfo("Defaults applied", "Applied page label defaults.")

    def save_profile(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            title="Save bundle profile",
        )
        if not path:
            return
        data = self._collect_profile_data()
        Path(path).write_text(json.dumps(data, indent=2), encoding="utf-8")
        self._log(f"Profile saved: {path}")

    def load_profile(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")],
            title="Load bundle profile",
        )
        if not path:
            return
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise RuntimeError("Invalid profile format.")
        self._apply_profile_data(data)
        self._log(f"Profile loaded: {path}")

    def _log(self, message: str) -> None:
        self.log_queue.put(message)

    def _drain_log_queue(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.after(150, self._drain_log_queue)

    def run_pipeline(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.run_btn.config(state="disabled")
        self.progress.start(10)
        self.worker = threading.Thread(target=self._run_pipeline_worker, daemon=True)
        self.worker.start()

    def _run_pipeline_worker(self) -> None:
        try:
            source_folder = Path(self.source_folder.get())
            index_docx = Path(self.index_docx.get())
            output_folder = Path(self.output_folder.get())
            pdf_list = [Path(p) for p in self._listbox_items()]
            include_index_enabled = bool(self.include_index.get())
            generate_index_enabled = include_index_enabled and bool(self.generate_index_from_template.get())
            use_existing_desc_enabled = include_index_enabled and bool(self.use_existing_index_descriptions.get())
            include_pdf_zero_index_enabled = include_index_enabled and bool(self.include_pdf_zero_index.get())
            link_rows_enabled = include_index_enabled and bool(self.link_rows.get())
            stamp_enabled = bool(self.enable_stamp.get())
            stamp_number_source = self.stamp_number_source.get()
            stamp_prefix = self.stamp_prefix.get()

            has_index_docx = index_docx.is_file()
            require_index_docx = include_index_enabled

            if not source_folder.is_dir():
                raise RuntimeError("Invalid source folder.")
            if require_index_docx and not has_index_docx:
                raise RuntimeError(
                    "Invalid index DOCX path. This run needs a DOCX because "
                    "index insertion or index-table generation is enabled."
                )
            if not output_folder.exists():
                output_folder.mkdir(parents=True, exist_ok=True)
            if not pdf_list:
                raise RuntimeError("No PDFs in the ordered list.")
            try:
                page_col_width_cm = float(self.page_column_width_cm.get() or "2.0")
            except ValueError as exc:
                raise RuntimeError("Page column width must be a number (cm).") from exc

            index_rows = make_rows_from_pdf_names(pdf_list)
            index_rows_from_docx = False
            if has_index_docx:
                self._log("Reading index metadata rows from DOCX...")
                try:
                    parsed_rows = read_index_rows(
                        index_docx,
                        page_label_column_title=self.page_label_column_title.get(),
                        pdf_page_column_title=self.pdf_page_column_title.get(),
                    )
                    self._log(f"Index metadata rows detected: {len(parsed_rows)}")
                    if len(parsed_rows) == len(pdf_list):
                        index_rows = parsed_rows
                        index_rows_from_docx = True
                    else:
                        self._log(
                            f"Warning: Index rows ({len(parsed_rows)}) do not match PDF count ({len(pdf_list)}). "
                            "Using PDF filename-based metadata for merge/bookmarks."
                        )
                except Exception as exc:
                    if generate_index_enabled and not use_existing_desc_enabled:
                        self._log(
                            f"Warning: Could not read index metadata rows ({exc}). "
                            "Falling back to PDF filename-based metadata."
                        )
                    else:
                        raise
            else:
                self._log("No index DOCX supplied; using PDF filename-based metadata.")

            merged_pdf = output_folder / "_bundle_merged_no_index.pdf"
            self._log("Merging ordered PDFs and creating document bookmarks...")
            pagination = merge_bundle_from_ordered_list(
                pdf_list,
                index_rows,
                merged_pdf,
                self._log,
                preserve_source_bookmarks=bool(self.preserve_source_bookmarks.get()),
            )
            self._log(f"Merged bundle pages: {sum(r.source_pages for r in pagination)}")

            report_csv = output_folder / "authorities_bundle_pagination_report.csv"
            write_pagination_report(report_csv, pagination)
            self._log(f"Pagination report: {report_csv}")

            updated_docx: Optional[Path] = None
            if include_index_enabled and has_index_docx:
                updated_docx = output_folder / f"{index_docx.stem} - populated pages.docx"
                if generate_index_enabled:
                    if use_existing_desc_enabled:
                        updated_rows = update_index_docx_pages(
                            index_docx,
                            updated_docx,
                            pagination,
                            page_label_column_title=self.page_label_column_title.get(),
                            include_page_range=bool(self.include_page_range.get()),
                            stamp_enabled=False,
                            stamp_number_source=stamp_number_source,
                            stamp_prefix=stamp_prefix,
                            stamp_index_page_offset=0,
                        )
                        self._log(
                            f"Updated existing index tables in DOCX: {updated_docx} "
                            f"(rows updated: {updated_rows})"
                        )
                    else:
                        self._log("Generating index tables from template DOCX...")
                        row_count, section_count, style_warnings = generate_index_docx_from_template(
                            index_docx,
                            updated_docx,
                            pdf_list,
                            index_rows,
                            pagination,
                            include_description_sorted=bool(self.include_sorted_by_description.get()),
                            use_existing_index_descriptions=False,
                            include_page_range=bool(self.include_page_range.get()),
                            page_column_width_cm=page_col_width_cm,
                            table_format=self._build_index_table_format_config(),
                            start_each_table_on_new_page=bool(self.start_each_table_on_new_page.get()),
                            page_label_column_title=self.page_label_column_title.get(),
                            pdf_page_column_title=self.pdf_page_column_title.get(),
                            include_doc_index_column=bool(self.include_doc_index_column.get()),
                            include_pdf_zero_index=False if include_pdf_zero_index_enabled else False,
                            stamp_enabled=False,
                            stamp_number_source=stamp_number_source,
                            stamp_prefix=stamp_prefix,
                            stamp_index_page_offset=0,
                        )
                        self._log(
                            f"Generated index DOCX: {updated_docx} "
                            f"(rows: {row_count}, sections: {section_count})"
                        )
                        for warning in style_warnings:
                            self._log(f"Warning: {warning}")
                else:
                    updated_rows = update_index_docx_pages(
                        index_docx,
                        updated_docx,
                        pagination,
                        page_label_column_title=self.page_label_column_title.get(),
                        include_page_range=bool(self.include_page_range.get()),
                        stamp_enabled=False,
                        stamp_number_source=stamp_number_source,
                        stamp_prefix=stamp_prefix,
                        stamp_index_page_offset=0,
                    )
                    self._log(f"Updated index DOCX: {updated_docx} (rows updated: {updated_rows})")
            else:
                if include_index_enabled:
                    self._log("Skipping DOCX index update/generation (no index DOCX path provided).")
                else:
                    self._log("Index disabled: skipping DOCX index update/generation.")

            final_pdf = merged_pdf
            index_pages_for_links: Optional[int] = None
            if include_index_enabled:
                if updated_docx is None:
                    raise RuntimeError("Cannot insert index: updated index DOCX was not generated.")
                self._log("Exporting updated DOCX index to exact-format PDF via Word...")
                index_pdf = output_folder / "Authorities Index (Word exact).pdf"
                export_docx_to_pdf_exact(updated_docx, index_pdf)
                self._log(f"Index PDF: {index_pdf}")

                with_index_pdf = output_folder / "_bundle_with_index.pdf"
                index_pages = prepend_index_pdf(index_pdf, merged_pdf, with_index_pdf)
                self._log(f"Prepended index pages: {index_pages}")
                final_pdf = with_index_pdf
                index_pages_for_links = index_pages

                needs_stamped_page_values = stamp_enabled and generate_index_enabled
                if generate_index_enabled and (include_pdf_zero_index_enabled or needs_stamped_page_values):
                    # Iteratively regenerate until the index-page offset used in rendered values is stable.
                    stable = False
                    for attempt in range(1, 4):
                        self._log(
                            "Regenerating index with final page offsets "
                            f"(offset {index_pages}, attempt {attempt}/3)..."
                        )
                        if use_existing_desc_enabled:
                            updated_rows = update_index_docx_pages(
                                updated_docx,
                                updated_docx,
                                pagination,
                                page_label_column_title=self.page_label_column_title.get(),
                                include_page_range=bool(self.include_page_range.get()),
                                stamp_enabled=stamp_enabled,
                                stamp_number_source=stamp_number_source,
                                stamp_prefix=stamp_prefix,
                                stamp_index_page_offset=index_pages,
                            )
                            self._log(f"Updated existing index table rows with final offsets: {updated_rows}")
                        else:
                            row_count, section_count, style_warnings = generate_index_docx_from_template(
                                index_docx,
                                updated_docx,
                                pdf_list,
                                index_rows,
                                pagination,
                                include_description_sorted=bool(self.include_sorted_by_description.get()),
                                use_existing_index_descriptions=False,
                                include_page_range=bool(self.include_page_range.get()),
                                page_column_width_cm=page_col_width_cm,
                                table_format=self._build_index_table_format_config(),
                                start_each_table_on_new_page=bool(self.start_each_table_on_new_page.get()),
                                page_label_column_title=self.page_label_column_title.get(),
                                pdf_page_column_title=self.pdf_page_column_title.get(),
                                include_doc_index_column=bool(self.include_doc_index_column.get()),
                                include_pdf_zero_index=include_pdf_zero_index_enabled,
                                pdf_zero_index_offset=index_pages,
                                stamp_enabled=stamp_enabled,
                                stamp_number_source=stamp_number_source,
                                stamp_prefix=stamp_prefix,
                                stamp_index_page_offset=index_pages,
                            )
                            for warning in style_warnings:
                                self._log(f"Warning: {warning}")

                        export_docx_to_pdf_exact(updated_docx, index_pdf)
                        refreshed_index_pages = prepend_index_pdf(index_pdf, merged_pdf, with_index_pdf)
                        self._log(f"Prepended index pages after index rebuild: {refreshed_index_pages}")
                        final_pdf = with_index_pdf
                        index_pages_for_links = refreshed_index_pages

                        if refreshed_index_pages == index_pages:
                            stable = True
                            self._log("Index values stabilized with final index-page offset.")
                            break
                        index_pages = refreshed_index_pages
                    if not stable:
                        self._log(
                            "Warning: Index-page offset did not stabilize after 3 attempts; "
                            f"using latest index-page offset {index_pages}."
                        )

            else:
                self._log("Index insertion disabled; keeping bundle without front index PDF.")

            if self.enable_stamp.get():
                self._log("Stamping final PDF with selected options...")
                number_source = self.stamp_number_source.get()
                stamp_prefix = self.stamp_prefix.get()
                stamp_pdf_in_place(
                    final_pdf,
                    self._log,
                    prefix=stamp_prefix,
                    number_source=number_source,
                    color_preset=color_preset_key_from_display(self.stamp_color_preset.get()),
                    position=self.stamp_position.get(),
                    stamp_mode=self.stamp_mode.get(),
                    bookmark_page_labels=bool(self.stamp_bookmark_page_labels.get()),
                    page_label_bookmark_parent=self.stamp_page_label_parent.get(),
                    set_page_labels_from_stamp=bool(self.stamp_set_page_labels_from_stamp.get()),
                )
                if number_source == "page-label" and stamp_prefix:
                    self._log("Updating PDF page labels to match stamped prefix...")
                    prefix_existing_page_labels_in_place(final_pdf, stamp_prefix)
                    if self.stamp_bookmark_page_labels.get():
                        updated = prefix_page_label_bookmark_titles_in_place(
                            final_pdf,
                            self.stamp_page_label_parent.get(),
                            stamp_prefix,
                        )
                        self._log(f"Page-label bookmarks updated: {updated}")

            if self.group_document_bookmarks.get():
                self._log("Grouping existing document bookmarks under parent folder...")
                grouped = group_top_level_bookmarks_under_parent_in_place(
                    final_pdf,
                    parent_title=self.document_bookmark_parent.get(),
                    preserve_top_titles=[
                        "Index",
                        self.stamp_page_label_parent.get(),
                    ],
                )
                self._log(f"Top-level bookmark groups moved: {grouped}")

            if self.compress_pdf.get():
                self._log("Compressing final PDF...")
                compress_pdf_in_place(final_pdf)

            if self.open_outlines.get():
                self._log("Setting PDF to open with bookmarks panel...")
                set_open_outlines_in_place(final_pdf)

            output_name = normalize_text(self.output_name.get()) or "Authorities Bundle Final.pdf"
            if not output_name.lower().endswith(".pdf"):
                output_name += ".pdf"
            target = output_folder / output_name
            final_pdf.replace(target)

            if include_index_enabled and link_rows_enabled and index_pages_for_links:
                self._log("Adding full-row index hyperlinks...")
                expected_starts = None
                if generate_index_enabled:
                    expected_starts = self._build_expected_index_start_pages(
                        pdf_list=pdf_list,
                        index_rows=index_rows,
                        pagination_rows=pagination,
                        use_existing_index_descriptions=use_existing_desc_enabled,
                    )
                links = add_full_row_links(
                    target,
                    pagination,
                    index_pages_for_links,
                    expected_start_pages=expected_starts,
                )
                self._log(f"Full-row links added: {links}")

            self._log("")
            self._log("Pipeline complete.")
            self._log(f"Final PDF: {target}")
            if updated_docx is not None:
                self._log(f"Updated DOCX: {updated_docx}")
            self._log(f"Pagination CSV: {report_csv}")
        except Exception as exc:
            self._log("")
            self._log(f"ERROR: {exc}")
            messagebox.showerror("Pipeline failed", str(exc))
        finally:
            self.progress.stop()
            self.run_btn.config(state="normal")


def main() -> None:
    app = BundleGui()
    app.mainloop()


if __name__ == "__main__":
    main()
