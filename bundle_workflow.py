import argparse
import csv
import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    from PyPDF2 import PdfReader, PdfWriter
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing required Python dependencies for the bundle workflow.\n"
        "Install them with:\n"
        "  python -m pip install -r requirements.txt\n"
        "or install the packages manually:\n"
        "  python -m pip install PyPDF2 python-docx"
    ) from exc


REQUIRED_INDEX_COLUMNS = ["Citation", "Case title", "C ¶", "D2 ¶", "Page"]

PROCEDURAL_EXPECTED = [
    ("CPR 3.4", "03.4 CPR.pdf"),
    ("PD 3A, para 1.2", "03.1.2 PD.pdf"),
    ("CPR 16.4", "16.4 CPR.pdf"),
    ("PD 16, para 8.2", "16.8.2 PD.pdf"),
    ("CPR 17.1", "17.1 CPR.pdf"),
    ("White Book 2026, paras 17.3.5–17.3.11", "17.3.5–17.3.11 WHITEBOOK 2026.pdf"),
    ("PD 53B, para 8.1", "53B.8.1 PD.pdf"),
]

LENGTH_LIMIT = 8000
WORKFLOW_BASENAME = "authorities_bundle"


@dataclass
class IndexEntry:
    position: int
    citation: str
    case_title: str
    c_para: str
    d2_para: str
    page: str


@dataclass
class PdfInfo:
    abs_path: Path
    rel_path: str
    filename: str
    stem: str
    page_count: int
    encrypted: bool
    unreadable: bool
    ocr_likely_needed: bool
    text_sample: str
    norm_stem: str
    norm_text: str


def normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", value or "")
    text = (
        text.replace("–", "-")
        .replace("—", "-")
        .replace("‑", "-")
        .replace("\u00A0", " ")
        .replace("\n", " ")
        .replace("\r", " ")
    )
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_loose(value: str) -> str:
    value = normalize_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def bookmark_label(citation: str, case_title: str) -> str:
    c = normalize_text(citation)
    t = normalize_text(case_title)
    return f"{c}:: {t}"


def split_citation_parts(citation: str) -> List[str]:
    cleaned = normalize_text(citation)
    parts = re.split(r"[;\n|]+|(?<!\w)/(?!\w)", cleaned)
    more = []
    for part in parts:
        more.extend(re.split(r"\s{2,}", part))
    out = []
    for part in more:
        p = normalize_text(part)
        if len(p) >= 4:
            out.append(p)
    return list(dict.fromkeys(out))


def title_keywords(case_title: str) -> List[str]:
    words = re.findall(r"[A-Za-z][A-Za-z'\-]{2,}", normalize_text(case_title))
    stop = {"the", "and", "for", "with", "from", "against", "para", "paras", "book", "white", "power"}
    return [w.lower() for w in words if w.lower() not in stop]


def require_python_docx() -> None:
    try:
        import docx  # noqa: F401
    except ModuleNotFoundError:
        raise SystemExit(
            "Missing dependency: python-docx.\n"
            "Install locally with:\n"
            "  pip install python-docx\n"
            "Then rerun the command."
        )


def get_index_table_and_headers(document) -> Tuple[object, Dict[str, int]]:
    table_with_headers = None
    header_map: Dict[str, int] = {}
    for table in document.tables:
        if not table.rows:
            continue
        headers = [normalize_text(cell.text) for cell in table.rows[0].cells]
        mapping = {}
        for col in REQUIRED_INDEX_COLUMNS:
            if col in headers:
                mapping[col] = headers.index(col)
        if len(mapping) == len(REQUIRED_INDEX_COLUMNS):
            table_with_headers = table
            header_map = mapping
            break
    if table_with_headers is None:
        raise SystemExit(
            "Could not find an index table containing columns: "
            + ", ".join(REQUIRED_INDEX_COLUMNS)
        )
    return table_with_headers, header_map


def read_index_entries(word_index_path: Path) -> List[IndexEntry]:
    require_python_docx()
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError

    try:
        document = Document(str(word_index_path))
    except PackageNotFoundError as exc:
        raise SystemExit(
            "Word index file could not be opened as a valid .docx package. "
            "Please open it in Word and Save As a modern .docx file, then rerun.\n"
            f"File: {word_index_path}"
        ) from exc
    table_with_headers, header_map = get_index_table_and_headers(document)

    entries: List[IndexEntry] = []
    for row in table_with_headers.rows[1:]:
        cells = row.cells
        citation = normalize_text(cells[header_map["Citation"]].text)
        case_title = normalize_text(cells[header_map["Case title"]].text)
        c_para = normalize_text(cells[header_map["C ¶"]].text)
        d2_para = normalize_text(cells[header_map["D2 ¶"]].text)
        page = normalize_text(cells[header_map["Page"]].text)
        if not citation and not case_title:
            continue
        entries.append(
            IndexEntry(
                position=len(entries) + 1,
                citation=citation,
                case_title=case_title,
                c_para=c_para,
                d2_para=d2_para,
                page=page,
            )
        )

    if not entries:
        raise SystemExit("Index table was found but no non-empty rows were detected.")
    return entries


def should_exclude_pdf(path: Path, source_root: Path, output_root: Path, output_pdf_name: str) -> bool:
    if path.name.lower() == output_pdf_name.lower():
        return True
    try:
        path.relative_to(output_root)
        rel_to_output = True
    except ValueError:
        rel_to_output = False
    if rel_to_output and path.stem.lower().startswith(WORKFLOW_BASENAME):
        return True
    stem = normalize_loose(path.stem)
    if "authoritiesbundle" in stem and ("merged" in stem or "final" in stem or "output" in stem):
        return True
    return False


def extract_pdf_text(pdf_path: Path, max_pages: int = 4) -> Tuple[int, bool, bool, bool, str]:
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return 0, False, True, False, ""

    encrypted = bool(getattr(reader, "is_encrypted", False))
    if encrypted:
        try:
            reader.decrypt("")
        except Exception:
            return 0, True, True, False, ""

    try:
        page_count = len(reader.pages)
    except Exception:
        return 0, encrypted, True, False, ""

    sample_parts = []
    upper = min(max_pages, page_count)
    for page_index in range(upper):
        try:
            extracted = reader.pages[page_index].extract_text() or ""
        except Exception:
            extracted = ""
        if extracted:
            sample_parts.append(extracted)

    sample = normalize_text(" ".join(sample_parts))
    alnum_chars = re.findall(r"[A-Za-z0-9]", sample)
    ocr_likely_needed = len(alnum_chars) < 30
    return page_count, encrypted, False, ocr_likely_needed, sample


def inventory_pdfs(
    source_folder: Path,
    output_folder: Path,
    output_pdf_name: str,
) -> Tuple[List[PdfInfo], List[str]]:
    excluded = []
    pdfs: List[PdfInfo] = []
    for path in sorted(source_folder.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.lower() != ".pdf":
            continue
        if should_exclude_pdf(path, source_folder, output_folder, output_pdf_name):
            excluded.append(str(path))
            continue

        rel = path.relative_to(source_folder).as_posix()
        page_count, encrypted, unreadable, ocr_likely_needed, text_sample = extract_pdf_text(path)
        pdfs.append(
            PdfInfo(
                abs_path=path,
                rel_path=rel,
                filename=path.name,
                stem=path.stem,
                page_count=page_count,
                encrypted=encrypted,
                unreadable=unreadable,
                ocr_likely_needed=ocr_likely_needed,
                text_sample=text_sample,
                norm_stem=normalize_loose(path.stem),
                norm_text=normalize_loose(text_sample[:LENGTH_LIMIT]),
            )
        )
    return pdfs, excluded


def score_candidate(entry: IndexEntry, pdf: PdfInfo) -> Tuple[int, List[str]]:
    score = 0
    evidence: List[str] = []
    citation_parts = split_citation_parts(entry.citation)
    case_title_norm = normalize_loose(entry.case_title)
    title_keys = title_keywords(entry.case_title)

    for part in citation_parts:
        p = normalize_loose(part)
        if not p:
            continue
        if p in pdf.norm_stem:
            score += 6
            evidence.append(f"citation '{part}' in filename")
        elif p in pdf.norm_text:
            score += 4
            evidence.append(f"citation '{part}' in extracted text")

    if case_title_norm and case_title_norm in pdf.norm_text:
        score += 10
        evidence.append("full case title in extracted text")
    elif case_title_norm and case_title_norm in pdf.norm_stem:
        score += 8
        evidence.append("full case title in filename")

    if title_keys:
        stem_hits = sum(1 for k in title_keys if k in pdf.norm_stem)
        text_hits = sum(1 for k in title_keys if k in pdf.norm_text)
        if stem_hits:
            score += min(6, stem_hits)
            evidence.append(f"{stem_hits} title keywords in filename")
        if text_hits:
            score += min(8, text_hits)
            evidence.append(f"{text_hits} title keywords in extracted text")

    if pdf.encrypted:
        evidence.append("file is encrypted")
    if pdf.unreadable:
        evidence.append("file is unreadable")
    if pdf.ocr_likely_needed:
        evidence.append("low extractable text; OCR likely required")
    return score, evidence


def preference_rule_for_entry(entry: IndexEntry) -> Optional[Dict[str, object]]:
    marker = f"{entry.citation} {entry.case_title}".lower()
    rules = [
        {
            "match": "lakatamia",
            "prefer": "[2023] EWHC 1874 (Comm).pdf",
            "avoid": ["[2023] EWHC 3023 (Comm).pdf"],
        },
        {
            "match": "three rivers",
            "prefer": "[2001] UKHL 16.pdf",
            "avoid": ["[2001] 2 All ER 513.pdf"],
        },
        {
            "match": "investment trust companies",
            "prefer": "[2012] STC 1150.pdf",
            "avoid": ["[2012] EWHC 458 (Ch).pdf"],
        },
        {
            "match": "mukhtar v saleem",
            "prefer": "[2018] EWHC 1729 (QB).pdf",
            "avoid": [],
        },
        {
            "match": "fm capital partners ltd v marino",
            "prefer": "[2019] EWHC 768 (Comm).pdf",
            "avoid": ["[2020] EWCA Civ 245.pdf"],
            "must_exist": True,
        },
    ]
    for rule in rules:
        if rule["match"] in marker:
            return rule
    return None


def procedural_expected_filename(entry: IndexEntry) -> Optional[str]:
    for citation, filename in PROCEDURAL_EXPECTED:
        if normalize_loose(citation) == normalize_loose(entry.citation):
            return filename
    return None


def apply_preference(
    entry: IndexEntry,
    candidates: List[Tuple[PdfInfo, int, List[str]]],
    all_pdfs: List[PdfInfo],
    issues: List[str],
) -> Tuple[Optional[PdfInfo], List[PdfInfo], str, Optional[str]]:
    rule = preference_rule_for_entry(entry)
    alternatives: List[PdfInfo] = []
    note = ""
    if rule:
        preferred = next((p for p in all_pdfs if p.filename.lower() == rule["prefer"].lower()), None)
        if preferred:
            alternatives = [c[0] for c in candidates if c[0].abs_path != preferred.abs_path]
            note = f"preferred file applied: {rule['prefer']}"
            for avoided in rule.get("avoid", []):
                if any(p.filename.lower() == avoided.lower() for p in all_pdfs):
                    issues.append(
                        f"Entry {entry.position} '{entry.citation}': kept preferred version "
                        f"and excluded alternative '{avoided}'."
                    )
            return preferred, alternatives, note, None
        if rule.get("must_exist"):
            issues.append(
                f"Entry {entry.position} '{entry.citation}': required preferred file "
                f"'{rule['prefer']}' not found; unresolved."
            )
            return None, alternatives, "required preferred file missing", (
                f"Required preferred file missing: {rule['prefer']}. "
                "No substitution allowed."
            )
        issues.append(
            f"Entry {entry.position} '{entry.citation}': preferred file '{rule['prefer']}' "
            "was not found in source inventory."
        )
    expected_proc = procedural_expected_filename(entry)
    if expected_proc:
        expected_file = next((p for p in all_pdfs if p.filename.lower() == expected_proc.lower()), None)
        if expected_file:
            alternatives = [c[0] for c in candidates if c[0].abs_path != expected_file.abs_path]
            note = f"procedural expected filename applied: {expected_proc}"
            return expected_file, alternatives, note, None
    return None, alternatives, note, None


def filename_confirms_match(entry: IndexEntry, pdf: PdfInfo) -> bool:
    norm_file = pdf.norm_stem
    if not norm_file:
        return False
    for part in split_citation_parts(entry.citation):
        token = normalize_loose(part)
        if token and token in norm_file:
            return True
    title_norm = normalize_loose(entry.case_title)
    if title_norm and title_norm in norm_file:
        return True
    keys = title_keywords(entry.case_title)
    key_hits = sum(1 for k in keys if k in norm_file)
    return key_hits >= 2


def verify_procedural_order(entries: List[IndexEntry], issues: List[str]) -> None:
    if len(entries) < len(PROCEDURAL_EXPECTED):
        issues.append("Index has fewer than seven rows; cannot satisfy required procedural extract ordering.")
        return
    for i, (expected_citation, _) in enumerate(PROCEDURAL_EXPECTED):
        actual = entries[i].citation
        if normalize_loose(actual) != normalize_loose(expected_citation):
            issues.append(
                f"Procedural order mismatch at row {i+1}: expected citation '{expected_citation}', "
                f"found '{actual}'."
            )


def match_entries(entries: List[IndexEntry], pdfs: List[PdfInfo]) -> Tuple[List[Dict[str, str]], Dict[str, object]]:
    issues: List[str] = []
    alternatives: List[Dict[str, str]] = []
    unresolved: List[Dict[str, str]] = []
    unmatched: List[Dict[str, str]] = []
    used_paths: set = set()
    verify_procedural_order(entries, issues)
    explicit_unindexed = {"[2020] ewca civ 245.pdf", "[2023] ewhc 3023 (comm).pdf"}
    manifest_rows: List[Dict[str, str]] = []

    for entry in entries:
        status = "pending_approval"
        inclusion = "proposed_include"
        evidence_parts = []
        rel_path = ""
        selected: Optional[PdfInfo] = None

        if not entry.citation or not entry.case_title:
            status = "blocked_blank_index_value"
            inclusion = "blocked"
            unresolved.append(
                {
                    "index_position": str(entry.position),
                    "citation": entry.citation,
                    "case_title": entry.case_title,
                    "issue": "Blank Citation or Case title cell.",
                }
            )

        candidates_scored = []
        for pdf in pdfs:
            score, ev = score_candidate(entry, pdf)
            if score >= 6:
                candidates_scored.append((pdf, score, ev))
        candidates_scored.sort(key=lambda item: item[1], reverse=True)

        pref_selected, pref_alts, pref_note, pref_block_reason = apply_preference(
            entry,
            candidates_scored,
            pdfs,
            issues,
        )
        if pref_note:
            evidence_parts.append(pref_note)
        if pref_selected:
            selected = pref_selected
            evidence_parts.append("selected by preference rule")
            if pref_alts:
                for alt in pref_alts:
                    alternatives.append(
                        {
                            "index_position": str(entry.position),
                            "citation": entry.citation,
                            "case_title": entry.case_title,
                            "selected_filename": pref_selected.filename,
                            "alternative_filename": alt.filename,
                            "reason": "Alternative version for same entry; not merged automatically.",
                        }
                    )
        elif pref_block_reason:
            status = "blocked_preference_requirement"
            inclusion = "blocked"
            evidence_parts.append(pref_block_reason)
            unresolved.append(
                {
                    "index_position": str(entry.position),
                    "citation": entry.citation,
                    "case_title": entry.case_title,
                    "issue": pref_block_reason,
                }
            )
        elif candidates_scored:
            top_score = candidates_scored[0][1]
            top_group = [c for c in candidates_scored if c[1] >= top_score - 2]
            if len(top_group) == 1:
                selected = top_group[0][0]
                evidence_parts.extend(top_group[0][2][:3])
            else:
                status = "blocked_ambiguous_match"
                inclusion = "blocked"
                candidate_names = ", ".join(c[0].filename for c in top_group)
                evidence_parts.append(f"ambiguous top candidates: {candidate_names}")
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": f"Ambiguous candidates; choose one: {candidate_names}",
                    }
                )
                for c in top_group:
                    alternatives.append(
                        {
                            "index_position": str(entry.position),
                            "citation": entry.citation,
                            "case_title": entry.case_title,
                            "selected_filename": "",
                            "alternative_filename": c[0].filename,
                            "reason": "Ambiguous candidate requiring user choice.",
                        }
                    )
        else:
            if status == "pending_approval":
                status = "blocked_no_match"
                inclusion = "blocked"
                unmatched.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": "No confident PDF match found.",
                    }
                )

        if selected:
            rel_path = selected.rel_path
            evidence_parts.append(f"matched file: {selected.filename}")
            if not filename_confirms_match(entry, selected):
                status = "blocked_content_only_match"
                inclusion = "blocked"
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": (
                            f"Match not confirmed by filename identifiers for '{selected.filename}'. "
                            "Likely text-only cross-reference; requires user confirmation."
                        ),
                    }
                )
            if selected.unreadable:
                status = "blocked_unreadable_pdf"
                inclusion = "blocked"
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": f"Matched file unreadable: {selected.filename}",
                    }
                )
            elif selected.encrypted:
                status = "blocked_encrypted_pdf"
                inclusion = "blocked"
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": f"Matched file encrypted: {selected.filename}",
                    }
                )
            elif selected.ocr_likely_needed:
                evidence_parts.append("scan/OCR warning")
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": f"Matched file likely scan requiring OCR: {selected.filename}",
                    }
                )
                if status == "pending_approval":
                    status = "warning_ocr_likely_needed"
            if selected.abs_path in used_paths:
                evidence_parts.append("possible duplicate use of same PDF across entries")
                unresolved.append(
                    {
                        "index_position": str(entry.position),
                        "citation": entry.citation,
                        "case_title": entry.case_title,
                        "issue": f"Possible duplicate document mapping: {selected.filename}",
                    }
                )
            else:
                used_paths.add(selected.abs_path)

            expected_proc_name = procedural_expected_filename(entry)
            if expected_proc_name and selected.filename.lower() != expected_proc_name.lower():
                issues.append(
                    f"Procedural entry row {entry.position} expected filename '{expected_proc_name}' "
                    f"but matched '{selected.filename}'."
                )
            if expected_proc_name:
                keys = title_keywords(entry.case_title)
                key_hits = sum(1 for k in keys[:6] if k in selected.norm_text)
                if keys and key_hits == 0:
                    unresolved.append(
                        {
                            "index_position": str(entry.position),
                            "citation": entry.citation,
                            "case_title": entry.case_title,
                            "issue": f"Procedural extract content mismatch warning for '{selected.filename}'",
                        }
                    )
                    if status == "pending_approval":
                        status = "warning_content_mismatch"

        manifest_rows.append(
            {
                "Index position": str(entry.position),
                "Citation": entry.citation,
                "Case title": entry.case_title,
                "Exact relative PDF path": rel_path,
                "Match evidence and status": "; ".join(dict.fromkeys(evidence_parts)) or status,
                "Inclusion status": inclusion if status.startswith("blocked_") else status,
                "Proposed bookmark label": bookmark_label(entry.citation, entry.case_title),
            }
        )

    unused = []
    for pdf in pdfs:
        if pdf.abs_path in used_paths:
            continue
        lower_name = pdf.filename.lower()
        reason = "Not mapped to any index entry."
        if lower_name in explicit_unindexed:
            reason = "Explicitly treated as unindexed judgment unless user approves inclusion."
        unused.append(
            {
                "filename": pdf.filename,
                "relative_path": pdf.rel_path,
                "reason": reason,
            }
        )

    return manifest_rows, {
        "issues": issues,
        "unresolved": unresolved,
        "unmatched": unmatched,
        "alternatives": alternatives,
        "unused_pdfs": unused,
    }


def write_csv(path: Path, fieldnames: List[str], rows: List[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def prepare_workflow(
    source_folder: Path,
    word_index_path: Path,
    output_folder: Path,
    output_pdf_name: str,
) -> None:
    if not source_folder.is_dir():
        raise SystemExit(f"Source folder not found: {source_folder}")
    if not word_index_path.is_file():
        raise SystemExit(f"Word index file not found: {word_index_path}")
    output_folder.mkdir(parents=True, exist_ok=True)

    entries = read_index_entries(word_index_path)
    pdfs, excluded = inventory_pdfs(source_folder, output_folder, output_pdf_name)
    manifest_rows, review = match_entries(entries, pdfs)

    manifest_path = output_folder / f"{WORKFLOW_BASENAME}_manifest.csv"
    write_csv(
        manifest_path,
        [
            "Index position",
            "Citation",
            "Case title",
            "Exact relative PDF path",
            "Match evidence and status",
            "Inclusion status",
            "Proposed bookmark label",
        ],
        manifest_rows,
    )

    write_csv(
        output_folder / f"{WORKFLOW_BASENAME}_unmatched_index_entries.csv",
        ["index_position", "citation", "case_title", "issue"],
        review["unmatched"],
    )
    write_csv(
        output_folder / f"{WORKFLOW_BASENAME}_alternative_versions.csv",
        ["index_position", "citation", "case_title", "selected_filename", "alternative_filename", "reason"],
        review["alternatives"],
    )
    write_csv(
        output_folder / f"{WORKFLOW_BASENAME}_unused_pdfs.csv",
        ["filename", "relative_path", "reason"],
        review["unused_pdfs"],
    )
    write_csv(
        output_folder / f"{WORKFLOW_BASENAME}_unresolved_issues.csv",
        ["index_position", "citation", "case_title", "issue"],
        review["unresolved"],
    )

    summary = {
        "source_folder": str(source_folder),
        "word_index_path": str(word_index_path),
        "output_folder": str(output_folder),
        "output_pdf_name": output_pdf_name,
        "inventory_count": len(pdfs),
        "excluded_generated_or_output_files": excluded,
        "issues": review["issues"],
        "counts": {
            "manifest_rows": len(manifest_rows),
            "unmatched_entries": len(review["unmatched"]),
            "unresolved_issues": len(review["unresolved"]),
            "alternative_versions": len(review["alternatives"]),
            "unused_pdfs": len(review["unused_pdfs"]),
        },
        "next_step": (
            "Review and edit manifest Inclusion status. "
            "Set approved rows to 'approved'. Set approved omissions to 'omit_approved'. "
            "Then run the merge command."
        ),
    }
    summary_path = output_folder / f"{WORKFLOW_BASENAME}_review_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"Manifest created: {manifest_path}")
    print(f"Review summary: {summary_path}")
    print("Stop for approval before merging.")


def read_manifest(manifest_path: Path) -> List[Dict[str, str]]:
    if not manifest_path.is_file():
        raise SystemExit(f"Manifest not found: {manifest_path}")
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)
    if not rows:
        raise SystemExit("Manifest is empty.")
    return rows


def prompt_yes_no(message: str) -> bool:
    answer = input(f"{message} [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def set_cell_text_preserving_style(cell, text: str) -> None:
    value = normalize_text(text)
    if not cell.paragraphs:
        cell.text = value
        return
    first_paragraph = cell.paragraphs[0]
    if first_paragraph.runs:
        first_paragraph.runs[0].text = value
        for run in first_paragraph.runs[1:]:
            run.text = ""
    else:
        first_paragraph.add_run(value)
    for paragraph in cell.paragraphs[1:]:
        for run in paragraph.runs:
            run.text = ""


def ensure_output_write_allowed(output_pdf_path: Path, overwrite: bool) -> None:
    if output_pdf_path.exists() and not overwrite:
        if not prompt_yes_no(f"Output file exists: {output_pdf_path}. Overwrite?"):
            raise SystemExit("Aborted by user to avoid overwrite.")


def read_pagination_rows(pagination_report_path: Path) -> List[Dict[str, str]]:
    if not pagination_report_path.is_file():
        raise SystemExit(f"Pagination report not found: {pagination_report_path}")
    with pagination_report_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if not rows:
        raise SystemExit("Pagination report is empty.")
    required = {"Entry", "Start PDF page position"}
    missing = [col for col in required if col not in rows[0]]
    if missing:
        raise SystemExit(
            "Pagination report is missing required columns: " + ", ".join(missing)
        )
    return rows


def populate_word_index_pages(
    word_index_path: Path,
    pagination_report_path: Path,
    output_word_index_path: Path,
    overwrite: bool,
) -> None:
    require_python_docx()
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError

    if not word_index_path.is_file():
        raise SystemExit(f"Word index file not found: {word_index_path}")
    if output_word_index_path.exists() and not overwrite:
        if not prompt_yes_no(f"Output Word file exists: {output_word_index_path}. Overwrite?"):
            raise SystemExit("Aborted by user to avoid overwrite.")

    pagination_rows = read_pagination_rows(pagination_report_path)
    page_by_index_position: Dict[int, str] = {}
    for row in pagination_rows:
        entry = normalize_text(row.get("Entry", ""))
        if not entry.isdigit():
            continue
        page_by_index_position[int(entry)] = normalize_text(row.get("Start PDF page position", ""))

    try:
        document = Document(str(word_index_path))
    except PackageNotFoundError as exc:
        raise SystemExit(
            "Word index file could not be opened as a valid .docx package.\n"
            f"File: {word_index_path}"
        ) from exc

    table, header_map = get_index_table_and_headers(document)
    page_col = header_map["Page"]
    index_row_position = 0
    updates = 0
    missing_positions: List[int] = []

    for row in table.rows[1:]:
        cells = row.cells
        citation = normalize_text(cells[header_map["Citation"]].text)
        case_title = normalize_text(cells[header_map["Case title"]].text)
        if not citation and not case_title:
            continue
        index_row_position += 1
        page_value = page_by_index_position.get(index_row_position)
        if page_value is None:
            missing_positions.append(index_row_position)
            continue
        set_cell_text_preserving_style(cells[page_col], page_value)
        updates += 1

    output_word_index_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_word_index_path))

    report = {
        "input_word_index": str(word_index_path),
        "pagination_report": str(pagination_report_path),
        "output_word_index": str(output_word_index_path),
        "updated_rows": updates,
        "missing_index_positions": missing_positions,
        "note": (
            "Page column values were set to the merged PDF start positions "
            "for matching bookmark/index entries."
        ),
    }
    report_path = output_word_index_path.parent / f"{WORKFLOW_BASENAME}_word_index_update_report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Updated Word index copy: {output_word_index_path}")
    print(f"Update report: {report_path}")


def add_top_level_bookmark(writer: PdfWriter, title: str, page_index: int) -> None:
    try:
        writer.add_outline_item(title, page_index)
    except Exception:
        writer.addBookmark(title, page_index)


def collect_merge_rows(manifest_rows: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    approved = []
    blocked = []
    for row in manifest_rows:
        status = normalize_text(row.get("Inclusion status", "")).lower()
        if status == "approved":
            approved.append(row)
        elif status == "omit_approved":
            continue
        else:
            blocked.append(row)
    return approved, blocked


def extract_top_level_outlines(reader: PdfReader) -> List[object]:
    try:
        outlines = reader.outline
    except Exception:
        return []
    top_level = []
    for item in outlines:
        if isinstance(item, list):
            continue
        top_level.append(item)
    return top_level


def merge_bundle(
    source_folder: Path,
    output_folder: Path,
    output_pdf_name: str,
    manifest_path: Path,
    include_index_pdf: str,
    index_pdf_path: Optional[Path],
    overwrite: bool,
) -> None:
    manifest_rows = read_manifest(manifest_path)
    approved_rows, blocked_rows = collect_merge_rows(manifest_rows)
    if blocked_rows:
        raise SystemExit(
            "Manifest still contains unresolved rows.\n"
            "Only 'approved' (included) and 'omit_approved' (omitted with approval) are allowed before merge."
        )
    if not approved_rows:
        raise SystemExit("No approved rows found in manifest.")

    include_index = include_index_pdf
    if include_index == "ask":
        include_index = "yes" if prompt_yes_no("Include an index PDF at the front?") else "no"
    if include_index == "yes" and not index_pdf_path:
        raise SystemExit(
            "Index inclusion requested but --index-pdf was not provided.\n"
            "Provide a PDF index file path or run with --include-index-pdf no."
        )
    if include_index == "yes" and index_pdf_path and not index_pdf_path.is_file():
        raise SystemExit(f"Index PDF not found: {index_pdf_path}")

    output_pdf_path = output_folder / output_pdf_name
    ensure_output_write_allowed(output_pdf_path, overwrite)
    output_folder.mkdir(parents=True, exist_ok=True)

    writer = PdfWriter()
    pagination_rows = []
    expected_total_pages = 0
    current_page = 1
    merged_inputs = []

    if include_index == "yes":
        index_reader = PdfReader(str(index_pdf_path))
        index_pages = len(index_reader.pages)
        for p in index_reader.pages:
            writer.add_page(p)
        add_top_level_bookmark(writer, "Index", 0)
        pagination_rows.append(
            {
                "Entry": "Index",
                "Citation": "",
                "Case title": "Index",
                "Filename": index_pdf_path.name,
                "Source page count": str(index_pages),
                "Start PDF page position": "1",
                "End PDF page position": str(index_pages),
            }
        )
        expected_total_pages += index_pages
        current_page += index_pages
        merged_inputs.append(str(index_pdf_path))

    for row in approved_rows:
        rel_path = row.get("Exact relative PDF path", "").strip()
        citation = normalize_text(row.get("Citation", ""))
        case_title = normalize_text(row.get("Case title", ""))
        if not citation or not case_title:
            raise SystemExit("Manifest contains approved row with blank Citation or Case title.")
        pdf_path = (source_folder / rel_path).resolve()
        if not pdf_path.is_file():
            raise SystemExit(f"Approved PDF not found: {pdf_path}")

        reader = PdfReader(str(pdf_path))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise SystemExit(f"Approved PDF is encrypted and cannot be read: {pdf_path}")
        page_count = len(reader.pages)
        start = current_page
        end = start + page_count - 1
        for page in reader.pages:
            writer.add_page(page)
        add_top_level_bookmark(writer, bookmark_label(citation, case_title), start - 1)
        pagination_rows.append(
            {
                "Entry": row.get("Index position", ""),
                "Citation": citation,
                "Case title": case_title,
                "Filename": rel_path,
                "Source page count": str(page_count),
                "Start PDF page position": str(start),
                "End PDF page position": str(end),
            }
        )
        expected_total_pages += page_count
        current_page = end + 1
        merged_inputs.append(str(pdf_path))

    with output_pdf_path.open("wb") as handle:
        writer.write(handle)

    pagination_path = output_folder / f"{WORKFLOW_BASENAME}_pagination_report.csv"
    write_csv(
        pagination_path,
        [
            "Entry",
            "Citation",
            "Case title",
            "Filename",
            "Source page count",
            "Start PDF page position",
            "End PDF page position",
        ],
        pagination_rows,
    )

    validation = {
        "output_pdf": str(output_pdf_path),
        "checks": [],
        "passed": True,
    }

    merged_reader = PdfReader(str(output_pdf_path))
    actual_total_pages = len(merged_reader.pages)
    check_total = actual_total_pages == expected_total_pages
    validation["checks"].append(
        {
            "name": "Total pages equal sum of approved inputs",
            "passed": check_total,
            "expected": expected_total_pages,
            "actual": actual_total_pages,
        }
    )

    outlines = extract_top_level_outlines(merged_reader)
    expected_bookmarks = len(approved_rows) + (1 if include_index == "yes" else 0)
    check_bookmark_count = len(outlines) == expected_bookmarks
    validation["checks"].append(
        {
            "name": "Exactly one top-level bookmark per included entry",
            "passed": check_bookmark_count,
            "expected": expected_bookmarks,
            "actual": len(outlines),
        }
    )

    expected_labels = []
    expected_starts = []
    if include_index == "yes":
        expected_labels.append("Index")
        expected_starts.append(1)
    expected_labels.extend([bookmark_label(r["Citation"], r["Case title"]) for r in approved_rows])
    expected_starts.extend([int(r["Start PDF page position"]) for r in pagination_rows if r["Entry"] != "Index"])
    if include_index == "yes":
        expected_starts = [1] + expected_starts

    actual_labels = []
    actual_starts = []
    for outline_item in outlines:
        title = getattr(outline_item, "title", "")
        actual_labels.append(normalize_text(title))
        try:
            actual_starts.append(merged_reader.get_destination_page_number(outline_item) + 1)
        except Exception:
            actual_starts.append(-1)

    check_labels = actual_labels == expected_labels
    check_starts = actual_starts == expected_starts
    validation["checks"].append(
        {
            "name": "Bookmark labels match Citation:: Case title exactly",
            "passed": check_labels,
            "expected": expected_labels,
            "actual": actual_labels,
        }
    )
    validation["checks"].append(
        {
            "name": "Bookmarks point to each document's first page",
            "passed": check_starts,
            "expected": expected_starts,
            "actual": actual_starts,
        }
    )

    order_ok = True
    for row in pagination_rows:
        start = int(row["Start PDF page position"])
        end = int(row["End PDF page position"])
        if not (1 <= start <= end <= actual_total_pages):
            order_ok = False
            break
    validation["checks"].append(
        {
            "name": "Pagination report positions align with merged PDF",
            "passed": order_ok,
        }
    )

    duplicate_rel_paths: Dict[str, int] = {}
    for row in approved_rows:
        rel = normalize_text(row.get("Exact relative PDF path", ""))
        if not rel:
            continue
        duplicate_rel_paths[rel] = duplicate_rel_paths.get(rel, 0) + 1
    duplicate_rel_paths = {k: v for k, v in duplicate_rel_paths.items() if v > 1}
    duplicate_inputs = bool(duplicate_rel_paths)
    approved_duplicate_markers = (
        "user approved current match",
        "user approved duplicate",
    )
    duplicate_inputs_approved = True
    if duplicate_inputs:
        for rel_path in duplicate_rel_paths:
            rel_has_approval_marker = False
            for row in approved_rows:
                if normalize_text(row.get("Exact relative PDF path", "")) != rel_path:
                    continue
                evidence = normalize_text(row.get("Match evidence and status", "")).lower()
                if any(marker in evidence for marker in approved_duplicate_markers):
                    rel_has_approval_marker = True
            if not rel_has_approval_marker:
                duplicate_inputs_approved = False
                break
    duplicate_check_passed = (not duplicate_inputs) or duplicate_inputs_approved
    validation["checks"].append(
        {
            "name": "No unapproved duplicates or unindexed documents included",
            "passed": duplicate_check_passed,
            "details": (
                f"Approved duplicate mappings present: {duplicate_rel_paths}"
                if duplicate_inputs and duplicate_inputs_approved
                else "Duplicate source files detected in merge inputs."
                if duplicate_inputs
                else "OK"
            ),
        }
    )

    validation["passed"] = all(check["passed"] for check in validation["checks"])
    validation_path = output_folder / f"{WORKFLOW_BASENAME}_validation_report.json"
    validation_path.write_text(json.dumps(validation, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"Merged output PDF: {output_pdf_path}")
    print(f"Pagination report: {pagination_path}")
    print(f"Validation report: {validation_path}")
    if not validation["passed"]:
        raise SystemExit("Validation failed. Review validation report for details.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Authorities bundle workflow: prepare a review manifest from Word index + PDFs, "
            "then merge approved entries with required bookmarks."
        )
    )
    sub = parser.add_subparsers(dest="command", required=True)

    prepare = sub.add_parser("prepare", help="Create review manifest and issue reports; does not merge.")
    prepare.add_argument("--source-folder", required=True, help="Folder containing source PDFs.")
    prepare.add_argument("--word-index", required=True, help="Path to updated Word index (.docx).")
    prepare.add_argument("--output-folder", required=True, help="Folder for manifest/reports.")
    prepare.add_argument("--output-pdf-name", default="Authorities Bundle.pdf", help="Target merged PDF filename.")

    merge = sub.add_parser("merge", help="Merge approved rows in manifest into output PDF and validate.")
    merge.add_argument("--source-folder", required=True, help="Folder containing source PDFs.")
    merge.add_argument("--output-folder", required=True, help="Folder for output PDF and reports.")
    merge.add_argument("--output-pdf-name", default="Authorities Bundle.pdf", help="Target merged PDF filename.")
    merge.add_argument("--manifest", required=True, help="Path to edited manifest CSV.")
    merge.add_argument(
        "--include-index-pdf",
        choices=["ask", "yes", "no"],
        default="ask",
        help="Whether to include index PDF at front.",
    )
    merge.add_argument("--index-pdf", help="Path to index PDF if including index at front.")
    merge.add_argument("--overwrite", action="store_true", help="Overwrite output PDF without prompt.")

    populate = sub.add_parser(
        "populate-index-pages",
        help="Populate Word index Page column from pagination report, writing to a copy.",
    )
    populate.add_argument("--word-index", required=True, help="Path to source Word index (.docx).")
    populate.add_argument(
        "--pagination-report",
        required=True,
        help="Path to authorities_bundle_pagination_report.csv.",
    )
    populate.add_argument("--output-word-index", required=True, help="Output .docx path for updated copy.")
    populate.add_argument("--overwrite", action="store_true", help="Overwrite output .docx without prompt.")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "prepare":
        source_folder = Path(args.source_folder).resolve()
        output_folder = Path(args.output_folder).resolve()
        prepare_workflow(
            source_folder=source_folder,
            word_index_path=Path(args.word_index).resolve(),
            output_folder=output_folder,
            output_pdf_name=args.output_pdf_name,
        )
        return

    if args.command == "merge":
        source_folder = Path(args.source_folder).resolve()
        output_folder = Path(args.output_folder).resolve()
        merge_bundle(
            source_folder=source_folder,
            output_folder=output_folder,
            output_pdf_name=args.output_pdf_name,
            manifest_path=Path(args.manifest).resolve(),
            include_index_pdf=args.include_index_pdf,
            index_pdf_path=Path(args.index_pdf).resolve() if args.index_pdf else None,
            overwrite=bool(args.overwrite),
        )
        return

    if args.command == "populate-index-pages":
        populate_word_index_pages(
            word_index_path=Path(args.word_index).resolve(),
            pagination_report_path=Path(args.pagination_report).resolve(),
            output_word_index_path=Path(args.output_word_index).resolve(),
            overwrite=bool(args.overwrite),
        )
        return

    parser.error("Unknown command")


if __name__ == "__main__":
    main()
